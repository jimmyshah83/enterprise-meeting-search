"""Microsoft Graph API authentication module.

Supports both:
- Application (daemon) permission flow via client credentials grant
- Delegated (user) permission flow via authorization code grant

Credentials are read exclusively from environment variables so that no
secrets are ever hard-coded in source code.
"""

from __future__ import annotations

import logging
import os
import time
from typing import Optional

import msal

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Default scopes required by this application
# ---------------------------------------------------------------------------
DEFAULT_SCOPES: list[str] = [
    "https://graph.microsoft.com/Calendars.Read",
    "https://graph.microsoft.com/Chat.Read",
    "https://graph.microsoft.com/Mail.Read",
    "https://graph.microsoft.com/Files.Read",
    "https://graph.microsoft.com/OnlineMeetings.Read",
]

# Scope used for the client-credentials (application) flow.  The resource
# URI /.default causes Azure AD to grant all application permissions that
# have been statically consented to in the app registration.
APP_DEFAULT_SCOPE: list[str] = ["https://graph.microsoft.com/.default"]

# How many seconds before the recorded expiry time we treat a token as
# already expired and proactively refresh it.
_TOKEN_EXPIRY_BUFFER_SECONDS: int = 300  # 5 minutes


class AuthenticationError(Exception):
    """Raised when token acquisition fails."""


class GraphAuthClient:
    """Handles OAuth2 token acquisition and refresh for Microsoft Graph API.

    Parameters
    ----------
    tenant_id:
        Azure AD tenant ID.  Falls back to the ``AZURE_TENANT_ID`` environment
        variable when *None*.
    client_id:
        Azure AD application (client) ID.  Falls back to ``AZURE_CLIENT_ID``.
    client_secret:
        Client secret for confidential clients.  Falls back to
        ``AZURE_CLIENT_SECRET``.  Required for the application flow and the
        authorization-code delegated flow when no certificate is provided.
    redirect_uri:
        Redirect URI registered in the app registration.  Falls back to
        ``AZURE_REDIRECT_URI``.  Required for the delegated flow.
    authority:
        Full authority URL (e.g. ``https://login.microsoftonline.com/<tenant>``).
        Constructed automatically when *None*.
    scopes:
        List of OAuth2 scopes to request.  Defaults to :data:`DEFAULT_SCOPES`.
    """

    def __init__(
        self,
        *,
        tenant_id: Optional[str] = None,
        client_id: Optional[str] = None,
        client_secret: Optional[str] = None,
        redirect_uri: Optional[str] = None,
        authority: Optional[str] = None,
        scopes: Optional[list[str]] = None,
    ) -> None:
        self._tenant_id: str = tenant_id or os.environ["AZURE_TENANT_ID"]
        self._client_id: str = client_id or os.environ["AZURE_CLIENT_ID"]
        self._client_secret: Optional[str] = client_secret or os.environ.get(
            "AZURE_CLIENT_SECRET"
        )
        self._redirect_uri: str = redirect_uri or os.environ.get(
            "AZURE_REDIRECT_URI", "http://localhost:8080"
        )
        self._authority: str = authority or (
            f"https://login.microsoftonline.com/{self._tenant_id}"
        )
        self._scopes: list[str] = scopes or DEFAULT_SCOPES

        # Cached token details
        self._access_token: Optional[str] = None
        self._token_expires_at: float = 0.0

        # MSAL application instances (lazily created)
        self._confidential_app: Optional[msal.ConfidentialClientApplication] = None
        self._public_app: Optional[msal.PublicClientApplication] = None

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _get_confidential_app(self) -> msal.ConfidentialClientApplication:
        """Return (or lazily create) the MSAL confidential client app."""
        if self._confidential_app is None:
            if not self._client_secret:
                raise AuthenticationError(
                    "A client secret is required for the confidential client flow. "
                    "Set AZURE_CLIENT_SECRET."
                )
            self._confidential_app = msal.ConfidentialClientApplication(
                client_id=self._client_id,
                client_credential=self._client_secret,
                authority=self._authority,
            )
        return self._confidential_app

    def _get_public_app(self) -> msal.PublicClientApplication:
        """Return (or lazily create) the MSAL public client app."""
        if self._public_app is None:
            self._public_app = msal.PublicClientApplication(
                client_id=self._client_id,
                authority=self._authority,
            )
        return self._public_app

    def _cache_token(self, result: dict) -> str:
        """Validate an MSAL token result and cache the access token.

        Parameters
        ----------
        result:
            The dict returned by an MSAL token-acquisition call.

        Returns
        -------
        str
            The access token string.

        Raises
        ------
        AuthenticationError
            When *result* does not contain an access token.
        """
        if "access_token" not in result:
            error = result.get("error", "unknown_error")
            description = result.get("error_description", "No description provided.")
            raise AuthenticationError(
                f"Token acquisition failed [{error}]: {description}"
            )

        self._access_token = result["access_token"]
        expires_in: int = result.get("expires_in", 3600)
        self._token_expires_at = time.monotonic() + expires_in
        logger.debug("Access token acquired; expires in %d seconds.", expires_in)
        return self._access_token

    def _is_token_valid(self) -> bool:
        """Return *True* when the cached token is present and not near expiry."""
        if not self._access_token:
            return False
        remaining = self._token_expires_at - time.monotonic()
        return remaining > _TOKEN_EXPIRY_BUFFER_SECONDS

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def get_token_for_app(self) -> str:
        """Acquire (or return cached) access token using the client credentials grant.

        This is the *application* (daemon) permission flow.  The token is
        scoped to ``https://graph.microsoft.com/.default``.

        Returns
        -------
        str
            A valid access token.

        Raises
        ------
        AuthenticationError
            When token acquisition fails.
        """
        if self._is_token_valid():
            return self._access_token  # type: ignore[return-value]

        logger.debug("Acquiring application token via client credentials grant.")
        app = self._get_confidential_app()

        # Try the token cache first (MSAL manages its own in-memory cache)
        result = app.acquire_token_silent(APP_DEFAULT_SCOPE, account=None)
        if not result:
            result = app.acquire_token_for_client(scopes=APP_DEFAULT_SCOPE)

        return self._cache_token(result)

    def get_token_for_user(self, username: Optional[str] = None) -> str:
        """Acquire (or return cached) access token using the resource owner password grant.

        .. warning::
            The ROPC flow is deprecated by Microsoft for new applications.  It
            is provided here for completeness and backwards compatibility.  For
            interactive sign-in prefer :meth:`get_token_interactive`.

        Parameters
        ----------
        username:
            The user's UPN (email address).  When *None*, falls back to the
            ``AZURE_USERNAME`` environment variable.

        Returns
        -------
        str
            A valid access token.
        """
        if self._is_token_valid():
            return self._access_token  # type: ignore[return-value]

        username = username or os.environ.get("AZURE_USERNAME", "")
        password = os.environ.get("AZURE_PASSWORD", "")

        logger.debug("Acquiring delegated token via ROPC grant for %s.", username)
        app = self._get_public_app()

        accounts = app.get_accounts(username=username)
        if accounts:
            result = app.acquire_token_silent(self._scopes, account=accounts[0])
            if result:
                return self._cache_token(result)

        result = app.acquire_token_by_username_password(
            username=username,
            password=password,
            scopes=self._scopes,
        )
        return self._cache_token(result)

    def get_token_interactive(self) -> str:
        """Acquire a delegated token interactively (device code flow).

        Prints a device-code prompt to stdout so the user can sign in via a
        browser on any device.

        Returns
        -------
        str
            A valid access token.
        """
        if self._is_token_valid():
            return self._access_token  # type: ignore[return-value]

        logger.debug("Acquiring delegated token via device code flow.")
        app = self._get_public_app()

        flow = app.initiate_device_flow(scopes=self._scopes)
        if "user_code" not in flow:
            raise AuthenticationError(
                f"Failed to create device flow: {flow.get('error_description', '')}"
            )

        print(flow["message"])  # noqa: T201 – intentional user-facing output
        result = app.acquire_token_by_device_flow(flow)
        return self._cache_token(result)

    def get_authorization_url(self) -> tuple[str, str]:
        """Return the authorization URL for the authorization code flow.

        Returns
        -------
        tuple[str, str]
            ``(authorization_url, state)`` where *state* should be stored and
            validated in :meth:`exchange_auth_code`.
        """
        app = self._get_confidential_app()
        result = app.initiate_auth_code_flow(
            scopes=self._scopes,
            redirect_uri=self._redirect_uri,
        )
        auth_url: str = result["auth_uri"]
        state: str = result["state"]
        # Store the flow dict on the instance so exchange_auth_code can use it
        self._auth_code_flow = result
        return auth_url, state

    def exchange_auth_code(self, auth_response: dict) -> str:
        """Exchange an authorization code for an access token.

        Parameters
        ----------
        auth_response:
            The query-parameter dict from the redirect URI (must contain
            ``code`` and ``state``).

        Returns
        -------
        str
            A valid access token.

        Raises
        ------
        AuthenticationError
            When the token exchange fails.
        """
        if not hasattr(self, "_auth_code_flow"):
            raise AuthenticationError(
                "Call get_authorization_url() before exchange_auth_code()."
            )
        app = self._get_confidential_app()
        result = app.acquire_token_by_auth_code_flow(
            self._auth_code_flow,
            auth_response,
        )
        return self._cache_token(result)

    def refresh_token(self) -> str:
        """Force a token refresh regardless of current expiry.

        For the application flow this re-runs the client credentials grant.
        Delegated tokens are refreshed through MSAL's silent-acquire mechanism.

        Returns
        -------
        str
            A fresh access token.
        """
        # Invalidate the cached token so _is_token_valid returns False
        self._access_token = None
        self._token_expires_at = 0.0
        return self.get_token_for_app()
