"""
Microsoft Graph API client wrapper.

This module provides a thin wrapper around the Microsoft Graph REST API that
handles authentication (via MSAL), token caching/refresh, and request-level
concerns such as rate-limit back-off and consistent error mapping.

Environment variables
---------------------
``GRAPH_TENANT_ID``
    Azure AD tenant ID.
``GRAPH_CLIENT_ID``
    Azure AD application (client) ID.
``GRAPH_CLIENT_SECRET``
    Application secret.  Required for app-only flows; omit for delegated
    flows where ``GRAPH_ACCESS_TOKEN`` is supplied directly.
``GRAPH_ACCESS_TOKEN``
    A pre-obtained delegated access token.  When set the client skips MSAL
    and uses this token directly (useful for testing/development).
"""

from __future__ import annotations

import logging
import os
import time
from typing import Any, Dict, Optional

import requests

logger = logging.getLogger(__name__)

_GRAPH_BASE = "https://graph.microsoft.com/v1.0"
_BETA_BASE = "https://graph.microsoft.com/beta"

# Retry parameters for 429 / 503 responses
_MAX_RETRIES = 3
_BASE_BACKOFF = 1.0  # seconds


class GraphAPIError(Exception):
    """Raised when the Graph API returns an error response."""

    def __init__(self, status_code: int, message: str, error_code: Optional[str] = None) -> None:
        super().__init__(message)
        self.status_code = status_code
        self.message = message
        self.error_code = error_code

    def __repr__(self) -> str:
        return f"GraphAPIError(status={self.status_code}, code={self.error_code!r}, message={self.message!r})"


class GraphPermissionError(GraphAPIError):
    """Raised for 403 Forbidden responses (insufficient permissions)."""


class GraphNotFoundError(GraphAPIError):
    """Raised for 404 Not Found responses."""


def _raise_for_response(response: requests.Response) -> None:
    """Inspect *response* and raise the appropriate :class:`GraphAPIError` subclass."""
    if response.ok:
        return

    error_code: Optional[str] = None
    message = response.reason or "Unknown error"

    try:
        body = response.json()
        error = body.get("error", {})
        error_code = error.get("code")
        message = error.get("message", message)
    except ValueError:
        pass

    if response.status_code == 403:
        raise GraphPermissionError(response.status_code, message, error_code)
    if response.status_code == 404:
        raise GraphNotFoundError(response.status_code, message, error_code)
    raise GraphAPIError(response.status_code, message, error_code)


class GraphClient:
    """
    Reusable client for Microsoft Graph API.

    Authentication supports two modes:

    1. **App-only** (daemon): Pass ``tenant_id``, ``client_id``, and
       ``client_secret``.  A client-credentials token is obtained via MSAL.
    2. **Pre-obtained token**: Pass ``access_token`` directly.  Useful when
       a delegated token has already been acquired upstream.

    Parameters default to the corresponding environment variables when not
    supplied explicitly.
    """

    def __init__(
        self,
        tenant_id: Optional[str] = None,
        client_id: Optional[str] = None,
        client_secret: Optional[str] = None,
        access_token: Optional[str] = None,
        use_beta: bool = False,
    ) -> None:
        self._tenant_id = tenant_id or os.environ.get("GRAPH_TENANT_ID", "")
        self._client_id = client_id or os.environ.get("GRAPH_CLIENT_ID", "")
        self._client_secret = client_secret or os.environ.get("GRAPH_CLIENT_SECRET", "")
        self._static_token = access_token or os.environ.get("GRAPH_ACCESS_TOKEN", "")
        self._base_url = _BETA_BASE if use_beta else _GRAPH_BASE
        self._session = requests.Session()
        self._msal_app = None  # lazily initialised

    # ------------------------------------------------------------------
    # Token management
    # ------------------------------------------------------------------

    def _get_msal_app(self):
        """Return (or create) an MSAL ConfidentialClientApplication."""
        if self._msal_app is None:
            try:
                import msal  # type: ignore
            except ImportError as exc:
                raise ImportError(
                    "msal is required for app-only auth. Install it with: pip install msal"
                ) from exc

            authority = f"https://login.microsoftonline.com/{self._tenant_id}"
            self._msal_app = msal.ConfidentialClientApplication(
                self._client_id,
                authority=authority,
                client_credential=self._client_secret,
            )
        return self._msal_app

    def _acquire_token(self) -> str:
        """Return a valid access token, refreshing via MSAL when necessary."""
        if self._static_token:
            return self._static_token

        scopes = ["https://graph.microsoft.com/.default"]
        app = self._get_msal_app()

        # Try silent (cache) first
        result = app.acquire_token_silent(scopes, account=None)
        if not result:
            result = app.acquire_token_for_client(scopes=scopes)

        if "access_token" not in result:
            error = result.get("error_description", result.get("error", "Unknown MSAL error"))
            raise GraphAPIError(401, f"Token acquisition failed: {error}")

        return result["access_token"]

    # ------------------------------------------------------------------
    # HTTP helpers
    # ------------------------------------------------------------------

    def _headers(self) -> Dict[str, str]:
        return {
            "Authorization": f"Bearer {self._acquire_token()}",
            "Accept": "application/json",
            "Content-Type": "application/json",
        }

    def _request(
        self,
        method: str,
        url: str,
        *,
        accept_content: Optional[str] = None,
        **kwargs: Any,
    ) -> requests.Response:
        """Execute an HTTP *method* against *url* with automatic retry on 429/503."""
        headers = self._headers()
        if accept_content:
            headers["Accept"] = accept_content

        for attempt in range(_MAX_RETRIES + 1):
            response = self._session.request(method, url, headers=headers, **kwargs)

            if response.status_code in (429, 503):
                retry_after = int(response.headers.get("Retry-After", _BASE_BACKOFF * (2 ** attempt)))
                logger.warning(
                    "Rate limited (%d). Retrying after %s seconds (attempt %d/%d).",
                    response.status_code,
                    retry_after,
                    attempt + 1,
                    _MAX_RETRIES,
                )
                time.sleep(retry_after)
                continue

            _raise_for_response(response)
            return response

        # Final attempt after all retries
        _raise_for_response(response)  # type: ignore[possibly-undefined]
        return response  # type: ignore[possibly-undefined]

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def get(self, path: str, **kwargs: Any) -> Dict[str, Any]:
        """GET *path* and return the parsed JSON body."""
        url = path if path.startswith("https://") else f"{self._base_url}{path}"
        response = self._request("GET", url, **kwargs)
        return response.json()

    def get_raw(self, path: str, accept: str = "text/plain") -> str:
        """GET *path* and return the raw response text (e.g. for transcript content)."""
        url = path if path.startswith("https://") else f"{self._base_url}{path}"
        response = self._request("GET", url, accept_content=accept)
        return response.text

    def get_paginated(self, path: str, **kwargs: Any):
        """
        Yield all items from a paginated Graph response.

        Follows ``@odata.nextLink`` automatically.
        """
        url: Optional[str] = path if path.startswith("https://") else f"{self._base_url}{path}"
        while url:
            data = self.get(url, **kwargs)
            yield from data.get("value", [])
            url = data.get("@odata.nextLink")
