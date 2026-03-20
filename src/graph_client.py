"""Reusable Microsoft Graph API HTTP client.

Features
--------
* Automatic bearer-token injection via :class:`~graph_auth.GraphAuthClient`
* Exponential back-off retry for transient errors (5xx) and rate limiting (429)
* Configurable base URL so the module works with GCC-High / DoD clouds
* Thin wrappers for the most common Graph endpoints used by this application
"""

from __future__ import annotations

import logging
import os
import time
from typing import Any, Optional

import requests
from requests import Response, Session

from graph_auth import GraphAuthClient

logger = logging.getLogger(__name__)

_DEFAULT_BASE_URL: str = os.environ.get(
    "GRAPH_API_BASE_URL", "https://graph.microsoft.com/v1.0"
)

# Retry configuration
_MAX_RETRIES: int = 5
_INITIAL_BACKOFF_SECONDS: float = 1.0
_BACKOFF_MULTIPLIER: float = 2.0
_MAX_BACKOFF_SECONDS: float = 32.0

# HTTP status codes that warrant a retry
_RETRYABLE_STATUS_CODES: frozenset[int] = frozenset({429, 500, 502, 503, 504})


class GraphClientError(Exception):
    """Raised when a Graph API request fails after all retries."""

    def __init__(self, message: str, status_code: Optional[int] = None) -> None:
        super().__init__(message)
        self.status_code = status_code


class GraphClient:
    """HTTP client for the Microsoft Graph API.

    Parameters
    ----------
    auth_client:
        A configured :class:`~graph_auth.GraphAuthClient` instance.  When
        *None* a new one is created using environment variables.
    base_url:
        Graph API base URL.  Defaults to ``GRAPH_API_BASE_URL`` env var or
        ``https://graph.microsoft.com/v1.0``.
    use_app_permissions:
        When *True* (default) the client uses the application-permission
        (client credentials) flow.  Set to *False* to use delegated tokens.
    max_retries:
        Maximum number of retry attempts for retryable errors.
    session:
        A custom :class:`requests.Session` to use.  Mainly useful for testing.
    """

    def __init__(
        self,
        *,
        auth_client: Optional[GraphAuthClient] = None,
        base_url: str = _DEFAULT_BASE_URL,
        use_app_permissions: bool = True,
        max_retries: int = _MAX_RETRIES,
        session: Optional[Session] = None,
    ) -> None:
        self._auth = auth_client or GraphAuthClient()
        self._base_url = base_url.rstrip("/")
        self._use_app_permissions = use_app_permissions
        self._max_retries = max_retries
        self._session: Session = session or requests.Session()

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _get_access_token(self) -> str:
        """Return a valid access token for the configured flow."""
        if self._use_app_permissions:
            return self._auth.get_token_for_app()
        return self._auth.get_token_for_user()

    def _build_headers(self, extra: Optional[dict[str, str]] = None) -> dict[str, str]:
        """Build request headers including the Authorization bearer token."""
        token = self._get_access_token()
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }
        if extra:
            headers.update(extra)
        return headers

    @staticmethod
    def _parse_retry_after(response: Response) -> float:
        """Return the number of seconds to wait from a *Retry-After* header."""
        retry_after = response.headers.get("Retry-After")
        if retry_after is not None:
            try:
                return float(retry_after)
            except ValueError:
                pass
        return _INITIAL_BACKOFF_SECONDS

    def _request_with_retry(
        self,
        method: str,
        url: str,
        **kwargs: Any,
    ) -> Response:
        """Execute *method* against *url* with exponential back-off retries.

        Parameters
        ----------
        method:
            HTTP verb (``"GET"``, ``"POST"``, etc.).
        url:
            Fully-qualified URL.
        **kwargs:
            Passed directly to :meth:`requests.Session.request`.

        Returns
        -------
        Response
            The successful HTTP response.

        Raises
        ------
        GraphClientError
            When all retries are exhausted or a non-retryable error occurs.
        """
        backoff = _INITIAL_BACKOFF_SECONDS
        last_response: Optional[Response] = None

        for attempt in range(self._max_retries + 1):
            try:
                response = self._session.request(method, url, **kwargs)
            except requests.exceptions.RequestException as exc:
                if attempt == self._max_retries:
                    raise GraphClientError(
                        f"Request to {url} failed after {self._max_retries} retries: {exc}"
                    ) from exc
                logger.warning(
                    "Request error on attempt %d/%d: %s – retrying in %.1fs",
                    attempt + 1,
                    self._max_retries + 1,
                    exc,
                    backoff,
                )
                time.sleep(backoff)
                backoff = min(backoff * _BACKOFF_MULTIPLIER, _MAX_BACKOFF_SECONDS)
                continue

            last_response = response

            if response.status_code not in _RETRYABLE_STATUS_CODES:
                break

            if attempt == self._max_retries:
                break

            wait: float
            if response.status_code == 429:
                wait = self._parse_retry_after(response)
                logger.warning(
                    "Rate limited (429) on attempt %d/%d – retrying in %.1fs",
                    attempt + 1,
                    self._max_retries + 1,
                    wait,
                )
            else:
                wait = backoff
                logger.warning(
                    "Transient error %d on attempt %d/%d – retrying in %.1fs",
                    response.status_code,
                    attempt + 1,
                    self._max_retries + 1,
                    wait,
                )
            time.sleep(wait)
            backoff = min(backoff * _BACKOFF_MULTIPLIER, _MAX_BACKOFF_SECONDS)

        if last_response is None:
            raise GraphClientError(f"No response received from {url}")

        if not last_response.ok:
            raise GraphClientError(
                f"Graph API request failed [{last_response.status_code}]: "
                f"{last_response.text}",
                status_code=last_response.status_code,
            )

        return last_response

    def _get(self, path: str, params: Optional[dict] = None) -> dict:
        """Perform an authenticated GET request and return the JSON body."""
        url = f"{self._base_url}/{path.lstrip('/')}"
        response = self._request_with_retry(
            "GET",
            url,
            headers=self._build_headers(),
            params=params,
        )
        return response.json()

    def _post(self, path: str, json: Optional[dict] = None) -> dict:
        """Perform an authenticated POST request and return the JSON body."""
        url = f"{self._base_url}/{path.lstrip('/')}"
        response = self._request_with_retry(
            "POST",
            url,
            headers=self._build_headers(),
            json=json,
        )
        return response.json()

    # ------------------------------------------------------------------
    # Graph endpoint wrappers
    # ------------------------------------------------------------------

    def get_me(self) -> dict:
        """Return the signed-in user's profile (requires delegated token).

        Endpoint: ``GET /me``
        """
        return self._get("/me")

    def list_calendar_events(
        self,
        user_id: str = "me",
        top: int = 50,
        filter_query: Optional[str] = None,
    ) -> list[dict]:
        """Return calendar events for *user_id*.

        Endpoint: ``GET /users/{user_id}/events``

        Parameters
        ----------
        user_id:
            UPN or object ID of the user.  Use ``"me"`` for the signed-in user.
        top:
            Maximum number of events to return per page (max 999).
        filter_query:
            OData ``$filter`` expression.

        Returns
        -------
        list[dict]
            All events across all pages.
        """
        params: dict[str, Any] = {"$top": top}
        if filter_query:
            params["$filter"] = filter_query

        path = f"/users/{user_id}/events" if user_id != "me" else "/me/events"
        return self._get_all_pages(path, params=params)

    def list_chats(self, user_id: str = "me") -> list[dict]:
        """Return chats the user is a member of.

        Endpoint: ``GET /users/{user_id}/chats``
        """
        path = f"/users/{user_id}/chats" if user_id != "me" else "/me/chats"
        return self._get_all_pages(path)

    def list_messages(
        self,
        user_id: str = "me",
        top: int = 50,
        filter_query: Optional[str] = None,
    ) -> list[dict]:
        """Return mail messages for *user_id*.

        Endpoint: ``GET /users/{user_id}/messages``
        """
        params: dict[str, Any] = {"$top": top}
        if filter_query:
            params["$filter"] = filter_query

        path = f"/users/{user_id}/messages" if user_id != "me" else "/me/messages"
        return self._get_all_pages(path, params=params)

    def list_drive_items(
        self,
        user_id: str = "me",
        item_path: str = "root",
    ) -> list[dict]:
        """Return drive items (files) for *user_id*.

        Endpoint: ``GET /users/{user_id}/drive/root/children``
        """
        if user_id == "me":
            path = f"/me/drive/{item_path}/children"
        else:
            path = f"/users/{user_id}/drive/{item_path}/children"
        return self._get_all_pages(path)

    def get_online_meeting(self, user_id: str, meeting_id: str) -> dict:
        """Return details of an online meeting.

        Endpoint: ``GET /users/{user_id}/onlineMeetings/{meeting_id}``
        """
        return self._get(f"/users/{user_id}/onlineMeetings/{meeting_id}")

    # ------------------------------------------------------------------
    # Pagination helper
    # ------------------------------------------------------------------

    def _get_all_pages(
        self,
        path: str,
        params: Optional[dict] = None,
    ) -> list[dict]:
        """Iterate through OData ``@odata.nextLink`` pages and collect all items.

        Parameters
        ----------
        path:
            API path relative to *base_url*.
        params:
            Optional query parameters for the first request.

        Returns
        -------
        list[dict]
            Combined list of ``value`` items from all pages.
        """
        results: list[dict] = []
        next_url: Optional[str] = f"{self._base_url}/{path.lstrip('/')}"
        current_params: Optional[dict] = params

        while next_url:
            url = next_url
            response = self._request_with_retry(
                "GET",
                url,
                headers=self._build_headers(),
                params=current_params,
            )
            data = response.json()
            results.extend(data.get("value", []))
            next_url = data.get("@odata.nextLink")
            # Params are embedded in nextLink; don't re-send them
            current_params = None

        return results
