"""Unit tests for graph_client.GraphClient."""

from __future__ import annotations

import time
from unittest.mock import MagicMock, call, patch

import pytest
import requests.exceptions
import responses as responses_lib
from requests import Session

from graph_auth import GraphAuthClient
from graph_client import GraphClient, GraphClientError


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture()
def mock_auth():
    """Return a GraphAuthClient mock that always returns a test token."""
    auth = MagicMock(spec=GraphAuthClient)
    auth.get_token_for_app.return_value = "test-bearer-token"
    auth.get_token_for_user.return_value = "test-user-token"
    return auth


@pytest.fixture()
def client(mock_auth):
    """Return a GraphClient backed by a mock auth client."""
    return GraphClient(auth_client=mock_auth, use_app_permissions=True)


@pytest.fixture()
def delegated_client(mock_auth):
    """Return a GraphClient using delegated (user) permissions."""
    return GraphClient(auth_client=mock_auth, use_app_permissions=False)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _json_response(data: dict, status: int = 200):
    """Create a responses-library mock JSON response."""
    return responses_lib.Response(
        method=responses_lib.GET,
        url="",  # overridden per test
        json=data,
        status=status,
        content_type="application/json",
    )


# ---------------------------------------------------------------------------
# Initialisation
# ---------------------------------------------------------------------------

class TestGraphClientInit:
    def test_default_base_url(self, mock_auth):
        c = GraphClient(auth_client=mock_auth)
        assert c._base_url == "https://graph.microsoft.com/v1.0"

    def test_trailing_slash_stripped(self, mock_auth):
        c = GraphClient(auth_client=mock_auth, base_url="https://graph.microsoft.com/v1.0/")
        assert c._base_url == "https://graph.microsoft.com/v1.0"

    def test_custom_session(self, mock_auth):
        session = Session()
        c = GraphClient(auth_client=mock_auth, session=session)
        assert c._session is session


# ---------------------------------------------------------------------------
# Token injection
# ---------------------------------------------------------------------------

class TestTokenInjection:
    def test_app_token_used_for_app_permissions(self, client, mock_auth):
        headers = client._build_headers()
        assert headers["Authorization"] == "Bearer test-bearer-token"
        mock_auth.get_token_for_app.assert_called_once()

    def test_user_token_used_for_delegated_permissions(self, delegated_client, mock_auth):
        headers = delegated_client._build_headers()
        assert headers["Authorization"] == "Bearer test-user-token"
        mock_auth.get_token_for_user.assert_called_once()

    def test_extra_headers_merged(self, client):
        headers = client._build_headers({"X-Custom": "value"})
        assert headers["X-Custom"] == "value"
        assert "Authorization" in headers


# ---------------------------------------------------------------------------
# Retry logic
# ---------------------------------------------------------------------------

class TestRetryLogic:
    @responses_lib.activate
    def test_success_on_first_attempt(self, client):
        responses_lib.add(
            responses_lib.GET,
            "https://graph.microsoft.com/v1.0/me",
            json={"id": "user-1"},
            status=200,
        )
        result = client._get("/me")
        assert result["id"] == "user-1"

    @responses_lib.activate
    def test_retries_on_500_then_succeeds(self, client):
        responses_lib.add(
            responses_lib.GET,
            "https://graph.microsoft.com/v1.0/me",
            status=500,
        )
        responses_lib.add(
            responses_lib.GET,
            "https://graph.microsoft.com/v1.0/me",
            json={"id": "user-1"},
            status=200,
        )
        with patch("graph_client.time.sleep"):
            result = client._get("/me")
        assert result["id"] == "user-1"
        assert len(responses_lib.calls) == 2

    @responses_lib.activate
    def test_raises_after_max_retries(self, client):
        for _ in range(6):  # max_retries=5 → 6 total attempts
            responses_lib.add(
                responses_lib.GET,
                "https://graph.microsoft.com/v1.0/me",
                status=500,
            )
        with patch("graph_client.time.sleep"):
            with pytest.raises(GraphClientError, match="500"):
                client._get("/me")

    @responses_lib.activate
    def test_retries_on_rate_limit_429(self, client):
        responses_lib.add(
            responses_lib.GET,
            "https://graph.microsoft.com/v1.0/me",
            status=429,
            headers={"Retry-After": "2"},
        )
        responses_lib.add(
            responses_lib.GET,
            "https://graph.microsoft.com/v1.0/me",
            json={"id": "user-1"},
            status=200,
        )
        with patch("graph_client.time.sleep") as mock_sleep:
            result = client._get("/me")
        assert result["id"] == "user-1"
        # Should have slept for 2 seconds as specified in Retry-After
        mock_sleep.assert_called_once_with(2.0)

    @responses_lib.activate
    def test_parse_retry_after_header(self, client):
        responses_lib.add(
            responses_lib.GET,
            "https://graph.microsoft.com/v1.0/me",
            status=429,
            headers={"Retry-After": "10"},
        )
        responses_lib.add(
            responses_lib.GET,
            "https://graph.microsoft.com/v1.0/me",
            json={"id": "u"},
            status=200,
        )
        with patch("graph_client.time.sleep") as mock_sleep:
            client._get("/me")
        mock_sleep.assert_called_once_with(10.0)

    @responses_lib.activate
    def test_non_retryable_error_raises_immediately(self, client):
        responses_lib.add(
            responses_lib.GET,
            "https://graph.microsoft.com/v1.0/me",
            status=403,
        )
        with pytest.raises(GraphClientError, match="403"):
            client._get("/me")
        assert len(responses_lib.calls) == 1  # no retries

    @responses_lib.activate
    def test_raises_on_request_exception(self, client):
        responses_lib.add(
            responses_lib.GET,
            "https://graph.microsoft.com/v1.0/me",
            body=requests.exceptions.ConnectionError("Network error"),
        )
        # responses only raises on first call; add extras for retries
        for _ in range(5):
            responses_lib.add(
                responses_lib.GET,
                "https://graph.microsoft.com/v1.0/me",
                body=requests.exceptions.ConnectionError("Network error"),
            )
        with patch("graph_client.time.sleep"):
            with pytest.raises(GraphClientError, match="Network error"):
                client._get("/me")

    def test_status_code_on_error(self, client):
        with patch.object(client._session, "request") as mock_req:
            mock_resp = MagicMock()
            mock_resp.status_code = 401
            mock_resp.ok = False
            mock_resp.text = "Unauthorized"
            mock_resp.headers = {}
            mock_req.return_value = mock_resp
            with pytest.raises(GraphClientError) as exc_info:
                client._get("/me")
        assert exc_info.value.status_code == 401


# ---------------------------------------------------------------------------
# Pagination
# ---------------------------------------------------------------------------

class TestPagination:
    @responses_lib.activate
    def test_follows_next_link(self, client):
        page1_url = "https://graph.microsoft.com/v1.0/me/events"
        page2_url = "https://graph.microsoft.com/v1.0/me/events?$skiptoken=page2"
        responses_lib.add(
            responses_lib.GET,
            page1_url,
            json={
                "value": [{"id": "event-1"}, {"id": "event-2"}],
                "@odata.nextLink": page2_url,
            },
            status=200,
        )
        responses_lib.add(
            responses_lib.GET,
            page2_url,
            json={"value": [{"id": "event-3"}]},
            status=200,
        )
        items = client._get_all_pages("/me/events")
        assert len(items) == 3
        assert items[0]["id"] == "event-1"
        assert items[2]["id"] == "event-3"

    @responses_lib.activate
    def test_single_page_no_next_link(self, client):
        responses_lib.add(
            responses_lib.GET,
            "https://graph.microsoft.com/v1.0/me/messages",
            json={"value": [{"id": "msg-1"}]},
            status=200,
        )
        items = client._get_all_pages("/me/messages")
        assert len(items) == 1


# ---------------------------------------------------------------------------
# Endpoint wrappers
# ---------------------------------------------------------------------------

class TestEndpointWrappers:
    @responses_lib.activate
    def test_get_me(self, client):
        responses_lib.add(
            responses_lib.GET,
            "https://graph.microsoft.com/v1.0/me",
            json={"displayName": "Test User"},
            status=200,
        )
        result = client.get_me()
        assert result["displayName"] == "Test User"

    @responses_lib.activate
    def test_list_calendar_events_default_user(self, client):
        responses_lib.add(
            responses_lib.GET,
            "https://graph.microsoft.com/v1.0/me/events",
            json={"value": [{"id": "evt-1", "subject": "Stand-up"}]},
            status=200,
        )
        events = client.list_calendar_events()
        assert len(events) == 1
        assert events[0]["subject"] == "Stand-up"

    @responses_lib.activate
    def test_list_calendar_events_specific_user(self, client):
        uid = "user@example.com"
        responses_lib.add(
            responses_lib.GET,
            f"https://graph.microsoft.com/v1.0/users/{uid}/events",
            json={"value": [{"id": "evt-2"}]},
            status=200,
        )
        events = client.list_calendar_events(user_id=uid)
        assert events[0]["id"] == "evt-2"

    @responses_lib.activate
    def test_list_chats(self, client):
        responses_lib.add(
            responses_lib.GET,
            "https://graph.microsoft.com/v1.0/me/chats",
            json={"value": [{"id": "chat-1"}]},
            status=200,
        )
        chats = client.list_chats()
        assert chats[0]["id"] == "chat-1"

    @responses_lib.activate
    def test_list_messages(self, client):
        responses_lib.add(
            responses_lib.GET,
            "https://graph.microsoft.com/v1.0/me/messages",
            json={"value": [{"id": "msg-1", "subject": "Hello"}]},
            status=200,
        )
        msgs = client.list_messages()
        assert msgs[0]["subject"] == "Hello"

    @responses_lib.activate
    def test_list_drive_items(self, client):
        responses_lib.add(
            responses_lib.GET,
            "https://graph.microsoft.com/v1.0/me/drive/root/children",
            json={"value": [{"id": "file-1", "name": "report.pdf"}]},
            status=200,
        )
        items = client.list_drive_items()
        assert items[0]["name"] == "report.pdf"

    @responses_lib.activate
    def test_get_online_meeting(self, client):
        uid = "user@example.com"
        mid = "meeting-id-123"
        responses_lib.add(
            responses_lib.GET,
            f"https://graph.microsoft.com/v1.0/users/{uid}/onlineMeetings/{mid}",
            json={"id": mid, "subject": "Weekly Sync"},
            status=200,
        )
        meeting = client.get_online_meeting(user_id=uid, meeting_id=mid)
        assert meeting["subject"] == "Weekly Sync"
