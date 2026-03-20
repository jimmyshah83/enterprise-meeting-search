"""Unit tests for graph_auth.GraphAuthClient."""

from __future__ import annotations

import time
from unittest.mock import MagicMock, patch

import pytest

from graph_auth import AuthenticationError, GraphAuthClient


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture()
def env_vars(monkeypatch):
    """Set required environment variables for every test."""
    monkeypatch.setenv("AZURE_TENANT_ID", "test-tenant-id")
    monkeypatch.setenv("AZURE_CLIENT_ID", "test-client-id")
    monkeypatch.setenv("AZURE_CLIENT_SECRET", "test-client-secret")
    monkeypatch.setenv("AZURE_REDIRECT_URI", "http://localhost:8080")


@pytest.fixture()
def auth_client(env_vars):
    """Return a GraphAuthClient with environment variables set."""
    return GraphAuthClient()


# ---------------------------------------------------------------------------
# Initialisation tests
# ---------------------------------------------------------------------------

class TestGraphAuthClientInit:
    def test_reads_env_vars(self, env_vars):
        client = GraphAuthClient()
        assert client._tenant_id == "test-tenant-id"
        assert client._client_id == "test-client-id"
        assert client._client_secret == "test-client-secret"

    def test_explicit_params_override_env(self, env_vars):
        client = GraphAuthClient(
            tenant_id="explicit-tenant",
            client_id="explicit-client",
            client_secret="explicit-secret",
        )
        assert client._tenant_id == "explicit-tenant"
        assert client._client_id == "explicit-client"
        assert client._client_secret == "explicit-secret"

    def test_authority_constructed_from_tenant(self, env_vars):
        client = GraphAuthClient()
        assert client._authority == (
            "https://login.microsoftonline.com/test-tenant-id"
        )

    def test_custom_authority(self, env_vars):
        client = GraphAuthClient(authority="https://custom.authority/tenant")
        assert client._authority == "https://custom.authority/tenant"

    def test_missing_tenant_id_raises(self, monkeypatch):
        monkeypatch.delenv("AZURE_TENANT_ID", raising=False)
        with pytest.raises(KeyError):
            GraphAuthClient()

    def test_missing_client_id_raises(self, monkeypatch):
        monkeypatch.setenv("AZURE_TENANT_ID", "tid")
        monkeypatch.delenv("AZURE_CLIENT_ID", raising=False)
        with pytest.raises(KeyError):
            GraphAuthClient()


# ---------------------------------------------------------------------------
# Token caching / validity helpers
# ---------------------------------------------------------------------------

class TestTokenValidity:
    def test_no_token_is_invalid(self, auth_client):
        assert not auth_client._is_token_valid()

    def test_fresh_token_is_valid(self, auth_client):
        auth_client._access_token = "my-token"
        auth_client._token_expires_at = time.monotonic() + 3600
        assert auth_client._is_token_valid()

    def test_expired_token_is_invalid(self, auth_client):
        auth_client._access_token = "my-token"
        # Expired 10 seconds ago
        auth_client._token_expires_at = time.monotonic() - 10
        assert not auth_client._is_token_valid()

    def test_token_near_expiry_is_invalid(self, auth_client):
        auth_client._access_token = "my-token"
        # Within the 5-minute buffer
        auth_client._token_expires_at = time.monotonic() + 60
        assert not auth_client._is_token_valid()


class TestCacheToken:
    def test_caches_access_token(self, auth_client):
        result = {"access_token": "tok123", "expires_in": 3600}
        token = auth_client._cache_token(result)
        assert token == "tok123"
        assert auth_client._access_token == "tok123"

    def test_raises_on_missing_access_token(self, auth_client):
        result = {"error": "invalid_client", "error_description": "bad secret"}
        with pytest.raises(AuthenticationError, match="invalid_client"):
            auth_client._cache_token(result)

    def test_uses_default_expiry_when_missing(self, auth_client):
        result = {"access_token": "tok"}
        auth_client._cache_token(result)
        # Should be roughly now + 3600
        assert auth_client._token_expires_at > time.monotonic() + 3000


# ---------------------------------------------------------------------------
# Application (client credentials) flow
# ---------------------------------------------------------------------------

class TestGetTokenForApp:
    def test_returns_cached_token(self, auth_client):
        auth_client._access_token = "cached-tok"
        auth_client._token_expires_at = time.monotonic() + 3600
        assert auth_client.get_token_for_app() == "cached-tok"

    def test_acquires_token_when_cache_empty(self, auth_client):
        mock_app = MagicMock()
        mock_app.acquire_token_silent.return_value = None
        mock_app.acquire_token_for_client.return_value = {
            "access_token": "new-app-token",
            "expires_in": 3600,
        }
        auth_client._confidential_app = mock_app
        token = auth_client.get_token_for_app()
        assert token == "new-app-token"
        mock_app.acquire_token_for_client.assert_called_once()

    def test_uses_msal_cache_when_available(self, auth_client):
        mock_app = MagicMock()
        mock_app.acquire_token_silent.return_value = {
            "access_token": "silent-tok",
            "expires_in": 3600,
        }
        auth_client._confidential_app = mock_app
        token = auth_client.get_token_for_app()
        assert token == "silent-tok"
        mock_app.acquire_token_for_client.assert_not_called()

    def test_raises_when_no_client_secret(self, env_vars, monkeypatch):
        monkeypatch.delenv("AZURE_CLIENT_SECRET", raising=False)
        client = GraphAuthClient(client_secret=None)
        with pytest.raises(AuthenticationError, match="client secret"):
            client.get_token_for_app()

    def test_raises_on_msal_error(self, auth_client):
        mock_app = MagicMock()
        mock_app.acquire_token_silent.return_value = None
        mock_app.acquire_token_for_client.return_value = {
            "error": "invalid_scope",
            "error_description": "Scope not found",
        }
        auth_client._confidential_app = mock_app
        with pytest.raises(AuthenticationError, match="invalid_scope"):
            auth_client.get_token_for_app()


# ---------------------------------------------------------------------------
# Delegated (user) flow
# ---------------------------------------------------------------------------

class TestGetTokenForUser:
    def test_returns_cached_token(self, auth_client):
        auth_client._access_token = "user-cached"
        auth_client._token_expires_at = time.monotonic() + 3600
        assert auth_client.get_token_for_user() == "user-cached"

    def test_acquires_token_via_ropc(self, auth_client, monkeypatch):
        monkeypatch.setenv("AZURE_USERNAME", "user@example.com")
        monkeypatch.setenv("AZURE_PASSWORD", "password123")

        mock_app = MagicMock()
        mock_app.get_accounts.return_value = []
        mock_app.acquire_token_by_username_password.return_value = {
            "access_token": "user-tok",
            "expires_in": 3600,
        }
        auth_client._public_app = mock_app
        token = auth_client.get_token_for_user(username="user@example.com")
        assert token == "user-tok"

    def test_uses_silent_acquire_for_existing_account(self, auth_client, monkeypatch):
        monkeypatch.setenv("AZURE_USERNAME", "user@example.com")

        mock_account = MagicMock()
        mock_app = MagicMock()
        mock_app.get_accounts.return_value = [mock_account]
        mock_app.acquire_token_silent.return_value = {
            "access_token": "silent-user-tok",
            "expires_in": 3600,
        }
        auth_client._public_app = mock_app
        token = auth_client.get_token_for_user(username="user@example.com")
        assert token == "silent-user-tok"
        mock_app.acquire_token_by_username_password.assert_not_called()


# ---------------------------------------------------------------------------
# Device code flow
# ---------------------------------------------------------------------------

class TestGetTokenInteractive:
    def test_returns_cached_token(self, auth_client):
        auth_client._access_token = "interactive-cached"
        auth_client._token_expires_at = time.monotonic() + 3600
        assert auth_client.get_token_interactive() == "interactive-cached"

    def test_device_code_flow(self, auth_client, capsys):
        mock_app = MagicMock()
        mock_app.initiate_device_flow.return_value = {
            "user_code": "ABC123",
            "message": "Go to https://microsoft.com/devicelogin and enter ABC123",
        }
        mock_app.acquire_token_by_device_flow.return_value = {
            "access_token": "device-tok",
            "expires_in": 3600,
        }
        auth_client._public_app = mock_app
        token = auth_client.get_token_interactive()
        assert token == "device-tok"
        captured = capsys.readouterr()
        assert "ABC123" in captured.out

    def test_raises_when_device_flow_fails(self, auth_client):
        mock_app = MagicMock()
        mock_app.initiate_device_flow.return_value = {
            "error": "authorization_pending",
            "error_description": "Flow init failed",
        }
        auth_client._public_app = mock_app
        with pytest.raises(AuthenticationError, match="device flow"):
            auth_client.get_token_interactive()


# ---------------------------------------------------------------------------
# Authorization code flow
# ---------------------------------------------------------------------------

class TestAuthCodeFlow:
    def test_get_authorization_url(self, auth_client):
        mock_app = MagicMock()
        mock_app.initiate_auth_code_flow.return_value = {
            "auth_uri": "https://login.microsoftonline.com/authorize?...",
            "state": "random-state-value",
        }
        auth_client._confidential_app = mock_app
        url, state = auth_client.get_authorization_url()
        assert url == "https://login.microsoftonline.com/authorize?..."
        assert state == "random-state-value"

    def test_exchange_auth_code(self, auth_client):
        # First set up the flow
        mock_app = MagicMock()
        mock_app.initiate_auth_code_flow.return_value = {
            "auth_uri": "https://auth.example.com",
            "state": "state-xyz",
        }
        auth_client._confidential_app = mock_app
        auth_client.get_authorization_url()

        mock_app.acquire_token_by_auth_code_flow.return_value = {
            "access_token": "auth-code-tok",
            "expires_in": 3600,
        }
        token = auth_client.exchange_auth_code({"code": "code-abc", "state": "state-xyz"})
        assert token == "auth-code-tok"

    def test_exchange_auth_code_without_prior_flow_raises(self, auth_client):
        mock_app = MagicMock()
        auth_client._confidential_app = mock_app
        with pytest.raises(AuthenticationError, match="get_authorization_url"):
            auth_client.exchange_auth_code({"code": "x", "state": "y"})


# ---------------------------------------------------------------------------
# refresh_token
# ---------------------------------------------------------------------------

class TestRefreshToken:
    def test_refresh_invalidates_cache_and_re_acquires(self, auth_client):
        auth_client._access_token = "old-token"
        auth_client._token_expires_at = time.monotonic() + 3600

        mock_app = MagicMock()
        mock_app.acquire_token_silent.return_value = None
        mock_app.acquire_token_for_client.return_value = {
            "access_token": "refreshed-token",
            "expires_in": 3600,
        }
        auth_client._confidential_app = mock_app

        token = auth_client.refresh_token()
        assert token == "refreshed-token"
        assert auth_client._access_token == "refreshed-token"
