# enterprise-meeting-search

Enterprise-wide meeting search and reconstruction tool across Microsoft 365 data sources.

---

## Microsoft Graph API – Setup Guide

### 1. Azure AD App Registration

1. Go to [Azure Portal → Azure Active Directory → App registrations](https://portal.azure.com/#view/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/~/RegisteredApps) and click **New registration**.
2. Give the application a name (e.g. `enterprise-meeting-search`).
3. Choose the appropriate **Supported account types** for your organisation.
4. For the **delegated (user) flow**, add a **Redirect URI** (e.g. `http://localhost:8080`).
5. Click **Register**.

### 2. Required API Permissions

Navigate to **API Permissions → Add a permission → Microsoft Graph** and add the following permissions:

| Permission | Type | Description |
|---|---|---|
| `Calendars.Read` | Application or Delegated | Read user calendar events |
| `Chat.Read` | Application or Delegated | Read Microsoft Teams chat messages |
| `Mail.Read` | Application or Delegated | Read user mail messages |
| `Files.Read` | Application or Delegated | Read user OneDrive files |
| `OnlineMeetings.Read` | Application or Delegated | Read online meeting details |

> **Application permissions** are used by the daemon (app) flow and require admin consent.  
> **Delegated permissions** are used by the user flow and can be consented to by an individual user (unless your tenant requires admin consent for all permissions).

After adding permissions, click **Grant admin consent** if you are using application permissions.

### 3. Client Credentials

1. Navigate to **Certificates & secrets → New client secret**.
2. Enter a description and choose an expiry, then click **Add**.
3. Copy the secret **Value** immediately – it is shown only once.

### 4. Environment Variables

Copy `.env.example` to `.env` and fill in the values:

```bash
cp .env.example .env
```

| Variable | Description |
|---|---|
| `AZURE_TENANT_ID` | Directory (tenant) ID from the app overview page |
| `AZURE_CLIENT_ID` | Application (client) ID from the app overview page |
| `AZURE_CLIENT_SECRET` | Client secret value created above |
| `AZURE_REDIRECT_URI` | Must match a redirect URI in the app registration (delegated flow only) |
| `GRAPH_API_BASE_URL` | Optional; defaults to `https://graph.microsoft.com/v1.0` |

> **Never commit `.env` to source control.**

---

## Installation

```bash
pip install -r requirements.txt
```

---

## Usage

### Application (daemon) permission flow

```python
from src.graph_auth import GraphAuthClient
from src.graph_client import GraphClient

auth = GraphAuthClient()          # reads env vars automatically
client = GraphClient(auth_client=auth, use_app_permissions=True)

events = client.list_calendar_events(user_id="user@example.com")
```

### Delegated (user) permission flow – device code

```python
auth = GraphAuthClient()
token = auth.get_token_interactive()   # prints a device-code prompt

client = GraphClient(auth_client=auth, use_app_permissions=False)
me = client.get_me()
```

### Delegated flow – authorization code (web app)

```python
auth = GraphAuthClient()
url, state = auth.get_authorization_url()
# Redirect the user to `url`, then when they return:
token = auth.exchange_auth_code({"code": "<code>", "state": state})
```

---

## Running the Tests

```bash
pip install -r requirements-dev.txt
pytest
```
