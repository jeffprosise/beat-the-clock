# Beat the Clock

Demo app from the *Beat the Clock* session at Techorama Netherlands 2026. It is an agentic app whose tools read e-mail from Outlook folders through Microsoft Graph.

This guide explains how to configure an e-mail account so the agent can read from it.

## How authentication works

- The app calls Microsoft Graph using the delegated `Mail.Read` permission.
- Sign-in uses MSAL's device-code flow with a public client (no client secret).
- After the first sign-in, the refresh token is cached in `data/token_cache.json`, so later runs sign in silently. This file contains sensitive authentication material. It is listed in `.gitignore` and must never be committed.

## Setup

### 1. Register an app in Microsoft Entra ID

1. Go to the [Entra admin center](https://entra.microsoft.com) and open **Identity > Applications > App registrations > New registration**. (In the Azure portal: **Microsoft Entra ID > App registrations**.)
2. Give the app a name, such as `beat-the-clock-agent`.
3. Under **Supported account types**, choose based on the mailbox you will read:
   - Work or school account only: single tenant is fine.
   - Personal Outlook.com/Hotmail account, or either kind: choose the option that includes personal Microsoft accounts.
4. Leave **Redirect URI** blank and click **Register**.
5. On the Overview page, copy the **Application (client) ID**. If you registered a single-tenant app, also copy the **Directory (tenant) ID**.

### 2. Enable public client flows

1. Open **Authentication** (it may be labeled **Authentication (Preview)** with a **Settings** tab).
2. Set **Allow public client flows** to **Yes** and save. Device-code sign-in fails without this.

### 3. Add the Microsoft Graph permission

1. Open **API permissions > Add a permission > Microsoft Graph > Delegated permissions**.
2. Select **Mail.Read** and click **Add permissions**.
3. If your work or school tenant requires it, click **Grant admin consent**. Otherwise you will be asked to consent at first sign-in.

### 4. Set environment variables

```
OUTLOOK_CLIENT_ID=<application (client) ID>
OUTLOOK_TENANT_ID=<tenant ID>    # optional
```

`OUTLOOK_TENANT_ID` defaults to `common`. Use the value that matches your account:

| Account | `OUTLOOK_TENANT_ID` |
|---|---|
| Work or school (single-tenant app) | Your tenant ID (required) |
| Personal Microsoft account | `consumers` (or `common` if the app allows personal accounts) |

Set these in your shell or in a `.env` file. `.env` is already in `.gitignore`.

### 5. Install dependencies

```
pip install msal requests
```

Install the app's other dependencies as well (see the imports in `app.py` and `agents.py`).

### 6. Sign in the first time

1. Run the app and trigger a request that reads mail.
2. The console prints a "Microsoft Outlook sign-in required" message with a URL (`https://microsoft.com/devicelogin`) and a code.
3. Open the URL, enter the code, and sign in with the mailbox account.
4. Approve the `Mail.Read` consent prompt.
5. The agent continues automatically. The token is saved to `data/token_cache.json`.

### 7. Prepare the mailbox

- The agent finds folders by display name and supports nested paths such as `Travel/Asia 2026`. Matching is case-insensitive.
- Create the folder in Outlook and move or copy the relevant e-mails into it. Give the agent that folder path.
- If a folder name does not match, the error message lists the folders that exist at that level.

## Troubleshooting

| Symptom | Likely cause and fix |
|---|---|
| `OUTLOOK_CLIENT_ID is not set` | The variable is not visible to the process. Set it and restart. |
| `AADSTS7000218` or a public-client error | **Allow public client flows** was not saved (step 2). |
| `AADSTS50020` or "account does not exist in tenant" | The account type chosen at registration does not match `OUTLOOK_TENANT_ID` (steps 1 and 4). |
| `AADSTS65001` (consent required) | Redo step 3, or ask a tenant admin to grant consent. |
| Prompted to sign in on every run | Delete `data/token_cache.json` and sign in again. |

## Limitations

The agent calls Graph's `/me/...` endpoints, so it can only read the mailbox of the account that signs in. It cannot read other users' or shared mailboxes.

## Security

Never commit `data/token_cache.json`. If it was ever pushed to a repo, remove it from tracking with `git rm --cached` and revoke the app's sessions.
