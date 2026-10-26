import json
import os
import re
import threading
from typing import List
from urllib.parse import quote
import msal
import requests

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
SCOPES = ["Mail.Read"]
DEFAULT_MAX_FOLDER_MESSAGES = 100
HARD_MAX_FOLDER_MESSAGES = 200
MAX_MESSAGES_PER_BODY_CALL = 25
REQUEST_TIMEOUT = 30
MAX_BODY_WORDS = 500

# Persist MSAL's token cache so a user normally authenticates only once.
# The cache contains sensitive authentication material, so keep it local and
# out of source control.
CACHE_FILE = os.path.join(os.path.dirname(__file__), "data", "token_cache.json")

_msal_app = None
_token_cache = None
_cache_lock = threading.Lock()
_message_map_lock = threading.Lock()
_message_id_map = {}
_message_map_folder = None

def _load_token_cache():
    """Load MSAL's serialized token cache from disk, if one exists."""
    cache = msal.SerializableTokenCache()

    if os.path.isfile(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                cache.deserialize(f.read())
        except Exception as exc:
            # A corrupt/stale cache should not prevent the app from starting.
            # MSAL can simply authenticate again and replace it.
            print(f"Warning: could not load Outlook token cache: {exc}")

    return cache

def _save_token_cache():
    """Persist MSAL's cache after tokens are added or refreshed."""
    if _token_cache is None or not _token_cache.has_state_changed:
        return

    os.makedirs(os.path.dirname(CACHE_FILE), exist_ok=True)
    temp_file = CACHE_FILE + ".tmp"

    with _cache_lock:
        with open(temp_file, "w", encoding="utf-8") as f:
            f.write(_token_cache.serialize())
        os.replace(temp_file, CACHE_FILE)

        # Best effort on POSIX. Windows protects the file with the user's normal
        # filesystem ACLs; chmod may be ignored there.
        try:
            os.chmod(CACHE_FILE, 0o600)
        except OSError:
            pass

def _get_msal_app():
    global _msal_app, _token_cache
    if _msal_app is not None:
        return _msal_app

    client_id = os.getenv("OUTLOOK_CLIENT_ID")
    if not client_id:
        raise RuntimeError(
            "OUTLOOK_CLIENT_ID is not set. Register a public-client app in "
            "Microsoft Entra ID, grant delegated Mail.Read permission, enable "
            "public client flows, and set OUTLOOK_CLIENT_ID to the app's client ID."
        )

    tenant_id = os.getenv("OUTLOOK_TENANT_ID", "common")
    authority = f"https://login.microsoftonline.com/{tenant_id}"

    _token_cache = _load_token_cache()
    _msal_app = msal.PublicClientApplication(
        client_id,
        authority=authority,
        token_cache=_token_cache,
    )
    return _msal_app

def _get_access_token():
    """Acquire a delegated Microsoft Graph token for the signed-in user."""
    app = _get_msal_app()

    # First try the persistent MSAL cache. If the access token has expired but
    # the refresh token is still valid, acquire_token_silent refreshes it without
    # asking the user to authenticate again.
    for account in app.get_accounts():
        result = app.acquire_token_silent(SCOPES, account=account)
        _save_token_cache()
        if result and "access_token" in result:
            return result["access_token"]

    # Fall back to device-code sign-in only when silent authentication cannot
    # produce a token (for example, on first use or after consent is revoked).
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(
            "Could not start Microsoft device-code authentication: "
            + json.dumps(flow, indent=2)
        )

    print("\n\x1b[36mMicrosoft Outlook sign-in required\x1b[0m")
    print(flow["message"])
    print("The agent will continue after sign-in completes.\n")

    result = app.acquire_token_by_device_flow(flow)
    _save_token_cache()

    if "access_token" not in result:
        message = result.get("error_description") or result.get("error") or "unknown error"
        raise RuntimeError(f"Microsoft sign-in failed: {message}")

    return result["access_token"]

def _graph_get(url, token, params=None, prefer_text=False):
    headers = {"Authorization": f"Bearer {token}"}
    if prefer_text:
        headers["Prefer"] = 'outlook.body-content-type="text"'

    response = requests.get(
        url,
        headers=headers,
        params=params,
        timeout=REQUEST_TIMEOUT,
    )
    if not response.ok:
        try:
            detail = response.json()
        except Exception:
            detail = response.text
        raise RuntimeError(
            f"Microsoft Graph returned HTTP {response.status_code}: {detail}"
        )
    return response.json()

def _paged_values(url, token, params=None, prefer_text=False, limit=None):
    items = []
    next_url = url
    next_params = params

    while next_url:
        data = _graph_get(next_url, token, next_params, prefer_text=prefer_text)
        items.extend(data.get("value", []))
        if limit is not None and len(items) >= limit:
            return items[:limit]
        next_url = data.get("@odata.nextLink")
        next_params = None  # nextLink already contains the query parameters

    return items

def _find_folder(token, folder_path):
    parts = [part.strip() for part in folder_path.replace("\\", "/").split("/") if part.strip()]
    if not parts:
        raise ValueError("folder_path cannot be empty")

    parent_id = None
    traversed = []

    for part in parts:
        if parent_id is None:
            url = f"{GRAPH_ROOT}/me/mailFolders"
        else:
            url = f"{GRAPH_ROOT}/me/mailFolders/{quote(parent_id, safe='')}/childFolders"

        folders = _paged_values(
            url,
            token,
            params={"$select": "id,displayName", "$top": "100"},
        )
        match = next(
            (f for f in folders if f.get("displayName", "").casefold() == part.casefold()),
            None,
        )
        if not match:
            where = "/".join(traversed) or "the mailbox root"
            available = ", ".join(sorted(f.get("displayName", "") for f in folders))
            raise ValueError(
                f"Outlook folder '{part}' was not found under {where}. "
                f"Available folders there: {available or '(none)'}"
            )

        parent_id = match["id"]
        traversed.append(match["displayName"])

    return parent_id, "/".join(traversed)

def _recipient_list(recipients):
    output = []
    for recipient in recipients or []:
        address = (recipient.get("emailAddress") or {})
        name = address.get("name") or ""
        email = address.get("address") or ""
        if name and email:
            output.append(f"{name} <{email}>")
        elif email:
            output.append(email)
        elif name:
            output.append(name)
    return output

def _sender(message):
    sender = ((message.get("from") or {}).get("emailAddress") or {})
    return {
        "name": sender.get("name") or "",
        "address": sender.get("address") or "",
    }

def _truncate_words(text, max_words=MAX_BODY_WORDS):
    """Return at most the first max_words whitespace-delimited words of text."""
    text = (text or "").strip()
    if not text:
        return text, 0, False

    matches = list(re.finditer(r"\S+", text))
    word_count = len(matches)
    if word_count <= max_words:
        return text, word_count, False

    end = matches[max_words - 1].end()
    return text[:end].rstrip(), word_count, True

def list_outlook_folder(
    folder_path: str,
    max_messages: int = DEFAULT_MAX_FOLDER_MESSAGES,
) -> str:
    """
    List messages in an Outlook folder without retrieving their bodies.

    Use this first for large folders. It returns lightweight metadata so you can
    decide which messages are worth reading before calling get_outlook_messages.

    Args:
        folder_path: Outlook folder path, e.g. "Travel/Asia 2026".
        max_messages: Maximum number of message summaries to return (1-200).
            Defaults to 100.

    Returns:
        JSON containing the resolved folder path and message summaries ordered
        oldest to newest. Each summary has a short integer ID that is safe to
        pass back to get_outlook_messages, plus subject, sender, timestamps, and
        whether attachments are present. The underlying Microsoft Graph IDs are
        retained only inside this process and are never exposed to the model.
    """
    global _message_id_map, _message_map_folder

    try:
        max_messages = int(max_messages)
    except (TypeError, ValueError):
        max_messages = DEFAULT_MAX_FOLDER_MESSAGES
    max_messages = max(1, min(max_messages, HARD_MAX_FOLDER_MESSAGES))

    token = _get_access_token()
    folder_id, resolved_path = _find_folder(token, folder_path)

    url = f"{GRAPH_ROOT}/me/mailFolders/{quote(folder_id, safe='')}/messages"
    messages = _paged_values(
        url,
        token,
        params={
            "$select": "id,subject,from,receivedDateTime,sentDateTime,hasAttachments",
            "$top": "100",
        },
        limit=max_messages,
    )

    messages.sort(key=lambda m: m.get("receivedDateTime") or m.get("sentDateTime") or "")

    # Replace Graph's long opaque IDs with simple local integer IDs. This keeps
    # the model from having to copy Graph IDs exactly, which can otherwise lead
    # to ErrorInvalidIdMalformed responses if even one character is changed.
    local_map = {}
    summaries = []
    for local_id, m in enumerate(messages, start=1):
        graph_id = m.get("id")
        if not graph_id:
            continue
        local_map[local_id] = graph_id
        summaries.append({
            "id": local_id,
            "subject": m.get("subject") or "",
            "from": _sender(m),
            "received": m.get("receivedDateTime"),
            "sent": m.get("sentDateTime"),
            "has_attachments": bool(m.get("hasAttachments")),
        })

    # The most recent folder listing is the active lookup table for subsequent
    # get_outlook_messages calls. A lock keeps replacement atomic.
    with _message_map_lock:
        _message_id_map = local_map
        _message_map_folder = resolved_path

    return json.dumps(
        {
            "folder": resolved_path,
            "message_count": len(summaries),
            "messages": summaries,
        },
        indent=2,
        ensure_ascii=False,
    )

def get_outlook_messages(message_ids: List[int]) -> str:
    """
    Retrieve selected Outlook message bodies using short IDs from the latest
    list_outlook_folder call.

    At most 25 messages may be retrieved in one call. Each body is limited to
    its first 500 words. Invalid or stale short IDs are reported per message
    and do not cause the whole batch to fail.

    Args:
        message_ids: Short integer message IDs returned by list_outlook_folder.
            Supply no more than 25 IDs per call.

    Returns:
        JSON containing the selected messages. Each message includes the same
        short ID, subject, sender, recipients, timestamps, attachment presence,
        and at most the first 500 words of its plain-text body.
    """
    if not isinstance(message_ids, list):
        raise ValueError("message_ids must be a list of short integer message IDs")

    # Normalize IDs to integers, preserve caller order, and remove duplicates.
    unique_ids = []
    seen = set()
    for message_id in message_ids:
        try:
            local_id = int(message_id)
        except (TypeError, ValueError):
            continue
        if local_id > 0 and local_id not in seen:
            seen.add(local_id)
            unique_ids.append(local_id)

    if not unique_ids:
        raise ValueError("message_ids cannot be empty")
    if len(unique_ids) > MAX_MESSAGES_PER_BODY_CALL:
        raise ValueError(
            f"get_outlook_messages accepts at most {MAX_MESSAGES_PER_BODY_CALL} "
            "message IDs per call. Select the most relevant messages and retrieve "
            "additional ones in another call only if needed."
        )

    with _message_map_lock:
        id_map = dict(_message_id_map)
        mapped_folder = _message_map_folder

    if not id_map:
        raise RuntimeError(
            "No Outlook folder listing is active. Call list_outlook_folder first, "
            "then pass the short IDs it returns to get_outlook_messages."
        )

    token = _get_access_token()
    cleaned = []

    for local_id in unique_ids:
        graph_id = id_map.get(local_id)
        if not graph_id:
            cleaned.append({
                "id": local_id,
                "error": (
                    f"Unknown message ID {local_id}. Call list_outlook_folder again "
                    "if the folder listing has changed."
                ),
            })
            continue

        url = f"{GRAPH_ROOT}/me/messages/{quote(graph_id, safe='')}"
        try:
            m = _graph_get(
                url,
                token,
                params={
                    "$select": (
                        "id,subject,from,toRecipients,ccRecipients,receivedDateTime,"
                        "sentDateTime,body,hasAttachments"
                    )
                },
                prefer_text=True,
            )
        except Exception as exc:
            # One bad/stale message should not discard the rest of the batch.
            cleaned.append({
                "id": local_id,
                "error": str(exc),
            })
            continue

        body = (m.get("body") or {}).get("content") or ""
        body, body_word_count, body_truncated = _truncate_words(body)
        cleaned.append({
            "id": local_id,
            "subject": m.get("subject") or "",
            "from": _sender(m),
            "to": _recipient_list(m.get("toRecipients")),
            "cc": _recipient_list(m.get("ccRecipients")),
            "received": m.get("receivedDateTime"),
            "sent": m.get("sentDateTime"),
            "has_attachments": bool(m.get("hasAttachments")),
            "body_word_count": body_word_count,
            "body_truncated": body_truncated,
            "body": body,
        })

    return json.dumps(
        {
            "folder": mapped_folder,
            "message_count": len(cleaned),
            "messages": cleaned,
        },
        indent=2,
        ensure_ascii=False,
    )
