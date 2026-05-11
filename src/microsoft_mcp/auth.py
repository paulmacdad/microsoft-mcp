import os
import msal
import pathlib as pl
from typing import NamedTuple
from dotenv import load_dotenv

load_dotenv()

CACHE_FILE = pl.Path.home() / ".microsoft_mcp_token_cache.json"
SCOPES = ["https://graph.microsoft.com/.default"]


class Account(NamedTuple):
    username: str
    account_id: str


def _read_cache() -> str | None:
    try:
        return CACHE_FILE.read_text()
    except FileNotFoundError:
        return None


def _write_cache(content: str) -> None:
    CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
    CACHE_FILE.write_text(content)


def get_app() -> msal.PublicClientApplication:
    client_id = os.getenv("MICROSOFT_MCP_CLIENT_ID")
    if not client_id:
        raise ValueError("MICROSOFT_MCP_CLIENT_ID environment variable is required")

    tenant_id = os.getenv("MICROSOFT_MCP_TENANT_ID", "common")
    authority = f"https://login.microsoftonline.com/{tenant_id}"

    cache = msal.SerializableTokenCache()
    cache_content = _read_cache()
    if cache_content:
        cache.deserialize(cache_content)

    app = msal.PublicClientApplication(
        client_id, authority=authority, token_cache=cache
    )

    return app


def _find_account(
    accounts: list[dict], account_id: str | None
) -> dict | None:
    """Find an account by home_account_id or username (email).

    The MCP tools expose account_id as home_account_id, but callers
    sometimes pass the username/email instead.  Try both.
    """
    if not account_id:
        return accounts[0] if accounts else None

    # Try exact home_account_id match first
    for a in accounts:
        if a["home_account_id"] == account_id:
            return a

    # Fall back to case-insensitive username/email match
    needle = account_id.lower()
    for a in accounts:
        if a.get("username", "").lower() == needle:
            return a

    return None


def get_token(account_id: str | None = None) -> str:
    app = get_app()

    accounts = app.get_accounts()
    account = _find_account(accounts, account_id)

    if account is None and account_id:
        raise Exception(
            f"No cached account matching '{account_id}'. "
            f"Available accounts: {[a.get('username') for a in accounts]}. "
            f"Run authenticate_account + complete_authentication first."
        )

    result = app.acquire_token_silent(SCOPES, account=account)

    if not result:
        if account_id:
            # Never silently start a new device flow for API calls —
            # that blocks the MCP server waiting for interactive input
            # that will never arrive.
            raise Exception(
                f"Token refresh failed for '{account_id}'. "
                f"The refresh token may have expired. "
                f"Run authenticate_account + complete_authentication to re-authenticate."
            )
        # Only fall back to device flow when no account_id was specified
        # (i.e. a manual/interactive caller, not an MCP tool).
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise Exception(
                f"Failed to get device code: {flow.get('error_description', 'Unknown error')}"
            )
        verification_uri = flow.get(
            "verification_uri",
            flow.get("verification_url", "https://microsoft.com/devicelogin"),
        )
        print(
            f"\nTo authenticate:\n1. Visit {verification_uri}\n2. Enter code: {flow['user_code']}"
        )
        result = app.acquire_token_by_device_flow(flow)

    if "error" in result:
        raise Exception(
            f"Auth failed: {result.get('error_description', result['error'])}"
        )

    cache = app.token_cache
    if isinstance(cache, msal.SerializableTokenCache) and cache.has_state_changed:
        _write_cache(cache.serialize())

    return result["access_token"]


def list_accounts() -> list[Account]:
    app = get_app()
    return [
        Account(username=a["username"], account_id=a["home_account_id"])
        for a in app.get_accounts()
    ]


def authenticate_new_account() -> Account | None:
    """Authenticate a new account interactively"""
    app = get_app()

    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise Exception(
            f"Failed to get device code: {flow.get('error_description', 'Unknown error')}"
        )

    print("\nTo authenticate:")
    print(
        f"1. Visit: {flow.get('verification_uri', flow.get('verification_url', 'https://microsoft.com/devicelogin'))}"
    )
    print(f"2. Enter code: {flow['user_code']}")
    print("3. Sign in with your Microsoft account")
    print("\nWaiting for authentication...")

    result = app.acquire_token_by_device_flow(flow)

    if "error" in result:
        raise Exception(
            f"Auth failed: {result.get('error_description', result['error'])}"
        )

    cache = app.token_cache
    if isinstance(cache, msal.SerializableTokenCache) and cache.has_state_changed:
        _write_cache(cache.serialize())

    # Get the newly added account
    accounts = app.get_accounts()
    if accounts:
        # Find the account that matches the token we just got
        for account in accounts:
            if (
                account.get("username", "").lower()
                == result.get("id_token_claims", {})
                .get("preferred_username", "")
                .lower()
            ):
                return Account(
                    username=account["username"], account_id=account["home_account_id"]
                )
        # If exact match not found, return the last account
        account = accounts[-1]
        return Account(
            username=account["username"], account_id=account["home_account_id"]
        )

    return None
