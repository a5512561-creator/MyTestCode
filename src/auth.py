"""Microsoft Graph OAuth 認證（MSAL）。"""
import os
from pathlib import Path
import msal
from dotenv import load_dotenv

load_dotenv(Path(__file__).resolve().parents[1] / ".env")

TENANT_ID = os.getenv("TENANT_ID", "common")
CLIENT_ID = os.getenv("CLIENT_ID", "")
CLIENT_SECRET = os.getenv("CLIENT_SECRET", "")
DEVICE_CODE_FLOW = os.getenv("DEVICE_CODE_FLOW", "").lower() in ("1", "true", "yes")

# 委派權限（裝置碼/互動登入）用個別 scope，不可加 .default
SCOPES_DELEGATED = [
    "User.Read",
    "Tasks.ReadWrite",
    "Calendars.ReadWrite",
    "Notes.ReadWrite",
]
# 應用程式權限（client secret）用 .default
SCOPES_APP = ["https://graph.microsoft.com/.default"]


def get_token_client_credentials():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(scopes=SCOPES_APP)
    if "access_token" in result:
        return result["access_token"]
    raise RuntimeError("取得 token 失敗: " + str(result.get("error_description", result)))


def get_token_device_code():
    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
    )
    flow = app.initiate_device_flow(scopes=SCOPES_DELEGATED)
    if "message" in flow:
        print(flow["message"])
    else:
        raise RuntimeError("Device flow 初始化失敗: " + str(flow))
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" in result:
        return result["access_token"]
    raise RuntimeError("取得 token 失敗: " + str(result.get("error_description", result)))


def get_access_token():
    if DEVICE_CODE_FLOW or not CLIENT_SECRET:
        return get_token_device_code()
    return get_token_client_credentials()
