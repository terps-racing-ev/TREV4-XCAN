Goal: turn an excel spreadsheet into a two DBC's. one for Controls, and one for DAQ.

Subsheets: Overview, Controls Bus, DAQ Bus, Messages, Templates

Overview: not needed for this.

'Controls Bus': contains a table called 'ControlsBus'
A table of signals, with the following headers:
'Message' - the name of the message the signal is in, referenced from the 'ControlsMessages' table in the Messages subsheet
'CAN ID' - the id of the message in hex. examples: '0x000000A0'. since there are 8 hex digit, this is an extended id. remember to add 2^31. if it were instead '0x0A0', it is a standard ID).
'Signal Name' - the name of the signal
'Start Byte' - the PHYSICAL start byte of the signal. for little endian this is straightforward, but for big endian it is trickier. Make sure you understand what the start bit in the dbc of a big endian signal looks like.
'Bit Offset' - the bit offset from the start byte. not applicable for big endian signals
'Bit Length' - the bit length of the message.
'Template' - the name of the template the signal will be decoded with, referenced from the 'Templates' table in the Templates subsheet
'Notes' - not needed for this.


'DAQ Bus': contains a table called 'DAQBus'
Same gist as controls bus. Note that message is now referenced from the 'DAQMessages" table (also in the messages subsheet). however, templates are still referenced from the universal 'Templates' table.

'Messages': contains two tables, 'ControlsMessages' and 'DAQMessages'
each table has the following headers:
'Message Name' - this is what the bus subsheets actually reference.
'CAN ID' - the bus subsheet also looks up this value. it is a column in the bus tables purely for convenience of a reader.
'Rate (ms)' - not needed for this.
'Notes' - not needed for this.


'Templates': contains a table called 'Templates' with the following headers:
'Template Name' - the name of the template that the bus tables reference.
'Endianness' - either 'Little' or 'Big'
'Signedness' - either 'Signed' or 'Unsigned'
'Scale' - a number like 1 or 0.1, etc.
'Offset' - a number
'Min' - a number
'Max' - a number
'Units' - text (this is allowed to be blank)
'Enum (0 indexed, separate by ',')' = a string, for example: 'REVERSE, FORWARD' means 0 corresponds to REVERSE and 1 corresponds to FORWARD. make a value table in the DBC

It is not guaranteed that the table will be sorted in any particular order. Thus, it is important to process the entire bus table, assigning  signals to messages, before generating the dbc.

Accessing the excel file:
You will access the excel file on SharePoint via Microsoft Graph API. Below is some sample code for accessing the sheet. The path to the sheet is "/_Electrical-EV26/Electrical Architecture/XCAN.xlsx". The graph base, hostname, site path, etc are the same as the sample code (used for finance purposes) below. 

load_dotenv()
CLIENT_ID = os.getenv("CLIENT_ID")
TENANT_ID = os.getenv("TENANT_ID")
# -----------------------------
# Config constants
# -----------------------------
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SCOPES = ["Sites.Selected"]

SITE_HOSTNAME = "umd0.sharepoint.com"
SITE_PATH = "/TeamsTerpsRacingEV"
VENDOR_INFO_PATH = "/General/_EV26/Finance/ECE/Vendor Info.xlsx"
VENDOR_SHEET = "Info"

# Finance workbook with many order sheets that end with "Subteam"
FINANCE_BOOK_PATH = "/General/_EV26/Finance/Master Finance Sheet.xlsx"
ORDER_SHEET_SUFFIX = "Subsheet"
OLD_APPROVAL = "Approved"

# The global requests session used for all api calls
s = requests.Session()
s.headers.update({"Accept": "application/json"})

# -----------------------------
# General request function
# -----------------------------
def _request(method: str, url: str, **kwargs) -> requests.Response:
    for attempt in range(6):
        r = s.request(method, url, timeout=60, **kwargs)
        if r.status_code in (429, 500, 502, 503, 504):
            delay = float(r.headers.get("Retry-After", 0)) or min(0.5 * (2 ** attempt), 8.0)
            time.sleep(delay)
            continue
        return r
    return r

# -----------------------------
# GET an item
# -----------------------------
def graph_get(url: str, **kwargs) -> Any:
    r = _request("GET", url, **kwargs)
    if not r.ok:
        raise RuntimeError(f"GET {url} -> {r.status_code}: {r.text}")
    return r.json()

# -----------------------------
# POST an item
# -----------------------------
def graph_post(url: str, payload: Any, headers: dict | None = None) -> Any:
    """POST an item to the Graph API."""
    hdrs = {"Content-Type": "application/json"}
    if headers:
        hdrs.update(headers)
    r = _request("POST", url, data=json.dumps(payload), headers=hdrs)
    if not r.ok:
        raise RuntimeError(f"POST {url} -> {r.status_code}: {r.text}")
    return r.json()


def upload_file(drive_id: str, parent_id: str, filename: str, content: bytes) -> dict:
    """Uploads a file to a specific location in a drive."""
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{parent_id}:/{filename}:/content"
    r = _request("PUT", url, data=content, headers={"Content-Type": "application/octet-stream"})
    if not r.ok:
        raise RuntimeError(f"PUT {url} -> {r.status_code}: {r.text}")
    return r.json()


# -----------------------------
# Auth (device code)
# -----------------------------
def login_device_code() -> str:
    app = msal.PublicClientApplication(client_id=CLIENT_ID, authority=f"https://login.microsoftonline.com/{TENANT_ID}")
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(f"Failed to create device flow: {flow}")
    print("\n=== Microsoft sign-in ===")
    print(flow["message"])
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(f"Token acquisition failed: {result}")
    token = result["access_token"]
    s.headers["Authorization"] = f"Bearer {token}"
    return token

# -----------------------------
# Site/Drive/Item helpers
# -----------------------------
def resolve_site_id(hostname: str, site_path: str) -> str:
    url = f"{GRAPH_BASE}/sites/{hostname}:/sites{site_path}"
    site = graph_get(url)
    return site["id"]
# Relies on ^
def get_default_drive_id(site_id: str) -> str:
    url = f"{GRAPH_BASE}/sites/{site_id}/drive"
    drive = graph_get(url)
    return drive["id"]
# Relies on ^
def get_item_by_path(drive_id: str, path: str) -> dict:
    url = f"{GRAPH_BASE}/drives/{drive_id}/root:{path}"
    return graph_get(url)

# -----------------------------
# Excel helpers
# -----------------------------
def create_excel_session(drive_id: str, item_id: str, persist: bool = False) -> str:
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/workbook/createSession"
    data = graph_post(url, {"persistChanges": persist})
    return data["id"]

def excel_used_range_values(drive_id: str, item_id: str, worksheet: str, session_id: str) -> List[List[Any]]:
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/workbook/worksheets('{worksheet}')/usedRange(valuesOnly=true)?$select=values"
    data = graph_get(url, headers={"workbook-session-id": session_id})
    return data.get("values", [])

def list_worksheets(drive_id: str, item_id: str, session_id: str) -> List[str]:
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/workbook/worksheets?$select=name"
    data = graph_get(url, headers={"workbook-session-id": session_id})
    return [w["name"] for w in data.get("value", [])]



Output: Two LOCAL DBC's, 'ControlsBus.dbc' and 'DAQBus.dbc'. You do not need to upload anything to sharepoint. Do NOT modify the excel sheet under any circumstances. Only read from it. Throw an error in the program if signals overlap, or if there are any other issues.

You do not need to dig through any existing dbc's or look for patterns in my filesystem. This is a fresh project starting from scratch, and all the information about the excel sheet is noted above. Use traditional methods for setting up the DBC. You may import any libraries that will be helpful.

One more thing: For BU_, All signals/messages will be prefixed with the component that is sending that message, for example, VCU_Torque_Message. don't remove the prefixes from the signal and message, but you may use them for the BU (even though it technically doesn't matter)

Come up with a plan to create this tool. Let me know if you have any clarifying questions.
