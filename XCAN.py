"""
XCAN.py — Read XCAN.xlsx from SharePoint and generate ControlsBus.dbc / DAQBus.dbc
"""

import json
import os
import time
from dataclasses import dataclass, field
from typing import Any, Dict, List, Tuple

import msal
import requests
from dotenv import load_dotenv

# ─────────────────────────────────────────────
# Config
# ─────────────────────────────────────────────
load_dotenv()
CLIENT_ID = os.getenv("CLIENT_ID")
TENANT_ID = os.getenv("TENANT_ID")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SCOPES = ["Sites.Selected"]

SITE_HOSTNAME = "umd0.sharepoint.com"
SITE_PATH = "/TeamsTerpsRacingEV"
WORKBOOK_PATH = "/_Electrical-EV26/Electrical Architecture/XCAN.xlsx"

s = requests.Session()
s.headers.update({"Accept": "application/json"})


# ─────────────────────────────────────────────
# Graph API helpers
# ─────────────────────────────────────────────
def _request(method: str, url: str, **kwargs) -> requests.Response:
    for attempt in range(6):
        r = s.request(method, url, timeout=60, **kwargs)
        if r.status_code in (429, 500, 502, 503, 504):
            delay = float(r.headers.get("Retry-After", 0)) or min(0.5 * (2 ** attempt), 8.0)
            time.sleep(delay)
            continue
        return r
    return r


def graph_get(url: str, **kwargs) -> Any:
    r = _request("GET", url, **kwargs)
    if not r.ok:
        raise RuntimeError(f"GET {url} -> {r.status_code}: {r.text}")
    return r.json()


def login_device_code() -> str:
    app = msal.PublicClientApplication(
        client_id=CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
    )
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


def resolve_site_id(hostname: str, site_path: str) -> str:
    url = f"{GRAPH_BASE}/sites/{hostname}:/sites{site_path}"
    return graph_get(url)["id"]


def get_default_drive_id(site_id: str) -> str:
    url = f"{GRAPH_BASE}/sites/{site_id}/drive"
    return graph_get(url)["id"]


def get_item_by_path(drive_id: str, path: str) -> dict:
    url = f"{GRAPH_BASE}/drives/{drive_id}/root:{path}"
    return graph_get(url)


def create_excel_session(drive_id: str, item_id: str, persist: bool = False) -> str:
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/workbook/createSession"
    hdrs = {"Content-Type": "application/json"}
    r = _request("POST", url, data=json.dumps({"persistChanges": persist}), headers=hdrs)
    if not r.ok:
        raise RuntimeError(f"POST {url} -> {r.status_code}: {r.text}")
    return r.json()["id"]


def read_excel_table(
    drive_id: str, item_id: str, table_name: str, session_id: str
) -> List[dict]:
    """
    Read an Excel table by name using the Graph API table endpoints.
    Returns a list of dicts keyed by column name.
    """
    base = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/workbook/tables('{table_name}')"
    sess_hdr = {"workbook-session-id": session_id}

    # Get column names
    cols_data = graph_get(f"{base}/columns?$select=name", headers=sess_hdr)
    headers = [col["name"] for col in cols_data["value"]]

    # Get row values
    rows_data = graph_get(f"{base}/rows?$select=values", headers=sess_hdr)
    rows: List[dict] = []
    for row_obj in rows_data["value"]:
        vals = row_obj["values"][0]  # values is [[cell, cell, ...]]
        row = {}
        for header, val in zip(headers, vals):
            row[header] = _normalise(val)
        rows.append(row)
    return rows


# ─────────────────────────────────────────────
# Data classes
# ─────────────────────────────────────────────
@dataclass
class TemplateInfo:
    name: str
    endianness: str          # "Little" or "Big"
    signedness: str          # "Signed" or "Unsigned"
    scale: float
    offset: float
    min_val: float
    max_val: float
    units: str
    enum_str: str            # e.g. "REVERSE, FORWARD" or ""


@dataclass
class MessageInfo:
    name: str
    can_id_raw: int          # numeric value parsed from hex string
    is_extended: bool


@dataclass
class SignalRow:
    message_name: str
    signal_name: str
    start_byte: int
    bit_offset: int
    bit_length: int
    template_name: str


@dataclass
class Signal:
    name: str
    dbc_start_bit: int
    bit_length: int
    is_big_endian: bool
    is_signed: bool
    scale: float
    offset: float
    min_val: float
    max_val: float
    units: str
    enum_pairs: List[Tuple[int, str]]   # [(0,"REVERSE"),(1,"FORWARD")]
    # for overlap detection
    physical_bits: set = field(default_factory=set)


@dataclass
class Message:
    name: str
    can_id_dbc: int          # with bit 31 set if extended
    is_extended: bool
    transmitter: str
    signals: List[Signal] = field(default_factory=list)


# ─────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────
def _normalise(val: Any) -> str:
    """Return a stripped string representation of a cell value."""
    if val is None:
        return ""
    return str(val).strip()


# ─────────────────────────────────────────────
# Table → domain object parsers
# ─────────────────────────────────────────────
def parse_templates(rows: List[dict]) -> Dict[str, TemplateInfo]:
    templates: Dict[str, TemplateInfo] = {}
    for r in rows:
        name = r["Template Name"]
        if not name:
            continue
        templates[name] = TemplateInfo(
            name=name,
            endianness=r["Endianness"],
            signedness=r["Signedness"],
            scale=float(r["Scale"]),
            offset=float(r["Offset"]),
            min_val=float(r["Min"]),
            max_val=float(r["Max"]),
            units=r.get("Units", ""),
            enum_str=r.get("Enum (0 indexed, separate by ',')", ""),
        )
    return templates


def parse_messages(rows: List[dict]) -> Dict[str, MessageInfo]:
    messages: Dict[str, MessageInfo] = {}
    for r in rows:
        name = r["Message Name"]
        if not name:
            continue
        raw_id, ext = parse_can_id(r["CAN ID"])
        messages[name] = MessageInfo(name=name, can_id_raw=raw_id, is_extended=ext)
    return messages


def parse_bus_signals(rows: List[dict]) -> List[SignalRow]:
    signals: List[SignalRow] = []
    for r in rows:
        sig_name = r["Signal Name"]
        if not sig_name:
            continue
        bit_offset_str = r["Bit Offset"]
        bit_offset = int(float(bit_offset_str)) if bit_offset_str not in ("", "N/A", "-") else 0
        signals.append(
            SignalRow(
                message_name=r["Message"],
                signal_name=sig_name,
                start_byte=int(float(r["Start Byte"])),
                bit_offset=bit_offset,
                bit_length=int(float(r["Bit Length"])),
                template_name=r["Template"],
            )
        )
    return signals


# ─────────────────────────────────────────────
# CAN ID helper
# ─────────────────────────────────────────────
def parse_can_id(hex_str: str) -> Tuple[int, bool]:
    """
    Parse a hex CAN ID string like '0x000000A0' or '0x0A0'.
    Returns (numeric_id, is_extended).
    8 hex digits → extended (29-bit). Otherwise standard (11-bit).
    """
    cleaned = hex_str.strip()
    if cleaned.lower().startswith("0x"):
        cleaned = cleaned[2:]
    is_extended = len(cleaned) >= 8
    raw_id = int(cleaned, 16)
    return raw_id, is_extended


# ─────────────────────────────────────────────
# Physical bit set computation (for overlap)
# ─────────────────────────────────────────────
def physical_bits_le(start_byte: int, bit_offset: int, bit_length: int) -> set:
    """Physical bit positions for a little-endian signal."""
    base = start_byte * 8 + bit_offset
    return {base + i for i in range(bit_length)}


def physical_bits_be(start_byte: int, bit_length: int) -> set:
    """Physical bit positions for a big-endian signal (MSB at bit 7 of start_byte)."""
    bits = set()
    for i in range(bit_length):
        byte_idx = start_byte + i // 8
        bit_in_byte = 7 - (i % 8)
        bits.add(byte_idx * 8 + bit_in_byte)
    return bits


# ─────────────────────────────────────────────
# DBC start bit
# ─────────────────────────────────────────────
def dbc_start_bit(start_byte: int, bit_offset: int, is_big_endian: bool) -> int:
    if is_big_endian:
        return start_byte * 8 + 7
    else:
        return start_byte * 8 + bit_offset


# ─────────────────────────────────────────────
# Enum parsing
# ─────────────────────────────────────────────
def parse_enum(enum_str: str) -> List[Tuple[int, str]]:
    """Parse 'REVERSE, FORWARD' → [(0,'REVERSE'),(1,'FORWARD')]."""
    if not enum_str:
        return []
    parts = [p.strip() for p in enum_str.split(",")]
    return [(i, p) for i, p in enumerate(parts) if p]


# ─────────────────────────────────────────────
# Build messages for one bus
# ─────────────────────────────────────────────
def build_bus(
    signal_rows: List[SignalRow],
    messages_table: Dict[str, MessageInfo],
    templates: Dict[str, TemplateInfo],
    bus_label: str,
) -> Tuple[List[Message], List[str]]:
    """
    Returns (list_of_Message, sorted_node_names).
    Raises on missing references or signal overlap.
    """
    # Group signals by message name
    groups: Dict[str, List[SignalRow]] = {}
    for sr in signal_rows:
        groups.setdefault(sr.message_name, []).append(sr)

    nodes: set = set()
    messages: List[Message] = []

    for msg_name, srows in groups.items():
        # Look up message
        if msg_name not in messages_table:
            raise RuntimeError(
                f"[{bus_label}] Signal(s) reference message '{msg_name}' "
                f"which is not in the messages table"
            )
        minfo = messages_table[msg_name]
        can_id_dbc = minfo.can_id_raw | 0x80000000 if minfo.is_extended else minfo.can_id_raw

        # Extract transmitter node from prefix before first '_'
        transmitter = msg_name.split("_")[0] if "_" in msg_name else msg_name
        nodes.add(transmitter)

        msg = Message(
            name=msg_name,
            can_id_dbc=can_id_dbc,
            is_extended=minfo.is_extended,
            transmitter=transmitter,
        )

        for sr in srows:
            # Look up template
            if sr.template_name not in templates:
                raise RuntimeError(
                    f"[{bus_label}] Signal '{sr.signal_name}' references template "
                    f"'{sr.template_name}' which is not in the templates table"
                )
            tmpl = templates[sr.template_name]
            is_be = tmpl.endianness.lower() == "big"
            is_signed = tmpl.signedness.lower() == "signed"

            start = dbc_start_bit(sr.start_byte, sr.bit_offset, is_be)

            if is_be:
                phys = physical_bits_be(sr.start_byte, sr.bit_length)
            else:
                phys = physical_bits_le(sr.start_byte, sr.bit_offset, sr.bit_length)

            sig = Signal(
                name=sr.signal_name,
                dbc_start_bit=start,
                bit_length=sr.bit_length,
                is_big_endian=is_be,
                is_signed=is_signed,
                scale=tmpl.scale,
                offset=tmpl.offset,
                min_val=tmpl.min_val,
                max_val=tmpl.max_val,
                units=tmpl.units,
                enum_pairs=parse_enum(tmpl.enum_str),
                physical_bits=phys,
            )
            msg.signals.append(sig)

        # Overlap detection
        for i, a in enumerate(msg.signals):
            for b in msg.signals[i + 1 :]:
                overlap = a.physical_bits & b.physical_bits
                if overlap:
                    raise RuntimeError(
                        f"[{bus_label}] Signal overlap in message '{msg_name}': "
                        f"'{a.name}' and '{b.name}' share physical bits {sorted(overlap)}"
                    )

        messages.append(msg)

    # Sort messages by CAN ID for deterministic output
    messages.sort(key=lambda m: m.can_id_dbc)

    return messages, sorted(nodes)


# ─────────────────────────────────────────────
# DBC generation
# ─────────────────────────────────────────────
def generate_dbc(messages: List[Message], nodes: List[str]) -> str:
    lines: List[str] = []

    lines.append('VERSION ""')
    lines.append("")
    lines.append("NS_ :")
    lines.append("")
    lines.append("BS_:")
    lines.append("")
    lines.append("BU_: " + " ".join(nodes))
    lines.append("")

    # Messages and signals
    for msg in messages:
        lines.append(f"BO_ {msg.can_id_dbc} {msg.name}: 8 {msg.transmitter}")
        for sig in msg.signals:
            byte_order = 0 if sig.is_big_endian else 1
            sign_char = "-" if sig.is_signed else "+"
            lines.append(
                f" SG_ {sig.name} : {sig.dbc_start_bit}|{sig.bit_length}"
                f"@{byte_order}{sign_char}"
                f" ({sig.scale},{sig.offset})"
                f" [{sig.min_val}|{sig.max_val}]"
                f' "{sig.units}"'
                f" Vector__XXX"
            )
        lines.append("")

    # Value tables
    for msg in messages:
        for sig in msg.signals:
            if sig.enum_pairs:
                val_entries = " ".join(f'{v} "{d}"' for v, d in sig.enum_pairs)
                lines.append(f"VAL_ {msg.can_id_dbc} {sig.name} {val_entries} ;")

    lines.append("")
    return "\n".join(lines)


# ─────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────
def main():
    login_device_code()

    print("Resolving SharePoint site...")
    site_id = resolve_site_id(SITE_HOSTNAME, SITE_PATH)
    drive_id = get_default_drive_id(site_id)
    item = get_item_by_path(drive_id, WORKBOOK_PATH)
    item_id = item["id"]

    print("Creating read-only Excel session...")
    session_id = create_excel_session(drive_id, item_id, persist=False)

    print("Reading tables...")
    templates_rows = read_excel_table(drive_id, item_id, "Templates", session_id)
    ctrl_msg_rows = read_excel_table(drive_id, item_id, "ControlsMessages", session_id)
    daq_msg_rows = read_excel_table(drive_id, item_id, "DAQMessages", session_id)
    ctrl_bus_rows = read_excel_table(drive_id, item_id, "ControlsBus", session_id)
    daq_bus_rows = read_excel_table(drive_id, item_id, "DAQBus", session_id)

    print("Parsing templates...")
    templates = parse_templates(templates_rows)
    print(f"  Found {len(templates)} templates")

    print("Parsing message tables...")
    ctrl_messages = parse_messages(ctrl_msg_rows)
    daq_messages = parse_messages(daq_msg_rows)
    print(f"  ControlsMessages: {len(ctrl_messages)}, DAQMessages: {len(daq_messages)}")

    print("Parsing bus tables...")
    ctrl_signals = parse_bus_signals(ctrl_bus_rows)
    daq_signals = parse_bus_signals(daq_bus_rows)
    print(f"  Controls Bus signals: {len(ctrl_signals)}, DAQ Bus signals: {len(daq_signals)}")

    print("Building Controls Bus DBC...")
    ctrl_msgs, ctrl_nodes = build_bus(ctrl_signals, ctrl_messages, templates, "ControlsBus")

    print("Building DAQ Bus DBC...")
    daq_msgs, daq_nodes = build_bus(daq_signals, daq_messages, templates, "DAQBus")

    print("Writing DBC files...")
    ctrl_dbc = generate_dbc(ctrl_msgs, ctrl_nodes)
    with open("ControlsBus.dbc", "w", newline="\n") as f:
        f.write(ctrl_dbc)

    daq_dbc = generate_dbc(daq_msgs, daq_nodes)
    with open("DAQBus.dbc", "w", newline="\n") as f:
        f.write(daq_dbc)

    print(f"\nControlsBus.dbc: {len(ctrl_msgs)} messages, "
          f"{sum(len(m.signals) for m in ctrl_msgs)} signals")
    print(f"DAQBus.dbc:      {len(daq_msgs)} messages, "
          f"{sum(len(m.signals) for m in daq_msgs)} signals")
    print("Done.")


if __name__ == "__main__":
    main()
