#!/usr/bin/env python3
"""
Selectel Server Inventory Script
Fetches dedicated and cloud servers via Selectel API
Exports to JSON and Excel (.xlsx)

Requirements:
    pip install requests openpyxl

Usage:
    export SELECTEL_TOKEN="your_iam_token_here"
    python selectel_servers.py

Or pass token directly:
    python selectel_servers.py --token YOUR_TOKEN
"""

import os
import sys
import json
import argparse
import requests
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─── Configuration ────────────────────────────────────────────────────────────

DEDICATED_BASE_URL = "https://api.selectel.ru/servers/v2"
CLOUD_BASE_URL     = "https://api.selectel.ru/vpc/resell/v2"
CLOUD_NOVA_URL     = "https://api.selectel.ru"   # Nova compute API

# ─── Auth helpers ─────────────────────────────────────────────────────────────

def get_token():
    """Get IAM token from env or CLI arg."""
    parser = argparse.ArgumentParser(description="Selectel server inventory")
    parser.add_argument("--token", help="Selectel IAM token")
    parser.add_argument(
        "--keystone-token",
        help="Keystone token for cloud servers (optional, auto-fetched if not provided)"
    )
    args = parser.parse_args()
    token = args.token or os.environ.get("SELECTEL_TOKEN")
    if not token:
        print("ERROR: Provide token via --token or SELECTEL_TOKEN env variable")
        sys.exit(1)
    return token, args.keystone_token


def auth_headers(token):
    return {"X-Auth-Token": token, "Content-Type": "application/json"}


# ─── Dedicated servers ────────────────────────────────────────────────────────

def fetch_dedicated_servers(token):
    """Fetch all dedicated server resources from Selectel API."""
    headers = auth_headers(token)
    servers = []

    # Get resources list (model=server)
    try:
        resp = requests.get(
            f"{DEDICATED_BASE_URL}/resource",
            headers=headers,
            params={"model": "server", "limit": 1000},
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        resources = data.get("result", [])
    except Exception as e:
        print(f"  [!] Could not fetch dedicated server resources: {e}")
        return []

    # Get location names
    locations = {}
    try:
        loc_resp = requests.get(f"{DEDICATED_BASE_URL}/location", headers=headers, timeout=15)
        if loc_resp.ok:
            for loc in loc_resp.json().get("result", []):
                locations[loc["uuid"]] = loc.get("name", loc["uuid"])
    except Exception:
        pass

    for res in resources:
        resource_uuid = res.get("uuid", "")
        server_info = {}

        # Get detailed server resource info
        try:
            detail_resp = requests.get(
                f"{DEDICATED_BASE_URL}/resource/{resource_uuid}/server",
                headers=headers,
                timeout=15,
            )
            if detail_resp.ok:
                server_info = detail_resp.json().get("result", {})
        except Exception:
            pass

        # Get OS info
        os_info = ""
        try:
            os_resp = requests.get(
                f"{DEDICATED_BASE_URL}/resource/{resource_uuid}/os",
                headers=headers,
                timeout=15,
            )
            if os_resp.ok:
                os_data = os_resp.json().get("result", {})
                os_info = os_data.get("name", "") or os_data.get("os_name", "")
        except Exception:
            pass

        # Get IP addresses
        ip_list = []
        try:
            ip_resp = requests.get(
                f"{DEDICATED_BASE_URL}/resource/{resource_uuid}/ip",
                headers=headers,
                timeout=15,
            )
            if ip_resp.ok:
                for ip_item in ip_resp.json().get("result", []):
                    ip_addr = ip_item.get("ip_address") or ip_item.get("address", "")
                    if ip_addr:
                        ip_list.append(ip_addr)
        except Exception:
            pass

        # Extract billing date
        billing_date = ""
        paid_till = res.get("paid_till") or res.get("payment_date") or res.get("expires")
        if paid_till:
            try:
                dt = datetime.fromisoformat(str(paid_till).replace("Z", "+00:00"))
                billing_date = dt.strftime("%Y-%m-%d")
            except Exception:
                billing_date = str(paid_till)

        # Build configuration string
        config_parts = []
        if server_info:
            cpu = server_info.get("cpu_info") or server_info.get("cpu", "")
            ram = server_info.get("ram") or server_info.get("memory", "")
            disk = server_info.get("disk_info") or server_info.get("storage", "")
            if cpu:
                config_parts.append(f"CPU: {cpu}")
            if ram:
                config_parts.append(f"RAM: {ram}")
            if disk:
                config_parts.append(f"Storage: {disk}")

        service_name = res.get("name") or res.get("title", "")
        location_uuid = res.get("location_uuid") or res.get("location", "")
        location_name = locations.get(location_uuid, location_uuid)

        servers.append({
            "name":          service_name or resource_uuid,
            "ip":            ", ".join(ip_list) if ip_list else "—",
            "configuration": " | ".join(config_parts) if config_parts else res.get("service_name", "—"),
            "region":        location_name or "—",
            "os":            os_info or "—",
            "billing_date":  billing_date or "—",
            "uuid":          resource_uuid,
            "status":        res.get("state", res.get("status", "—")),
        })

    return servers


# ─── Cloud servers ────────────────────────────────────────────────────────────

def get_keystone_token(iam_token):
    """Exchange IAM token for Keystone token to access OpenStack APIs."""
    url = "https://api.selectel.ru/identity/v3/auth/tokens"
    payload = {
        "auth": {
            "identity": {
                "methods": ["application_credential"],
                "application_credential": {"token": iam_token}
            }
        }
    }
    try:
        resp = requests.post(url, json=payload, timeout=15)
        if resp.ok:
            return resp.headers.get("X-Subject-Token")
    except Exception:
        pass

    # Alternative: direct token auth
    try:
        payload2 = {
            "auth": {
                "identity": {
                    "methods": ["token"],
                    "token": {"id": iam_token}
                }
            }
        }
        resp2 = requests.post(url, json=payload2, timeout=15)
        if resp2.ok:
            return resp2.headers.get("X-Subject-Token")
    except Exception:
        pass

    return None


def fetch_cloud_projects(token):
    """Fetch cloud projects to get endpoints."""
    headers = auth_headers(token)
    try:
        resp = requests.get(
            f"{CLOUD_BASE_URL}/projects",
            headers=headers,
            timeout=15,
        )
        if resp.ok:
            return resp.json() if isinstance(resp.json(), list) else resp.json().get("projects", [])
    except Exception:
        pass
    return []


def fetch_cloud_servers(iam_token, keystone_token=None):
    """Fetch all cloud (virtual) servers via Selectel resell + OpenStack Nova API."""
    headers = auth_headers(iam_token)
    servers = []

    # Try resell v2 API first for project list
    projects = []
    try:
        resp = requests.get(f"{CLOUD_BASE_URL}/projects", headers=headers, timeout=15)
        if resp.ok:
            raw = resp.json()
            projects = raw if isinstance(raw, list) else raw.get("projects", raw.get("result", []))
    except Exception as e:
        print(f"  [!] Could not fetch cloud projects: {e}")

    if not keystone_token:
        keystone_token = get_keystone_token(iam_token)

    if not keystone_token:
        print("  [!] Could not obtain Keystone token for cloud API. "
              "Try passing --keystone-token manually.")
        return []

    ks_headers = {
        "X-Auth-Token": keystone_token,
        "Content-Type": "application/json",
    }

    # Selectel pools / regions
    regions = [
        ("ru-1", "https://api.selectel.ru/servers/api/v2/servers"),
        ("ru-2", "https://api.selectel.ru/servers/api/v2/servers"),
        ("ru-3", "https://api.selectel.ru/servers/api/v2/servers"),
        ("ru-7", "https://api.selectel.ru/servers/api/v2/servers"),
        ("ru-9", "https://api.selectel.ru/servers/api/v2/servers"),
    ]

    # Use Nova API for each project
    nova_endpoints = [
        "https://api.selectel.ru/v2.1",
        "https://compute.cloud.selectel.ru/v2.1",
    ]

    for project in projects:
        project_id = project.get("id") or project.get("uuid", "")
        if not project_id:
            continue

        for nova_base in nova_endpoints:
            try:
                url = f"{nova_base}/{project_id}/servers/detail"
                resp = requests.get(url, headers=ks_headers, params={"all_tenants": 0}, timeout=20)
                if not resp.ok:
                    continue
                nova_servers = resp.json().get("servers", [])

                for s in nova_servers:
                    # Extract IP
                    ip_list = []
                    for net_name, addrs in s.get("addresses", {}).items():
                        for addr in addrs:
                            if addr.get("OS-EXT-IPS:type") == "floating" or addr.get("version") == 4:
                                ip_list.append(addr.get("addr", ""))

                    # Flavor / configuration
                    flavor = s.get("flavor", {})
                    flavor_name = flavor.get("original_name") or flavor.get("id", "")
                    vcpu = flavor.get("vcpus", "")
                    ram_mb = flavor.get("ram", "")
                    disk_gb = flavor.get("disk", "")
                    config = flavor_name
                    if vcpu and ram_mb:
                        config = f"{flavor_name} ({vcpu} vCPU, {int(ram_mb)//1024}GB RAM, {disk_gb}GB disk)"

                    # OS
                    image = s.get("image", {})
                    os_name = ""
                    if isinstance(image, dict):
                        os_name = image.get("name", "")
                    metadata = s.get("metadata", {})
                    if not os_name:
                        os_name = metadata.get("os_distro", "") or metadata.get("image_name", "")

                    # Region from availability zone
                    az = s.get("OS-EXT-AZ:availability_zone", "")

                    # Billing date - not directly in Nova, use created date as fallback
                    created = s.get("created", "")
                    updated = s.get("updated", "")

                    servers.append({
                        "name":          s.get("name", s.get("id", "unknown")),
                        "ip":            ", ".join(filter(None, ip_list)) or "—",
                        "configuration": config or "—",
                        "region":        az or "—",
                        "os":            os_name or "—",
                        "billing_date":  "—",  # Cloud servers are billed hourly; see billing panel
                        "uuid":          s.get("id", ""),
                        "status":        s.get("status", "—"),
                        "project_id":    project_id,
                    })
                break  # success, stop trying endpoints
            except Exception:
                continue

    return servers


# ─── Export helpers ───────────────────────────────────────────────────────────

COLUMNS = ["name", "ip", "configuration", "region", "os", "billing_date"]
HEADERS = {
    "name":          "Имя сервера",
    "ip":            "IP-адрес",
    "configuration": "Конфигурация",
    "region":        "Регион",
    "os":            "Операционная система",
    "billing_date":  "Дата оплаты",
}


def save_json(dedicated, cloud, path):
    output = {
        "generated_at": datetime.utcnow().isoformat() + "Z",
        "dedicated_servers": dedicated,
        "cloud_servers": cloud,
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"  ✓ JSON saved: {path}")


def _header_fill():
    return PatternFill("solid", start_color="1F3864")


def _alt_fill():
    return PatternFill("solid", start_color="DCE6F1")


def _thin_border():
    s = Side(style="thin", color="AAAAAA")
    return Border(left=s, right=s, top=s, bottom=s)


def _write_sheet(ws, servers, title):
    ws.title = title

    # Title row
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(COLUMNS))
    title_cell = ws.cell(row=1, column=1, value=title)
    title_cell.font = Font(name="Arial", bold=True, size=13, color="FFFFFF")
    title_cell.fill = PatternFill("solid", start_color="1F3864")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # Header row
    for col_idx, col_key in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=2, column=col_idx, value=HEADERS[col_key])
        cell.font = Font(name="Arial", bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", start_color="2E75B6")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _thin_border()
    ws.row_dimensions[2].height = 22

    # Data rows
    for row_idx, server in enumerate(servers, start=3):
        fill = _alt_fill() if row_idx % 2 == 0 else PatternFill("solid", start_color="FFFFFF")
        for col_idx, col_key in enumerate(COLUMNS, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=server.get(col_key, "—"))
            cell.font = Font(name="Arial", size=10)
            cell.fill = fill
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.border = _thin_border()
        ws.row_dimensions[row_idx].height = 18

    # Summary row
    summary_row = len(servers) + 3
    ws.cell(row=summary_row, column=1, value=f"Всего серверов: {len(servers)}")
    ws.cell(row=summary_row, column=1).font = Font(name="Arial", bold=True, italic=True)

    # Column widths
    col_widths = [30, 22, 45, 20, 35, 15]
    for i, width in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width

    # Freeze header
    ws.freeze_panes = "A3"


def save_excel(dedicated, cloud, path):
    wb = Workbook()
    ws_dedicated = wb.active
    _write_sheet(ws_dedicated, dedicated, "Выделенные серверы")

    ws_cloud = wb.create_sheet()
    _write_sheet(ws_cloud, cloud, "Облачные серверы")

    wb.save(path)
    print(f"  ✓ Excel saved: {path}")


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    token, keystone_token = get_token()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    json_path  = f"/mnt/user-data/outputs/selectel_servers_{timestamp}.json"
    excel_path = f"/mnt/user-data/outputs/selectel_servers_{timestamp}.xlsx"

    print("=" * 55)
    print("  Selectel Server Inventory")
    print("=" * 55)

    print("\n[1/2] Fetching dedicated servers...")
    dedicated = fetch_dedicated_servers(token)
    print(f"      Found: {len(dedicated)} dedicated server(s)")

    print("\n[2/2] Fetching cloud (virtual) servers...")
    cloud = fetch_cloud_servers(token, keystone_token)
    print(f"      Found: {len(cloud)} cloud server(s)")

    print("\n[Saving] Exporting results...")
    save_json(dedicated, cloud, json_path)
    save_excel(dedicated, cloud, excel_path)

    print("\n" + "=" * 55)
    print(f"  Done! Files saved:")
    print(f"    JSON:  {json_path}")
    print(f"    Excel: {excel_path}")
    print("=" * 55)


if __name__ == "__main__":
    main()
