#!/usr/bin/env python3
"""
Selectel Server Inventory Script
Fetches dedicated and cloud servers via Selectel API
Exports to JSON and Excel (.xlsx)

Requirements:
    pip install requests openpyxl

Usage:
    python selectel_servers.py \
        --username SERVICE_USER \
        --account  ACCOUNT_ID  \
        --password PASSWORD
"""

import os, sys, json, argparse, requests
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

KEYSTONE_URL   = "https://cloud.api.selcloud.ru/identity/v3/auth/tokens"
DEDICATED_BASE = "https://api.selectel.ru/servers/v2"
RESELL_BASE    = "https://api.selectel.ru/vpc/resell/v2"


# ── Auth ──────────────────────────────────────────────────────────────────────

def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument("--username", default=os.environ.get("SEL_USERNAME"))
    p.add_argument("--account",  default=os.environ.get("SEL_ACCOUNT"))
    p.add_argument("--password", default=os.environ.get("SEL_PASSWORD"))
    args = p.parse_args()
    missing = [k for k, v in [("--username", args.username), ("--account", args.account), ("--password", args.password)] if not v]
    if missing:
        print(f"ERROR: Missing: {', '.join(missing)}")
        print("Use CLI args or env vars: SEL_USERNAME, SEL_ACCOUNT, SEL_PASSWORD")
        sys.exit(1)
    return args


def get_account_token(username, account_id, password):
    """Account-scoped IAM token for dedicated servers and project listing."""
    payload = {"auth": {"identity": {"methods": ["password"], "password": {"user": {
        "name": username, "domain": {"name": account_id}, "password": password
    }}}, "scope": {"domain": {"name": account_id}}}}
    r = requests.post(KEYSTONE_URL, json=payload, timeout=20)
    r.raise_for_status()
    token = r.headers.get("X-Subject-Token")
    if not token:
        raise RuntimeError("No X-Subject-Token in response headers")
    return token


def get_project_token(username, account_id, password, project_name):
    """Project-scoped IAM token + service catalog for OpenStack API."""
    payload = {"auth": {"identity": {"methods": ["password"], "password": {"user": {
        "name": username, "domain": {"name": account_id}, "password": password
    }}}, "scope": {"project": {"name": project_name, "domain": {"name": account_id}}}}}
    r = requests.post(KEYSTONE_URL, json=payload, timeout=20)
    r.raise_for_status()
    token   = r.headers.get("X-Subject-Token")
    catalog = r.json().get("token", {}).get("catalog", [])
    return token, catalog


def get_nova_url(catalog):
    for svc in catalog:
        if svc.get("type") == "compute":
            for ep in svc.get("endpoints", []):
                if ep.get("interface") == "public":
                    return ep.get("url", "").rstrip("/")
    return None


def list_projects(account_token):
    r = requests.get(f"{RESELL_BASE}/projects", headers={"X-Auth-Token": account_token}, timeout=15)
    r.raise_for_status()
    raw = r.json()
    return raw if isinstance(raw, list) else raw.get("projects", [])


# ── Dedicated servers ─────────────────────────────────────────────────────────

def fetch_dedicated(account_token):
    h = {"X-Auth-Token": account_token}
    servers = []

    locations = {}
    try:
        r = requests.get(f"{DEDICATED_BASE}/location", headers=h, timeout=15)
        if r.ok:
            for loc in r.json().get("result", []):
                locations[loc["uuid"]] = loc.get("name", loc["uuid"])
    except Exception:
        pass

    try:
        r = requests.get(f"{DEDICATED_BASE}/resource", headers=h,
                         params={"model": "server", "limit": 1000}, timeout=30)
        r.raise_for_status()
        resources = r.json().get("result", [])
    except Exception as e:
        print(f"  [!] Dedicated resources error: {e}")
        return []

    for res in resources:
        uuid = res.get("uuid", "")

        hw = {}
        try:
            r = requests.get(f"{DEDICATED_BASE}/resource/{uuid}/server", headers=h, timeout=15)
            if r.ok: hw = r.json().get("result", {})
        except Exception: pass

        os_name = "—"
        try:
            r = requests.get(f"{DEDICATED_BASE}/resource/{uuid}/os", headers=h, timeout=15)
            if r.ok:
                d = r.json().get("result", {})
                os_name = d.get("name") or d.get("os_name") or "—"
        except Exception: pass

        ip_list = []
        try:
            r = requests.get(f"{DEDICATED_BASE}/resource/{uuid}/ip", headers=h, timeout=15)
            if r.ok:
                for item in r.json().get("result", []):
                    addr = item.get("ip_address") or item.get("address", "")
                    if addr: ip_list.append(addr)
        except Exception: pass

        billing = "—"
        raw_date = res.get("paid_till") or res.get("payment_date") or res.get("expires")
        if raw_date:
            try:
                dt = datetime.fromisoformat(str(raw_date).replace("Z", "+00:00"))
                billing = dt.strftime("%d.%m.%Y")
            except Exception:
                billing = str(raw_date)

        parts = []
        if hw.get("cpu_info"):  parts.append(f"CPU: {hw['cpu_info']}")
        if hw.get("ram"):       parts.append(f"RAM: {hw['ram']}")
        if hw.get("disk_info"): parts.append(f"Disk: {hw['disk_info']}")
        config = " | ".join(parts) if parts else res.get("service_name", "—")

        loc_uuid = res.get("location_uuid") or res.get("location", "")
        servers.append({
            "name":          res.get("name") or res.get("title") or uuid,
            "ip":            ", ".join(ip_list) or "—",
            "configuration": config,
            "region":        locations.get(loc_uuid, loc_uuid) or "—",
            "os":            os_name,
            "billing_date":  billing,
        })

    return servers


# ── Cloud servers ─────────────────────────────────────────────────────────────

def fetch_cloud(username, account_id, password, projects):
    all_servers = []

    for proj in projects:
        proj_name = proj.get("name", "")
        if not proj_name:
            continue
        print(f"    → project: {proj_name}")

        try:
            token, catalog = get_project_token(username, account_id, password, proj_name)
        except Exception as e:
            print(f"      [!] Auth error: {e}")
            continue

        nova_url = get_nova_url(catalog)
        if not nova_url:
            print(f"      [!] No Nova endpoint in catalog, skipping")
            continue

        try:
            r = requests.get(f"{nova_url}/servers/detail",
                             headers={"X-Auth-Token": token},
                             params={"all_tenants": 0}, timeout=25)
            r.raise_for_status()
            nova_servers = r.json().get("servers", [])
        except Exception as e:
            print(f"      [!] Nova error: {e}")
            continue

        for s in nova_servers:
            # IPs — floating first
            ip_pairs = []
            for addrs in s.get("addresses", {}).values():
                for a in addrs:
                    ip_pairs.append((a.get("OS-EXT-IPS:type", "fixed"), a.get("addr", "")))
            ip_pairs.sort(key=lambda x: x[0] != "floating")
            ips = ", ".join(ip for _, ip in ip_pairs if ip) or "—"

            # Flavor
            fl = s.get("flavor", {})
            fname  = fl.get("original_name") or fl.get("id", "")
            vcpu   = fl.get("vcpus", "")
            ram_mb = fl.get("ram", "")
            disk   = fl.get("disk", "")
            if vcpu and ram_mb:
                ram_gb = int(ram_mb) // 1024 if str(ram_mb).isdigit() else ram_mb
                config = f"{fname} ({vcpu} vCPU, {ram_gb} GB RAM, {disk} GB disk)"
            else:
                config = fname or "—"

            # OS
            image = s.get("image", {})
            os_name = (image.get("name") if isinstance(image, dict) else "") \
                      or s.get("metadata", {}).get("os_distro", "") or "—"

            az     = s.get("OS-EXT-AZ:availability_zone", "")
            region = f"{az} / {proj_name}" if az else proj_name

            all_servers.append({
                "name":          s.get("name") or s.get("id", "—"),
                "ip":            ips,
                "configuration": config,
                "region":        region,
                "os":            os_name,
                "billing_date":  "почасовой",
            })

    return all_servers


# ── Export ────────────────────────────────────────────────────────────────────

COLS      = ["name", "ip", "configuration", "region", "os", "billing_date"]
HEADS_RU  = ["Имя сервера", "IP-адрес", "Конфигурация", "Регион", "Операционная система", "Дата оплаты"]
COL_W     = [28, 22, 50, 28, 34, 15]


def _thin():
    s = Side(style="thin", color="B0B8C8")
    return Border(left=s, right=s, top=s, bottom=s)


def _write_sheet(ws, servers, title):
    ws.title = title
    n = len(COLS)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n)
    tc = ws.cell(row=1, column=1, value=title)
    tc.font      = Font(name="Arial", bold=True, size=13, color="FFFFFF")
    tc.fill      = PatternFill("solid", start_color="1B3A6B")
    tc.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    for ci, h in enumerate(HEADS_RU, 1):
        c = ws.cell(row=2, column=ci, value=h)
        c.font      = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        c.fill      = PatternFill("solid", start_color="2E75B6")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = _thin()
    ws.row_dimensions[2].height = 24

    for ri, srv in enumerate(servers, 3):
        bg = "EAF0FB" if ri % 2 == 0 else "FFFFFF"
        fill = PatternFill("solid", start_color=bg)
        for ci, key in enumerate(COLS, 1):
            c = ws.cell(row=ri, column=ci, value=srv.get(key, "—"))
            c.font      = Font(name="Arial", size=10)
            c.fill      = fill
            c.alignment = Alignment(vertical="center", wrap_text=(ci == 3))
            c.border    = _thin()
        ws.row_dimensions[ri].height = 18

    sc = ws.cell(row=len(servers) + 3, column=1, value=f"Итого: {len(servers)}")
    sc.font = Font(name="Arial", bold=True, italic=True, color="1B3A6B")

    for i, w in enumerate(COL_W, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A3"


def save_json(dedicated, cloud, path):
    obj = {
        "generated_at":      datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"),
        "dedicated_servers": dedicated,
        "cloud_servers":     cloud,
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)
    print(f"  ✓ JSON  → {path}")


def save_excel(dedicated, cloud, path):
    wb = Workbook()
    _write_sheet(wb.active, dedicated, "Выделенные серверы")
    _write_sheet(wb.create_sheet(), cloud, "Облачные серверы")
    wb.save(path)
    print(f"  ✓ Excel → {path}")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    args = parse_args()
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_json = f"/mnt/user-data/outputs/selectel_servers_{ts}.json"
    out_xlsx = f"/mnt/user-data/outputs/selectel_servers_{ts}.xlsx"

    print("=" * 60)
    print("  Selectel Server Inventory")
    print("=" * 60)

    print("\n[auth] Getting account IAM token...")
    try:
        acc_token = get_account_token(args.username, args.account, args.password)
        print("       ✓ OK (valid 24h)")
    except Exception as e:
        print(f"       ✗ Failed: {e}")
        sys.exit(1)

    print("\n[1/2] Dedicated servers...")
    dedicated = fetch_dedicated(acc_token)
    print(f"      Found: {len(dedicated)}")

    print("\n[2/2] Cloud (virtual) servers...")
    projects = []
    try:
        projects = list_projects(acc_token)
        print(f"      Projects: {len(projects)}")
    except Exception as e:
        print(f"      [!] Project list error: {e}")

    cloud = fetch_cloud(args.username, args.account, args.password, projects) if projects else []
    print(f"      Found: {len(cloud)}")

    print("\n[save] Writing files...")
    save_json(dedicated, cloud, out_json)
    save_excel(dedicated, cloud, out_xlsx)

    print(f"\n{'='*60}")
    print(f"  Done!")
    print(f"    JSON:  {out_json}")
    print(f"    Excel: {out_xlsx}")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
