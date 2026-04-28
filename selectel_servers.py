#!/usr/bin/env python3
"""
Selectel Server Inventory Script

pip install requests openpyxl

Запуск:
    export SEL_USERNAME="checks"
    export SEL_ACCOUNT="177345"
    export SEL_PASSWORD='ваш_пароль'
    python selectel_servers.py
"""

import os, sys, json, argparse, requests
from datetime import datetime, timezone
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

KEYSTONE_URL   = "https://cloud.api.selcloud.ru/identity/v3/auth/tokens"
DEDICATED_BASE = "https://api.selectel.ru/servers/v2"


# ── Auth ──────────────────────────────────────────────────────────────────────

def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument("--username", default=os.environ.get("SEL_USERNAME"))
    p.add_argument("--account",  default=os.environ.get("SEL_ACCOUNT"))
    p.add_argument("--password", default=os.environ.get("SEL_PASSWORD"))
    p.add_argument("--projects", default=os.environ.get("SEL_PROJECTS", ""),
                   help="Comma-separated project names if auto-discovery fails")
    args = p.parse_args()
    missing = [k for k, v in [("--username", args.username),
                               ("--account",  args.account),
                               ("--password", args.password)] if not v]
    if missing:
        print(f"ERROR: Missing: {', '.join(missing)}")
        sys.exit(1)
    return args


def get_account_token(username, account_id, password):
    payload = {"auth": {
        "identity": {"methods": ["password"], "password": {"user": {
            "name": username, "domain": {"name": account_id}, "password": password,
        }}},
        "scope": {"domain": {"name": account_id}},
    }}
    r = requests.post(KEYSTONE_URL, json=payload, timeout=20)
    r.raise_for_status()
    token = r.headers.get("X-Subject-Token")
    if not token:
        raise RuntimeError("No X-Subject-Token in response")
    return token


def get_project_token(username, account_id, password, project_name):
    payload = {"auth": {
        "identity": {"methods": ["password"], "password": {"user": {
            "name": username, "domain": {"name": account_id}, "password": password,
        }}},
        "scope": {"project": {"name": project_name, "domain": {"name": account_id}}},
    }}
    r = requests.post(KEYSTONE_URL, json=payload, timeout=20)
    r.raise_for_status()
    token   = r.headers.get("X-Subject-Token")
    body    = r.json().get("token", {})
    catalog = body.get("catalog", [])
    return token, catalog


def get_nova_url(catalog):
    for svc in catalog:
        if svc.get("type") == "compute":
            for ep in svc.get("endpoints", []):
                if ep.get("interface") == "public":
                    return ep.get("url", "").rstrip("/")
    return None


def list_projects(account_token):
    r = requests.get(
        "https://api.selectel.ru/vpc/resell/v2/projects",
        headers={"X-Auth-Token": account_token},
        timeout=15,
    )
    r.raise_for_status()
    raw = r.json()
    return raw if isinstance(raw, list) else raw.get("projects", [])


def fmt_date(raw):
    if not raw:
        return "—"
    try:
        dt = datetime.fromisoformat(str(raw).replace("Z", "+00:00"))
        return dt.strftime("%d.%m.%Y")
    except Exception:
        return str(raw)


# ── Dedicated servers ─────────────────────────────────────────────────────────

def fetch_dedicated(account_token):
    h = {"X-Auth-Token": account_token}
    servers = []

    # Locations map
    locations = {}
    try:
        r = requests.get(f"{DEDICATED_BASE}/location", headers=h, timeout=15)
        if r.ok:
            for loc in r.json().get("result", []):
                locations[loc.get("uuid", "")] = loc.get("name", "")
    except Exception:
        pass

    # Try all resource models that correspond to physical servers
    for model in ("server", "serverchip", "custom", None):
        params = {"limit": 1000}
        if model:
            params["model"] = model
        try:
            r = requests.get(f"{DEDICATED_BASE}/resource", headers=h,
                             params=params, timeout=30)
            if not r.ok:
                print(f"    [!] /resource?model={model} → {r.status_code}")
                continue
            resources = r.json().get("result", [])
            if not resources:
                continue
            print(f"    model={model}: {len(resources)} resource(s)")

            for res in resources:
                uuid  = res.get("uuid", "")
                rtype = res.get("model", model or "?")

                # Hardware details endpoint depends on model
                hw = {}
                for ep in (f"/resource/{uuid}/server",
                           f"/resource/{uuid}/serverchip",
                           f"/resource/{uuid}/custom"):
                    try:
                        rr = requests.get(f"{DEDICATED_BASE}{ep}", headers=h, timeout=10)
                        if rr.ok:
                            hw = rr.json().get("result", {})
                            if hw:
                                break
                    except Exception:
                        pass

                # OS
                os_name = "—"
                try:
                    rr = requests.get(f"{DEDICATED_BASE}/resource/{uuid}/os",
                                      headers=h, timeout=10)
                    if rr.ok:
                        d = rr.json().get("result", {})
                        os_name = d.get("name") or d.get("os_name") or "—"
                except Exception:
                    pass

                # IPs — try network management API
                ip_list = []
                try:
                    rr = requests.get(f"{DEDICATED_BASE}/network/ip",
                                      headers=h,
                                      params={"resource_uuid": uuid},
                                      timeout=10)
                    if rr.ok:
                        for item in rr.json().get("result", []):
                            addr = item.get("ip_address") or item.get("address", "")
                            if addr:
                                ip_list.append(addr)
                except Exception:
                    pass

                # Config string
                parts = []
                for key, label in [("cpu_info","CPU"),("cpu","CPU"),
                                    ("ram","RAM"),("memory","RAM"),
                                    ("disk_info","Disk"),("storage","Disk")]:
                    val = hw.get(key)
                    if val and label not in [p.split(":")[0] for p in parts]:
                        parts.append(f"{label}: {val}")
                config = " | ".join(parts) if parts else (
                    res.get("service_name") or res.get("name") or rtype)

                loc_uuid = res.get("location_uuid") or res.get("location", "")
                billing  = fmt_date(res.get("paid_till") or
                                    res.get("payment_date") or
                                    res.get("expires"))

                servers.append({
                    "name":          res.get("name") or res.get("title") or uuid,
                    "ip":            ", ".join(ip_list) or "—",
                    "configuration": config,
                    "region":        locations.get(loc_uuid, loc_uuid) or "—",
                    "os":            os_name,
                    "billing_date":  billing,
                })

            # If we got results with no model filter, don't repeat
            if model is None:
                break

        except Exception as e:
            print(f"    [!] model={model} error: {e}")

    # De-duplicate by name+ip
    seen = set()
    unique = []
    for s in servers:
        key = (s["name"], s["ip"])
        if key not in seen:
            seen.add(key)
            unique.append(s)
    return unique


# ── Cloud servers ─────────────────────────────────────────────────────────────

def fetch_cloud(username, account_id, password, project_names):
    all_servers = []

    for proj_name in project_names:
        print(f"    → project: {proj_name}")
        try:
            token, catalog = get_project_token(username, account_id, password, proj_name)
        except requests.HTTPError as e:
            print(f"      [!] Auth {e.response.status_code}: {e.response.text[:100]}")
            continue
        except Exception as e:
            print(f"      [!] Auth error: {e}")
            continue

        nova_url = get_nova_url(catalog)
        if not nova_url:
            print(f"      [!] No Nova endpoint in catalog")
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

        print(f"      Servers: {len(nova_servers)}")

        for s in nova_servers:
            ip_pairs = []
            for addrs in s.get("addresses", {}).values():
                for a in addrs:
                    ip_pairs.append((a.get("OS-EXT-IPS:type","fixed"), a.get("addr","")))
            ip_pairs.sort(key=lambda x: x[0] != "floating")
            ips = ", ".join(ip for _, ip in ip_pairs if ip) or "—"

            fl    = s.get("flavor", {})
            fname = fl.get("original_name") or fl.get("id", "")
            vcpu  = fl.get("vcpus", "")
            ram   = fl.get("ram", "")
            disk  = fl.get("disk", "")
            if vcpu and ram:
                ram_gb = int(ram) // 1024 if str(ram).isdigit() else ram
                config = f"{fname} ({vcpu} vCPU, {ram_gb} GB RAM, {disk} GB disk)"
            else:
                config = fname or "—"

            image   = s.get("image", {})
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

COLS     = ["name", "ip", "configuration", "region", "os", "billing_date"]
HEADS_RU = ["Имя сервера", "IP-адрес", "Конфигурация", "Регион",
            "Операционная система", "Дата оплаты"]
COL_W    = [28, 22, 50, 28, 34, 15]


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
        bg   = "EAF0FB" if ri % 2 == 0 else "FFFFFF"
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
        "generated_at":      datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "dedicated_servers": dedicated,
        "cloud_servers":     cloud,
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)
    print(f"  ✓ JSON  → {path}")


def save_excel(dedicated, cloud, path):
    wb = Workbook()
    _write_sheet(wb.active,        dedicated, "Выделенные серверы")
    _write_sheet(wb.create_sheet(), cloud,    "Облачные серверы")
    wb.save(path)
    print(f"  ✓ Excel → {path}")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    args = parse_args()
    ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_json = f"selectel_servers_{ts}.json"
    out_xlsx = f"selectel_servers_{ts}.xlsx"

    print("=" * 60)
    print("  Selectel Server Inventory")
    print("=" * 60)

    # Account token
    print("\n[auth] Getting account IAM token...")
    try:
        acc_token = get_account_token(args.username, args.account, args.password)
        print("       ✓ OK (valid 24h)")
    except Exception as e:
        print(f"       ✗ Failed: {e}")
        sys.exit(1)

    # Projects
    print("\n[projects] Discovering projects...")
    project_names = []
    try:
        projects      = list_projects(acc_token)
        project_names = [p.get("name","") for p in projects if p.get("name")]
        print(f"           Found: {project_names}")
    except Exception as e:
        print(f"           [!] {e}")
        if args.projects:
            project_names = [n.strip() for n in args.projects.split(",") if n.strip()]
            print(f"           Using --projects: {project_names}")

    # Dedicated servers
    print("\n[1/2] Dedicated servers...")
    dedicated = fetch_dedicated(acc_token)
    print(f"      Found: {len(dedicated)}")

    # Cloud servers
    print("\n[2/2] Cloud (virtual) servers...")
    cloud = fetch_cloud(args.username, args.account, args.password, project_names)
    print(f"      Found: {len(cloud)}")

    # Save
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
