#!/usr/bin/env python3
"""Phase 3 — POST a workbook containing an empty input table at the fact grain.

Builds the input-table STRUCTURE (the part that IS scriptable). After this runs,
do the three UI steps from refs/input-tables.md: paste/upload the CSV, Publish,
and create the warehouse view.

Auth: reads ~/.sigma-migration/env (SIGMA_BASE_URL/CLIENT_ID/CLIENT_SECRET).

Usage:
    python build-input-table-wb.py --name "Forecast Entry" \
        --connection cb2f5180-… \
        --columns REGION:text,BRANCH:text,MONTH_DATE:datetime,CATEGORY_CODE:number,FORECAST_AMOUNT:number \
        [--folder <uuid>]
"""
import argparse
import json
import os
import subprocess
import urllib.request
import urllib.parse

SYSTEM_COLS = ["ID", "CREATED_AT", "CREATED_BY", "UPDATED_AT", "UPDATED_BY"]


def token():
    env = {}
    with open(os.path.expanduser("~/.sigma-migration/env")) as f:
        for line in f:
            line = line.strip().replace("export ", "")
            if "=" in line:
                k, v = line.split("=", 1)
                env[k] = v.strip().strip('"')
    data = urllib.parse.urlencode({
        "grant_type": "client_credentials",
        "client_id": env["SIGMA_CLIENT_ID"],
        "client_secret": env["SIGMA_CLIENT_SECRET"],
    }).encode()
    req = urllib.request.Request(env["SIGMA_BASE_URL"] + "/v2/auth/token", data=data,
                                 headers={"Content-Type": "application/x-www-form-urlencoded"})
    tok = json.load(urllib.request.urlopen(req))["access_token"]
    return env["SIGMA_BASE_URL"], tok


def home_folder(base, tok):
    req = urllib.request.Request(base + "/v2/whoami",
                                 headers={"Authorization": f"Bearer {tok}"})
    uid = json.load(urllib.request.urlopen(req))["userId"]
    req = urllib.request.Request(base + f"/v2/members/{uid}",
                                 headers={"Authorization": f"Bearer {tok}"})
    return json.load(urllib.request.urlopen(req)).get("homeFolderId")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--name", required=True)
    ap.add_argument("--connection", required=True, help="write-enabled connectionId")
    ap.add_argument("--columns", required=True, help="comma list of NAME:type")
    ap.add_argument("--folder", default=None)
    args = ap.parse_args()

    base, tok = token()
    folder = args.folder or home_folder(base, tok)

    cols = [{"id": "ID"}]  # row id first; system cols carry NO type
    for spec in args.columns.split(","):
        name, _, typ = spec.partition(":")
        cols.append({"id": name.strip(), "type": (typ or "text").strip()})
    for sc in SYSTEM_COLS[1:]:
        cols.append({"id": sc})

    spec = {
        "name": args.name,
        "description": "Excel→Sigma input table (empty). Paste/upload the seed CSV, "
                       "publish, then create a warehouse view.",
        "folderId": folder,
        "schemaVersion": 1,
        "pages": [{
            "id": "entryPage",
            "name": args.name,
            "elements": [{
                "id": "inputTable",
                "kind": "input-table",
                "name": args.name,
                "source": {"kind": "empty", "connectionId": args.connection},
                "inputMode": "explore",
                "columns": cols,
            }],
        }],
    }

    body = json.dumps(spec).encode()
    req = urllib.request.Request(base + "/v2/workbooks/spec", data=body, method="POST",
                                 headers={"Authorization": f"Bearer {tok}",
                                          "Content-Type": "application/json",
                                          "Accept": "application/json"})
    resp = json.load(urllib.request.urlopen(req))
    wbid = resp.get("workbookId")
    print("workbookId:", wbid)
    if wbid:
        req = urllib.request.Request(base + f"/v2/workbooks/{wbid}",
                                     headers={"Authorization": f"Bearer {tok}"})
        print("url:", json.load(urllib.request.urlopen(req)).get("url"))
        print("\nNEXT (UI): open the input table → paste the seed CSV → PUBLISH → "
              "Warehouse views → Create new → copy the view path.")


if __name__ == "__main__":
    main()
