#!/usr/bin/env python3
"""
Fetch Copilot agent package metadata from Microsoft Graph (beta).

This is test/playground code only. It has not been tested in a customer production environment.
Use in a non-customer-impact environment only and perform your own due diligence.
There is no warranty, and this code should not be treated as legally binding.

Supports client credentials (app-only) and device-code (delegated) auth via MSAL.

Example:
  python agent_registry_exporter.py --mode client-credentials \
    --tenant-id <TENANT> --client-id <APP_ID> --client-secret <SECRET>

See README.md for permissions and details.
"""
import argparse
import json
import os
import sys
from typing import List, Dict, Any

import msal
import requests


GRAPH_ENDPOINT = "https://graph.microsoft.com/beta/copilot/admin/catalog/packages"
DEFAULT_DELEGATED_SCOPES = ["CopilotPackages.Read.All"]


def acquire_token_client_credentials(tenant_id: str, client_id: str, client_secret: str) -> str:
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.ConfidentialClientApplication(
        client_id, authority=authority, client_credential=client_secret
    )
    token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in token:
        raise RuntimeError(f"Failed to acquire token: {token}")
    return token["access_token"]


def acquire_token_device_code(tenant_id: str, client_id: str, scopes: List[str]) -> str:
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.PublicClientApplication(client_id, authority=authority)
    flow = app.initiate_device_flow(scopes=scopes)
    if "user_code" not in flow:
        raise RuntimeError(f"Failed to start device flow: {flow}")
    print(flow["message"])
    token = app.acquire_token_by_device_flow(flow)
    if token is None or "access_token" not in token:
        guidance = (
            "Device-code requests require a public client application registration. "
            "Do not use a confidential client app registration with client secret for this mode. "
            "In Azure AD, enable public client/native flows or add a native redirect URI such as "
            "https://login.microsoftonline.com/common/oauth2/nativeclient. "
            "If you are using this script in device-code mode, remove --client-secret and ensure "
            "the app registration is configured for public client flows."
        )
        raise RuntimeError(f"Failed to acquire token via device flow: {token}\n\n{guidance}")
    return token["access_token"]


def fetch_all_packages(access_token: str, endpoint: str = GRAPH_ENDPOINT) -> List[Dict[str, Any]]:
    headers = {"Authorization": f"Bearer {access_token}", "Accept": "application/json"}
    items: List[Dict[str, Any]] = []
    url = endpoint
    print(f"Querying Graph API: {endpoint}")
    print(headers)
    while url:
        resp = requests.get(url, headers=headers)
        if resp.status_code >= 400:
            raise RuntimeError(f"Graph API error {resp.status_code}: {resp.text}")
        data = resp.json()
        value = data.get("value", [])
        items.extend(value)
        # follow OData nextLink for pagination
        url = data.get("@odata.nextLink")
    return items


def save_json(path: str, data: Any) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def save_csv(path: str, items: List[Dict[str, Any]]) -> None:
    import csv

    # Flatten top-level keys; nested values will be JSON-dumped
    keys = set()
    for it in items:
        keys.update(it.keys())
    keys = sorted(keys)
    with open(path, "w", newline='', encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(keys)
        for it in items:
            row = [
                (json.dumps(it.get(k), ensure_ascii=False) if isinstance(it.get(k), (dict, list)) else it.get(k))
                for k in keys
            ]
            writer.writerow(row)


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Export Copilot agent package metadata from Microsoft Graph")
    p.add_argument("--mode", choices=["client-credentials", "device-code"], default="device-code",
                   help="Authentication mode: delegated device-code by default, or client-credentials for app-only.")
    p.add_argument("--tenant-id", required=True)
    p.add_argument("--client-id", required=True)
    p.add_argument("--client-secret", help="Client secret for client credentials flow (required for client-credentials mode)")
    p.add_argument("--scopes", help="Comma-separated delegated scopes for device-code flow (default: CopilotPackages.Read.All)")
    p.add_argument("--out-json", default="agents_metadata.json", help="Output JSON file")
    p.add_argument("--out-csv", help="Optional CSV output file")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    if args.mode == "client-credentials":
        if not args.client_secret:
            print("--client-secret is required for client-credentials mode", file=sys.stderr)
            sys.exit(2)
        token = acquire_token_client_credentials(args.tenant_id, args.client_id, args.client_secret)
    else:
        if args.client_secret:
            print(
                "Warning: --client-secret is ignored in device-code mode. "
                "Use a public client app registration instead.",
                file=sys.stderr,
            )
        scopes = [s.strip() for s in args.scopes.split(",") if s.strip()] if args.scopes else DEFAULT_DELEGATED_SCOPES
        print(f"Using delegated scopes: {scopes}")
        token = acquire_token_device_code(args.tenant_id, args.client_id, scopes)

    print("Fetching packages from Graph...")
    items = fetch_all_packages(token)
    print(f"Fetched {len(items)} packages")
    save_json(args.out_json, items)
    print(f"Saved JSON: {args.out_json}")
    if args.out_csv:
        save_csv(args.out_csv, items)
        print(f"Saved CSV: {args.out_csv}")


if __name__ == "__main__":
    main()
