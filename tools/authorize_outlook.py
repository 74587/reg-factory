"""Batch-authorize Outlook accounts and persist Microsoft Graph refresh tokens.

Input is one account per line: ``email----password``.  This command performs
only the Microsoft OAuth authorization flow; it never opens the Outlook
unlock/recovery flow or attempts a press-and-hold challenge.
"""

from __future__ import annotations

import argparse
import concurrent.futures
import os
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


def load_accounts(path: str) -> list[tuple[str, str]]:
    accounts: list[tuple[str, str]] = []
    seen: set[str] = set()
    for number, raw in enumerate(Path(path).expanduser().read_text(encoding="utf-8-sig").splitlines(), 1):
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        parts = re.split(r"-{4,}", line, maxsplit=1)
        if len(parts) < 2 or not parts[0].strip() or not parts[1].strip():
            raise ValueError(f"line {number}: expected email----password")
        email = parts[0].strip().lower()
        if email in seen:
            continue
        seen.add(email)
        accounts.append((email, parts[1].strip()))
    return accounts


def authorize_one(item: tuple[str, str], index: int) -> dict | None:
    email, password = item
    from tools.extract_graph_tokens import get_graph_token

    print(f"[authorize {index}] {email}: starting Graph OAuth", flush=True)
    result = get_graph_token(email, password, index)
    if not result or not result.get("refresh_token"):
        print(f"[authorize {index}] {email}: failed", flush=True)
        return None
    print(f"[authorize {index}] {email}: authorized", flush=True)
    return {
        "email": email,
        "password": password,
        "refresh_token": result["refresh_token"],
        "client_id": result.get("client_id") or "",
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Batch authorize Outlook accounts for Microsoft Graph")
    parser.add_argument("--input", "-i", required=True, help="one email----password per line")
    parser.add_argument("--concurrency", "-c", type=int, default=3)
    parser.add_argument("--no-update-pool", action="store_true")
    args = parser.parse_args()
    if not 1 <= args.concurrency <= 10:
        parser.error("--concurrency must be between 1 and 10")
    try:
        accounts = load_accounts(args.input)
    except (OSError, ValueError) as exc:
        print(f"[authorize] input error: {exc}", file=sys.stderr)
        return 2
    if not accounts:
        print("[authorize] no accounts")
        return 0

    with concurrent.futures.ThreadPoolExecutor(max_workers=min(args.concurrency, len(accounts))) as pool:
        futures = [pool.submit(authorize_one, account, index) for index, account in enumerate(accounts, 1)]
        results = [result for future in futures if (result := future.result())]

    if results and not args.no_update_pool:
        from common.outlook_recovery import upsert_refresh_tokens

        update = upsert_refresh_tokens(results)
        print(
            f"[authorize] pool updated: updated={update['updated']} appended={update['appended']} "
            f"errors_cleared={update['errors_cleared']}",
            flush=True,
        )
    print(f"[authorize] complete: {len(results)}/{len(accounts)} authorized")
    return 0 if len(results) == len(accounts) else 1


if __name__ == "__main__":
    from common import proxy_switch

    proxy_switch.apply_platform_environment("outlook")
    raise SystemExit(main())
