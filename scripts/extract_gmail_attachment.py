#!/usr/bin/env python3
"""
Extract xlsx attachments from a Gmail message fetched in RAW format.

Input is the JSON file produced by the Gmail MCP tool's get_message call with
messageFormat=RAW (schema: {..., "raw": "<base64url RFC822 message>", ...}).
Gmail's get_message/get_thread tools do not expose attachment bytes directly
(FULL_CONTENT only returns attachment id/filename/mimeType), so RAW is the
only way to pull the actual file content through the MCP surface.

Usage:
    python3 extract_gmail_attachment.py <raw_message.json> <output_dir>

Prints one JSON line per saved attachment: {"filename": ..., "path": ..., "bytes": N}
"""
import sys
import json
import base64
import email
from email import policy
from pathlib import Path


def extract(raw_json_path: str, output_dir: str):
    with open(raw_json_path) as f:
        data = json.load(f)

    raw_b64 = data["raw"]
    # Gmail uses base64url; pad to a multiple of 4.
    padded = raw_b64 + "=" * (-len(raw_b64) % 4)
    raw_bytes = base64.urlsafe_b64decode(padded)
    msg = email.message_from_bytes(raw_bytes, policy=policy.default)

    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    saved = []
    for part in msg.walk():
        filename = part.get_filename()
        if not filename:
            continue
        payload = part.get_payload(decode=True)
        if payload is None:
            continue
        dest = out_dir / filename
        dest.write_bytes(payload)
        saved.append({"filename": filename, "path": str(dest), "bytes": len(payload)})

    return saved


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: extract_gmail_attachment.py <raw_message.json> <output_dir>", file=sys.stderr)
        sys.exit(1)

    results = extract(sys.argv[1], sys.argv[2])
    for r in results:
        print(json.dumps(r))
    if not results:
        print(json.dumps({"warning": "no attachments found"}))
