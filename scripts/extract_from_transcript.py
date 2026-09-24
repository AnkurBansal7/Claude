#!/usr/bin/env python3
"""
Extract a Gmail attachment from a get_message(messageFormat=RAW) tool result
recorded in a Claude Code agent transcript (.jsonl), instead of from a
manually-retyped copy of that tool result.

Why this exists: when a RAW email is small enough that the harness returns
it inline (rather than auto-saving it to a tool-result file), an agent that
needs it in a file has no way to get it there except by reproducing the
tool's ~30-50KB+ base64 output as literal text in a Write/Edit call. That
reproduction is not reliable (verified directly: three independent attempts
at retyping one such blob all silently truncated or corrupted it around the
same length). The harness itself, however, already wrote the *exact* tool
result into the calling agent's own .jsonl transcript — so extracting the
`raw` field from that file with this script is a mechanical operation, not
a text-generation one, and is exact by construction.

Usage:
    python3 extract_from_transcript.py <transcript.jsonl> <out_dir> [--message-id ID]

Scans the transcript for tool_result entries whose content parses as JSON
with a `raw` field (i.e. a get_message/get_thread RAW response), decodes the
MIME message, and writes any attachments to <out_dir>. If more than one such
result is found, pass --message-id to disambiguate (matches the response's
`id` field); otherwise all matches are extracted.

Prints one JSON line per saved attachment: {"filename": ..., "path": ..., "bytes": N}
"""
import sys
import json
import base64
import email
import argparse
from email import policy
from pathlib import Path


def find_raw_results(jsonl_path):
    """Yield (message_id, raw_b64) for every tool_result carrying a RAW email payload."""
    with open(jsonl_path) as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                obj = json.loads(line)
            except json.JSONDecodeError:
                continue
            message = obj.get("message")
            if not isinstance(message, dict):
                continue
            content = message.get("content")
            if not isinstance(content, list):
                continue
            for item in content:
                if not isinstance(item, dict) or item.get("type") != "tool_result":
                    continue
                inner = item.get("content")
                text = inner if isinstance(inner, str) else (
                    inner[0]["text"] if isinstance(inner, list) and inner and "text" in inner[0] else None
                )
                if not text:
                    continue
                try:
                    data = json.loads(text)
                except (json.JSONDecodeError, TypeError):
                    continue
                if isinstance(data, dict) and "raw" in data and "id" in data:
                    yield data["id"], data["raw"]


def extract_attachments(raw_b64, out_dir):
    padded = raw_b64 + "=" * (-len(raw_b64) % 4)
    raw_bytes = base64.urlsafe_b64decode(padded)
    msg = email.message_from_bytes(raw_bytes, policy=policy.default)

    out_dir = Path(out_dir)
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
    parser = argparse.ArgumentParser()
    parser.add_argument("transcript")
    parser.add_argument("out_dir")
    parser.add_argument("--message-id", default=None)
    args = parser.parse_args()

    results = list(find_raw_results(args.transcript))
    if args.message_id:
        results = [(mid, raw) for mid, raw in results if mid == args.message_id]

    if not results:
        print(json.dumps({"error": "no RAW get_message tool_result found in transcript"}), file=sys.stderr)
        sys.exit(1)

    seen = set()
    for mid, raw in results:
        if mid in seen:
            continue
        seen.add(mid)
        for entry in extract_attachments(raw, args.out_dir):
            print(json.dumps(entry))
