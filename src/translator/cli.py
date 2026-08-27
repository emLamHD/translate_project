from __future__ import annotations

import argparse
import csv
import json
import sys
from datetime import UTC, datetime
from pathlib import Path

from translator.config import PipelineConfig
from translator.errors import ProvenanceError, TokenProtectionError, TranslatorError
from translator.models import TMEntry
from translator.pipeline.orchestrator import run_pipeline
from translator.qa.reports import write_private_report
from translator.translation.memory import normalize_source, source_hash
from translator.translation.protect import validate_tokens


def import_tm(args: argparse.Namespace) -> int:
    source_path = Path(args.input)
    destination = Path(args.output)
    entries: dict[str, TMEntry] = {}
    conflicts: list[str] = []
    with source_path.open(encoding="utf-8-sig", newline="") as stream:
        reader = csv.DictReader(stream)
        required = {"unit_id", "source_text", "target_text", "approved_by"}
        if not reader.fieldnames or not required.issubset(reader.fieldnames):
            raise ProvenanceError(f"CSV requires columns: {sorted(required)}")
        for row in reader:
            source = normalize_source(row["source_text"])
            target = row["target_text"].strip()
            approver = row["approved_by"].strip() or args.approved_by
            if not source or not target or not approver:
                raise ProvenanceError(f"Incomplete approved row: {row.get('unit_id')}")
            if source == target:
                raise ProvenanceError(f"Source equals target: {row['unit_id']}")
            validate_tokens(source, target)
            digest = source_hash(source)
            entry = TMEntry(
                args.source,
                args.target,
                source,
                target,
                digest,
                "human_approved",
                row["unit_id"],
                approver,
                datetime.now(UTC).isoformat(),
                True,
            )
            previous = entries.get(digest)
            if previous and previous.target_text != target:
                conflicts.append(row["unit_id"])
            entries[digest] = entry
    if conflicts:
        raise ProvenanceError(f"Conflicting duplicate source rows: {conflicts}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(
        json.dumps(
            {"schema_version": 1, "entries": [item.to_dict() for item in entries.values()]},
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    audit = {
        "status": "PASS",
        "imported": len(entries),
        "conflicts": 0,
        "approved_by": args.approved_by,
    }
    write_private_report(Path(args.audit_report), audit)
    print(json.dumps(audit, indent=2))
    return 0


def migrate_legacy(args: argparse.Namespace) -> int:
    legacy = json.loads(Path(args.input).read_text(encoding="utf-8"))
    entries = []
    for index, (source, target) in enumerate(legacy.items()):
        normalized = normalize_source(source)
        entries.append(
            TMEntry(
                args.source,
                args.target,
                normalized,
                target,
                source_hash(normalized),
                args.provenance,
                f"legacy:{index}",
                None,
                None,
                False,
            ).to_dict()
        )
    Path(args.output).write_text(
        json.dumps({"schema_version": 1, "entries": entries}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(
        json.dumps(
            {
                "status": "MIGRATED_UNAPPROVED",
                "entries": len(entries),
                "provenance": args.provenance,
            }
        )
    )
    return 0


def translate(args: argparse.Namespace) -> int:
    config = PipelineConfig(
        args.execution_profile,
        args.source,
        args.target,
        args.format_profile,
        Path(args.translation_memory) if args.translation_memory else None,
        Path(args.output_dir),
        args.show_missing_markers,
        Path(args.format_config) if args.format_config else None,
    )
    output, report, payload = run_pipeline(Path(args.input), config)
    print(
        json.dumps(
            {"output": str(output), "report": str(report), "status": payload["status"]}, indent=2
        )
    )
    return 0


def parser() -> argparse.ArgumentParser:
    root = argparse.ArgumentParser(prog="translator")
    commands = root.add_subparsers(dest="command", required=True)
    run = commands.add_parser("translate")
    run.add_argument("input")
    run.add_argument("--execution-profile", choices=["no-ai"], required=True)
    run.add_argument("--source", required=True)
    run.add_argument("--target", required=True)
    run.add_argument(
        "--format-profile", choices=["preserve", "clean", "etech-sop"], default="preserve"
    )
    run.add_argument("--format-config")
    run.add_argument("--translation-memory")
    run.add_argument("--output-dir", required=True)
    run.add_argument("--show-missing-markers", action="store_true")
    run.set_defaults(handler=translate)
    imp = commands.add_parser("tm-import")
    imp.add_argument("--input", required=True)
    imp.add_argument("--source", required=True)
    imp.add_argument("--target", required=True)
    imp.add_argument("--approved-by", required=True)
    imp.add_argument("--output", required=True)
    imp.add_argument("--audit-report", required=True)
    imp.set_defaults(handler=import_tm)
    migrate = commands.add_parser("tm-migrate-legacy")
    migrate.add_argument("--input", required=True)
    migrate.add_argument("--output", required=True)
    migrate.add_argument("--source", required=True)
    migrate.add_argument("--target", required=True)
    migrate.add_argument(
        "--provenance",
        choices=["claude_silver_reference", "google_machine_draft", "unknown"],
        required=True,
    )
    migrate.set_defaults(handler=migrate_legacy)
    return root


def main() -> None:
    args = parser().parse_args()
    try:
        raise SystemExit(args.handler(args))
    except (TranslatorError, TokenProtectionError, ValueError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        raise SystemExit(2) from error


if __name__ == "__main__":
    main()
