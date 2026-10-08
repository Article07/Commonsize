#!/usr/bin/env python
"""
Commonsize automation -- no-API prototype CLI.

Examples:
    python main.py --input sample_input/sample_line_items.xlsx \\
        --company "DOOALL CORPRO (INDIA) PVT LTD" --output output/DOOALL_commonsize.xlsx

    python main.py --input annual_report.pdf --company "Ficus Pax" --extract-only
"""

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from core import analysis, classifier, computation, excel_writer, io_loader, irl, template_writer


def parse_args():
    parser = argparse.ArgumentParser(description="Commonsize automation (rule-based prototype)")
    parser.add_argument("--input", required=True, help="Path to a structured .xlsx/.csv or a PDF annual report")
    parser.add_argument("--company", required=True, help="Company name, used in sheet titles")
    parser.add_argument("--output", help="Output workbook path (required unless --extract-only)")
    parser.add_argument(
        "--extract-only",
        action="store_true",
        help="For PDF input: only extract tables for manual review, skip the full pipeline",
    )
    parser.add_argument(
        "--use-template",
        action="store_true",
        help="Populate the real Reference file template's P&L schedules instead of "
        "writing a standalone workbook (requires --template and Microsoft Excel installed)",
    )
    parser.add_argument(
        "--template",
        help="Path to the Reference file template (required with --use-template)",
    )
    parser.add_argument(
        "--equity-shares",
        type=int,
        default=None,
        help="Number of equity shares issued, used to rewrite the Sch. sheet's "
        "'\"x\" equity shares of Rs. 10 each' placeholder label (--use-template only)",
    )
    return parser.parse_args()


def run_pdf_extraction(input_path, output_dir):
    review_path, table_count = io_loader.extract_pdf_tables_for_review(input_path, output_dir)
    print(f"Extracted {table_count} table(s) from {input_path.name}.")
    print(f"Review workbook written to: {review_path}")
    print(
        "Next step: convert the relevant rows into the structured Statement/Line Item/"
        "Statement Tag/FY.. schema, then re-run with --input pointing at that file."
    )


def run_pipeline(
    input_path,
    company_name,
    output_path,
    use_template=False,
    template_path=None,
    equity_shares=None,
):
    line_items = io_loader.load_structured(input_path)
    if not line_items:
        print("No line items found in the input file.", file=sys.stderr)
        sys.exit(1)

    classifier.classify_all(line_items)

    fy_columns = computation.ordered_fy_columns(line_items)
    computation.compute_common_size(line_items, fy_columns)
    computation.compute_yoy_change(line_items, fy_columns)

    metrics, sections = analysis.run_analysis(line_items, fy_columns)
    irl_rows = irl.build_irl(sections)

    output_path.parent.mkdir(parents=True, exist_ok=True)

    if use_template:
        tw_log, unmapped_items = template_writer.populate_template(
            template_path,
            output_path,
            company_name,
            line_items,
            fy_columns,
            sections=sections,
            irl_rows=irl_rows,
            equity_shares=equity_shares,
        )
        print("\n".join(tw_log))
        if unmapped_items:
            print(
                f"{len(unmapped_items)} item(s) could not be routed to a template schedule "
                "(Uncertain classification, an untagged BS row, or a tag with no Sch. "
                "destination in this phase, e.g. CF tags) -- left out of Sch., see the "
                "Analysis/IRL sheet:"
            )
            for li in unmapped_items:
                print(f"  - {li.line_item}")
    else:
        excel_writer.write_workbook(output_path, company_name, line_items, fy_columns, sections, irl_rows)

    uncertain_count = sum(1 for li in line_items if li.confidence.startswith("Uncertain"))
    print(f"Loaded {len(line_items)} line item(s) across FY columns: {', '.join(fy_columns)}")
    print(f"{uncertain_count} item(s) flagged Uncertain -- Needs Review.")
    print(f"{len(irl_rows)} item(s) added to the Information Requirement List.")
    print(f"Workbook written to: {output_path}")


def main():
    args = parse_args()
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"Input file not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    is_pdf = input_path.suffix.lower() == ".pdf"

    if is_pdf:
        output_dir = Path(args.output).parent if args.output else Path("output")
        run_pdf_extraction(input_path, output_dir)
        if not args.extract_only:
            print(
                "Note: PDF input only supports table extraction for review in this "
                "no-API prototype. Re-run with the converted structured file to run "
                "the full classification/analysis pipeline."
            )
        return

    if args.extract_only:
        print("--extract-only only applies to PDF input.", file=sys.stderr)
        sys.exit(1)

    if not args.output:
        print("--output is required for structured input.", file=sys.stderr)
        sys.exit(1)

    if args.use_template and not args.template:
        print("--template is required when --use-template is set.", file=sys.stderr)
        sys.exit(1)

    template_path = Path(args.template) if args.template else None
    if args.use_template and not template_path.exists():
        print(f"Template file not found: {template_path}", file=sys.stderr)
        sys.exit(1)

    run_pipeline(
        input_path,
        args.company,
        Path(args.output),
        use_template=args.use_template,
        template_path=template_path,
        equity_shares=args.equity_shares,
    )


if __name__ == "__main__":
    main()
