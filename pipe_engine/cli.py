"""
cli.py — command-line entry for pipe_engine.

Usage:

  python -m pipe_engine.cli fill --extract <extract.json> --output <vor.xlsx> [...]

Closes the dwg→xlsx end-to-end pipeline:

  AutoCAD dwg
    → (cad-agent-mcp: CLAUDE_EXTRACT_ROUTE → extract_route.json)
    → (this CLI: read JSON → interpret.build_branches → derive_exit_sand
                 → xlsx_writer.write_xlsx)
    → vor.xlsx (formulas-on-output)

Engineer opens the xlsx, audits formulas, modifies any input cell,
Excel recalculates — same workflow as the original reference workbook.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

# Force UTF-8 stdout/stderr so cyrillic in --help and printout doesn't crash
# on PowerShell's default cp1251 code-page on Russian Windows.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from .earthwork import PitParams, ProjectParams
from .interpret import build_branches, derive_exit_sand_from_extract
from .io_schema import ExtractDocument
from .trench_profile import Section, SoilType
from .xlsx_writer import XlsxRenderInputs, write_xlsx


# ---- Subcommand: fill -------------------------------------------------------

def _add_fill_parser(subparsers: argparse._SubParsersAction) -> None:
    p = subparsers.add_parser(
        "fill",
        help="Render an earthwork xlsx from a dwg extract JSON.",
        description="Read a JSON extract produced by cad-agent-mcp (CLAUDE_EXTRACT_ROUTE), "
                    "interpret it (branches + exit_sand), and write an xlsx where every "
                    "computed cell is an Excel formula.",
    )
    p.add_argument("--extract", required=True, type=Path,
                   help="Path to the JSON extract from cad-agent-mcp (CLAUDE_EXTRACT_ROUTE).")
    p.add_argument("--output", "-o", required=True, type=Path,
                   help="Path to the xlsx to write (overwrites if exists).")

    # ---- Project params ----
    g = p.add_argument_group("project parameters (ИсхДанные)")
    g.add_argument("--depth", type=float, default=1.8,
                   help="Глубина заложения газопровода, м (default 1.8).")
    g.add_argument("--sand-pad", type=float, default=0.1,
                   help="Толщина песчаной подушки, м (default 0.1).")
    g.add_argument("--snow", "--snow-thickness", type=float, default=0.5,
                   dest="snow_thickness",
                   help="Толщина снежного покрова, м (default 0.5).")

    # ---- Trench profile ----
    g = p.add_argument_group("trench profile (геометрия траншеи)")
    g.add_argument("--bottom-width", type=float, default=0.7,
                   help="Ширина траншеи понизу для одиночной прокладки, м (default 0.7).")
    g.add_argument("--joint-bottom-width", type=float, default=1.2,
                   help="Ширина траншеи понизу для совместной прокладки, м (default 1.2).")
    g.add_argument("--slope", "--slope-ratio", type=float, default=0.0,
                   dest="slope_ratio",
                   help="Уклон откоса 1:k (default 0 = вертикальные стенки). "
                        "Для песчаного грунта летом по СНиП 12-04-2002 ~0.5.")
    g.add_argument("--sand-cushion", action="store_true",
                   help="Включить песчаную подушку под трубой (10 см) + присыпку (20 см) "
                        "из ПРИВОЗНОГО песка. Default: False (используем местный грунт).")

    # ---- Section / branch defaults ----
    g = p.add_argument_group("section defaults (общие для всех веток)")
    g.add_argument("--soil", choices=[s.name.lower() for s in SoilType],
                   default=SoilType.SAND.name.lower(),
                   help="Тип грунта (используется для автоподбора откоса летом). "
                        "Default: sand.")
    g.add_argument("--season", choices=["winter", "summer"], default="winter",
                   help="Сезон работ. winter → вертикальные стенки (нормативно по "
                        "СП 45.13330 п.6.1.15, мёрзлый грунт держит) и расчёт снега. "
                        "summer → откос по типу грунта, снег=0. Default: winter.")
    g.add_argument("--laying-method", choices=["plet", "individual"], default="plet",
                   help="Способ укладки: 'plet' (плеть, сваренная) или 'individual' "
                        "(отдельные трубы). Default: plet.")

    # ---- Exit-sand inputs that don't come from dwg ----
    g = p.add_argument_group("exit-sand counts (что не извлекается из dwg автоматически)")
    g.add_argument("--hand-trench-count", type=int, default=0,
                   help="Кол-во пересечений инж.коммуникаций требующих ручной траншеи "
                        "в футляре. Default: 0.")
    g.add_argument("--gnb-pit-count", type=int, default=0,
                   help="Кол-во рабочих/приёмных котлованов ГНБ. Default: 0.")
    g.add_argument("--joint-kt-count", type=int, default=0,
                   help="Кол-во КТ при совместной прокладке. Default: 0.")
    g.add_argument("--working-pit-count", type=int, default=0,
                   help="Кол-во рабочих котлованов для ГНБ (для блока котлованов). "
                        "Default: 0.")

    # Manual-count templates (sophisticated nodes that aren't block-triggered)
    g = p.add_argument_group("manual-trigger templates (complex nodes count from CLI)")
    g.add_argument("--grpsh-count", type=int, default=0,
                   help="Кол-во ГРПШ/ГРП (Газорегуляторных пунктов). Default: 0.")
    g.add_argument("--zazemlitel-count", type=int, default=0,
                   help="Кол-во заземлителей. Default: 0.")
    g.add_argument("--vrezka-count", type=int, default=0,
                   help="Кол-во врезок в действующий газопровод. Default: 0.")
    g.add_argument("--svecha-count", type=int, default=0,
                   help="Кол-во вытяжных свечей. Default: 0.")
    g.add_argument("--uzel-b-count", type=int, default=0,
                   help="Кол-во узлов Б (двухкрановая схема вывода у потребителя). "
                        "Default: 0. Узел А detect'ится автоматом по полилиниям _Вводы; "
                        "Узел Б — manual, отличается проектным решением.")

    p.set_defaults(func=_cmd_fill)


def _cmd_fill(args: argparse.Namespace) -> int:
    if not args.extract.exists():
        print(f"ERROR: extract file not found: {args.extract}", file=sys.stderr)
        return 1

    # 1. Load extract
    try:
        doc = ExtractDocument.from_json_file(args.extract)
    except (json.JSONDecodeError, ValueError, KeyError) as e:
        print(f"ERROR: failed to parse extract JSON: {e}", file=sys.stderr)
        return 1

    print(f"Loaded extract: {len(doc.polylines)} polylines, {len(doc.blocks)} blocks")
    print(f"  Source dwg: {doc.source.dwg_path}")
    print(f"  Drawing units assumed: {doc.source.drawing_units_assumed}")

    # 2. Build section
    soil = SoilType[args.soil.upper()]
    section = Section(
        name=f"cli_{args.soil}_{args.season}",
        soil=soil,
        season=args.season,
        laying_method=args.laying_method,
        sand_cushion_required=args.sand_cushion,
    )

    # 3. Build branches
    branches = build_branches(doc, default_section=section)
    print(f"\nBranches built: {len(branches)}")
    for b in branches:
        dns = ", ".join(f"DN{p.dn_mm}={p.length_grp1:.1f}м" for p in b.pipes)
        print(f"  {b.id}: {dns}")

    # 4. Derive exit-sand nodes
    exit_sand = derive_exit_sand_from_extract(
        doc,
        hand_trench_count=args.hand_trench_count,
        gnb_pit_count=args.gnb_pit_count,
        joint_kt_count=args.joint_kt_count,
    )
    print(f"\nExit-sand nodes: {len(exit_sand)}")
    for n in exit_sand:
        print(f"  {n.name}: count={n.count} × kratность={n.sand_units}")

    # 5. Bundle inputs
    project_params = ProjectParams(
        depth=args.depth,
        snow_thickness=args.snow_thickness,
        sand_pad=args.sand_pad,
    )
    pits = PitParams(working_pit_count=args.working_pit_count) if args.working_pit_count > 0 else None

    manual_counts = {
        "grpsh":      args.grpsh_count,
        "zazemlitel": args.zazemlitel_count,
        "vrezka":     args.vrezka_count,
        "svecha":     args.svecha_count,
        "uzel_b":     args.uzel_b_count,
    }
    inputs = XlsxRenderInputs(
        project_params=project_params,
        branches=branches,
        exit_sand_nodes=exit_sand,
        pits=pits,
        hand_trench_count=args.hand_trench_count,
        bottom_width=args.bottom_width,
        slope_ratio=args.slope_ratio,
        sand_cushion_required=args.sand_cushion,
        season_winter=(args.season == "winter"),
        joint_bottom_width=args.joint_bottom_width,
        joint_pipe_pairs=[],   # CLI doesn't accept joint pairs yet — future flag
        blocks=doc.blocks,     # for ВР node-library auto-emit (block-triggered)
        polylines=doc.polylines,         # for polyline-triggered templates (Узел А)
        manual_counts=manual_counts,     # for manual-triggered templates
        source_dwg_path=doc.source.dwg_path,
        source_extracted_at_utc=doc.source.extracted_at_utc,
    )

    # 6. Write xlsx
    args.output.parent.mkdir(parents=True, exist_ok=True)
    try:
        write_xlsx(inputs, args.output)
    except PermissionError as e:
        print(f"\nERROR: cannot write {args.output} — file may be open in Excel: {e}",
              file=sys.stderr)
        return 1
    except Exception as e:
        print(f"\nERROR: write_xlsx failed: {e}", file=sys.stderr)
        return 1

    size_kb = args.output.stat().st_size / 1024
    print(f"\nWrote: {args.output} ({size_kb:.1f} KB)")
    return 0


# ---- Top-level entry --------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pipe-engine",
        description="pipe_engine — deterministic engineering engine for AutoCAD → ВОР pipeline. "
                    "Input: extract.v1.json from the CAD agent; output: xlsx with formulas.",
    )
    subparsers = parser.add_subparsers(dest="cmd", required=True)
    _add_fill_parser(subparsers)
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
