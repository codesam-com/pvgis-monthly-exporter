#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import math
import sys
from calendar import month_name
from pathlib import Path
from typing import Any

import pandas as pd
import requests
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

API_URL = "https://re.jrc.ec.europa.eu/api/v5_3/MRcalc"
START_YEAR = 2005
END_YEAR = 2023
RADIATION_DATABASE = "PVGIS-SARAH3"
REQUEST_TIMEOUT_SECONDS = 60

OUTPUT_COLUMNS = [
    "month_number",
    "month_name",
    "latitude",
    "longitude",
    "solar_radiation_database",
    "start_year",
    "end_year",
    "global_horizontal_irradiation_kwh_m2_month",
    "direct_normal_irradiation_kwh_m2_month",
    "global_irradiation_optimum_angle_kwh_m2_month",
    "diffuse_to_global_ratio",
    "average_temperature_c",
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Download monthly PVGIS data and export CSV/JSON/XLSX."
    )
    parser.add_argument(
        "--latitude",
        required=True,
        type=float,
        help="Latitude in decimal degrees.",
    )
    parser.add_argument(
        "--longitude",
        required=True,
        type=float,
        help="Longitude in decimal degrees.",
    )
    return parser.parse_args()


def validate_coordinates(latitude: float, longitude: float) -> None:
    if not -90 <= latitude <= 90:
        raise ValueError(f"Latitude must be between -90 and 90. Received: {latitude}")
    if not -180 <= longitude <= 180:
        raise ValueError(f"Longitude must be between -180 and 180. Received: {longitude}")


def build_output_slug(latitude: float, longitude: float) -> str:
    def normalize(value: float, positive_prefix: str, negative_prefix: str) -> str:
        prefix = positive_prefix if value >= 0 else negative_prefix
        safe_value = f"{abs(value):.4f}".replace(".", "_")
        return f"{prefix}{safe_value}"

    lat_slug = normalize(latitude, "lat", "latm")
    lon_slug = normalize(longitude, "lon", "lonm")
    return f"{lat_slug}__{lon_slug}"


def request_pvgis_monthly_data(latitude: float, longitude: float) -> dict[str, Any]:
    params = {
        "lat": latitude,
        "lon": longitude,
        "raddatabase": RADIATION_DATABASE,
        "startyear": START_YEAR,
        "endyear": END_YEAR,
        "horirrad": 1,
        "optrad": 1,
        "mr_dni": 1,
        "d2g": 1,
        "avtemp": 1,
        "outputformat": "json",
    }

    response = requests.get(API_URL, params=params, timeout=REQUEST_TIMEOUT_SECONDS)

    try:
        response.raise_for_status()
    except requests.HTTPError as exc:
        raise RuntimeError(
            f"PVGIS request failed with status {response.status_code}: {response.text}"
        ) from exc

    payload = response.json()
    if not isinstance(payload, dict):
        raise RuntimeError("Unexpected PVGIS response: root JSON object is not a dictionary.")

    return payload


def first_number(value: Any) -> float | None:
    if value is None:
        return None

    if isinstance(value, bool):
        return None

    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)

    if isinstance(value, str):
        text = value.strip().replace(",", ".")
        if not text:
            return None
        try:
            return float(text)
        except ValueError:
            return None

    return None


def get_first_present(record: dict[str, Any], keys: list[str]) -> Any:
    for key in keys:
        if key in record:
            return record[key]
    return None


def is_monthly_table(value: Any) -> bool:
    if not isinstance(value, list) or len(value) != 12:
        return False

    if not all(isinstance(item, dict) for item in value):
        return False

    months: list[int] = []
    for row in value:
        month_raw = row.get("month", row.get("Month"))
        month_num = first_number(month_raw)
        if month_num is None:
            return False
        months.append(int(month_num))

    if sorted(months) != list(range(1, 13)):
        return False

    keys = set().union(*(row.keys() for row in value))
    expected_keys = {"H(h)_m", "H(i_opt)_m", "Hb(n)_m", "Kd", "T2m"}

    return len(keys & expected_keys) >= 3


def walk_and_collect_monthly_tables(node: Any, candidates: list[list[dict[str, Any]]]) -> None:
    if is_monthly_table(node):
        candidates.append(node)
        return

    if isinstance(node, dict):
        for child in node.values():
            walk_and_collect_monthly_tables(child, candidates)
    elif isinstance(node, list):
        for child in node:
            walk_and_collect_monthly_tables(child, candidates)


def find_monthly_records(payload: dict[str, Any]) -> list[dict[str, Any]]:
    candidates: list[list[dict[str, Any]]] = []
    walk_and_collect_monthly_tables(payload, candidates)

    if not candidates:
        debug_path = Path("debug_pvgis_response.json")
        debug_path.write_text(
            json.dumps(payload, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        raise RuntimeError(
            "Unable to locate the 12-row monthly MRcalc table in PVGIS JSON response. "
            f"Full response saved to {debug_path}."
        )

    def score(records: list[dict[str, Any]]) -> int:
        keys = set().union(*(row.keys() for row in records))
        expected = {"month", "Month", "H(h)_m", "H(i_opt)_m", "Hb(n)_m", "Kd", "T2m"}
        return len(keys & expected)

    return sorted(candidates, key=score, reverse=True)[0]


def normalize_records(records: list[dict[str, Any]], latitude: float, longitude: float) -> pd.DataFrame:
    normalized_rows: list[dict[str, Any]] = []

    for record in records:
        month_value = get_first_present(record, ["month", "Month"])
        month_number_raw = first_number(month_value)
        if month_number_raw is None:
            continue

        month_number = int(month_number_raw)
        if not 1 <= month_number <= 12:
            continue

        normalized_rows.append(
            {
                "month_number": month_number,
                "month_name": month_name[month_number],
                "latitude": round(latitude, 6),
                "longitude": round(longitude, 6),
                "solar_radiation_database": RADIATION_DATABASE,
                "start_year": START_YEAR,
                "end_year": END_YEAR,
                "global_horizontal_irradiation_kwh_m2_month": first_number(
                    get_first_present(record, ["H(h)_m"])
                ),
                "direct_normal_irradiation_kwh_m2_month": first_number(
                    get_first_present(record, ["Hb(n)_m"])
                ),
                "global_irradiation_optimum_angle_kwh_m2_month": first_number(
                    get_first_present(record, ["H(i_opt)_m"])
                ),
                "diffuse_to_global_ratio": first_number(
                    get_first_present(record, ["Kd"])
                ),
                "average_temperature_c": first_number(
                    get_first_present(record, ["T2m"])
                ),
            }
        )

    if len(normalized_rows) != 12:
        raise RuntimeError(
            f"Expected 12 monthly rows after normalization, got {len(normalized_rows)}."
        )

    df = pd.DataFrame(normalized_rows)
    df = df.sort_values("month_number").drop_duplicates(subset=["month_number"], keep="first")
    df = df.reset_index(drop=True)

    if len(df) != 12:
        raise RuntimeError(
            f"Expected 12 unique months after deduplication, got {len(df)}."
        )

    df = df[OUTPUT_COLUMNS]

    float_columns = [
        "global_horizontal_irradiation_kwh_m2_month",
        "direct_normal_irradiation_kwh_m2_month",
        "global_irradiation_optimum_angle_kwh_m2_month",
        "diffuse_to_global_ratio",
        "average_temperature_c",
    ]
    for col in float_columns:
        df[col] = df[col].round(3)

    return df


def ensure_output_dir(latitude: float, longitude: float) -> Path:
    slug = build_output_slug(latitude, longitude)
    output_dir = Path("data") / slug
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir


def write_csv(df: pd.DataFrame, destination: Path) -> None:
    df.to_csv(destination, index=False, encoding="utf-8")


def write_json(df: pd.DataFrame, destination: Path) -> None:
    records = df.to_dict(orient="records")
    destination.write_text(
        json.dumps(records, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


def write_xlsx(df: pd.DataFrame, destination: Path) -> None:
    with pd.ExcelWriter(destination, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="monthly_data")

        workbook = writer.book
        worksheet = writer.sheets["monthly_data"]

        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions

        header_font = Font(bold=True)
        for cell in worksheet[1]:
            cell.font = header_font

        for idx, column_name in enumerate(df.columns, start=1):
            values = df.iloc[:, idx - 1].tolist()
            max_len = max(len(str(column_name)), *(len(str(value)) for value in values))
            worksheet.column_dimensions[get_column_letter(idx)].width = min(max_len + 2, 40)

        workbook.save(destination)


def write_manifest(output_dir: Path, csv_path: Path, json_path: Path, xlsx_path: Path) -> None:
    manifest = {
        "files": {
            "csv": csv_path.name,
            "json": json_path.name,
            "xlsx": xlsx_path.name,
        }
    }
    (output_dir / "manifest.json").write_text(
        json.dumps(manifest, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


def main() -> int:
    args = parse_args()

    try:
        validate_coordinates(args.latitude, args.longitude)

        payload = request_pvgis_monthly_data(args.latitude, args.longitude)
        monthly_records = find_monthly_records(payload)
        df = normalize_records(monthly_records, args.latitude, args.longitude)

        output_dir = ensure_output_dir(args.latitude, args.longitude)
        base_name = f"pvgis_monthly_{build_output_slug(args.latitude, args.longitude)}"

        csv_path = output_dir / f"{base_name}.csv"
        json_path = output_dir / f"{base_name}.json"
        xlsx_path = output_dir / f"{base_name}.xlsx"

        write_csv(df, csv_path)
        write_json(df, json_path)
        write_xlsx(df, xlsx_path)
        write_manifest(output_dir, csv_path, json_path, xlsx_path)

        print(f"Generated files in: {output_dir}")
        print(f" - {csv_path.name}")
        print(f" - {json_path.name}")
        print(f" - {xlsx_path.name}")
        print("Done.")
        return 0

    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
