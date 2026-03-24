#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import math
import re
import sys
from calendar import month_name
from pathlib import Path
from typing import Any, Iterable

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
    parser.add_argument("--latitude", required=True, type=float, help="Latitude in decimal degrees.")
    parser.add_argument("--longitude", required=True, type=float, help="Longitude in decimal degrees.")
    return parser.parse_args()


def validate_coordinates(latitude: float, longitude: float) -> None:
    if not -90 <= latitude <= 90:
        raise ValueError(f"Latitude must be between -90 and 90. Received: {latitude}")
    if not -180 <= longitude <= 180:
        raise ValueError(f"Longitude must be between -180 and 180. Received: {longitude}")


def build_output_slug(latitude: float, longitude: float) -> str:
    def normalize(value: float, prefix_pos: str, prefix_neg: str) -> str:
        prefix = prefix_pos if value >= 0 else prefix_neg
        safe = f"{abs(value):.4f}".replace(".", "_")
        return f"{prefix}{safe}"

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


def iter_nested_values(node: Any) -> Iterable[Any]:
    if isinstance(node, dict):
        for value in node.values():
            yield value
            yield from iter_nested_values(value)
    elif isinstance(node, list):
        for item in node:
            yield item
            yield from iter_nested_values(item)


def find_monthly_records(payload: dict[str, Any]) -> list[dict[str, Any]]:
    candidates: list[list[dict[str, Any]]] = []

    direct_paths = [
        payload.get("outputs", {}).get("monthly"),
        payload.get("outputs", {}).get("monthly_data"),
        payload.get("monthly"),
    ]

    for candidate in direct_paths:
        if isinstance(candidate, list) and all(isinstance(x, dict) for x in candidate):
            candidates.append(candidate)

    for value in iter_nested_values(payload):
        if isinstance(value, list) and value and all(isinstance(x, dict) for x in value):
            first = value[0]
            if "month" in first or "Month" in first:
                candidates.append(value)

    if not candidates:
        raise RuntimeError("Unable to locate monthly records in PVGIS JSON response.")

    def score(records: list[dict[str, Any]]) -> tuple[int, int]:
        first = records[0]
        keys = set(first.keys())
        expected = {"month", "Month", "H(h)_m", "Hb(n)_m", "H(i_opt)_m", "Kd", "T2m"}
        return (len(keys & expected), len(records))

    best = sorted(candidates, key=score, reverse=True)[0]
    return best


def first_number(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if math.isnan(value):
            return None
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", ".")
        if text == "":
            return None
        try:
            return float(text)
        except ValueError:
            match = re.search(r"-?\d+(?:\.\d+)?", text)
            if match:
                return float(match.group(0))
    return None


def get_first_present(record: dict[str, Any], keys: list[str]) -> Any:
    for key in keys:
        if key in record:
            return record[key]
    return None


def normalize_records(records: list[dict[str, Any]], latitude: float, longitude: float) -> pd.DataFrame:
    normalized_rows: list[dict[str, Any]] = []

    for record in records:
        month_value = get_first_present(record, ["month", "Month"])
        month_number = int(first_number(month_value) or 0)
        if not 1 <= month_number <= 12:
            continue

        row = {
            "month_number": month_number,
            "month_name": month_name[month_number],
            "latitude": round(latitude, 6),
            "longitude": round(longitude, 6),
            "solar_radiation_database": RADIATION_DATABASE,
            "start_year": START_YEAR,
            "end_year": END_YEAR,
            "global_horizontal_irradiation_kwh_m2_month": first_number(
                get_first_present(record, ["H(h)_m", "Hh", "global_horizontal_irradiation"])
            ),
            "direct_normal_irradiation_kwh_m2_month": first_number(
                get_first_present(record, ["Hb(n)_m", "DNI", "direct_normal_irradiation"])
            ),
            "global_irradiation_optimum_angle_kwh_m2_month": first_number(
                get_first_present(record, ["H(i_opt)_m", "Hi_opt", "global_irradiation_optimum_angle"])
            ),
            "diffuse_to_global_ratio": first_number(
                get_first_present(record, ["Kd", "d2g", "diffuse_to_global_ratio"])
            ),
            "average_temperature_c": first_number(
                get_first_present(record, ["T2m", "temperature", "average_temperature"])
            ),
        }
        normalized_rows.append(row)

    if len(normalized_rows) != 12:
        raise RuntimeError(
            f"Expected 12 monthly rows after normalization, got {len(normalized_rows)}."
        )

    df = pd.DataFrame(normalized_rows)
    df = df.sort_values("month_number").reset_index(drop=True)
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
    destination.write_text(json.dumps(records, indent=2, ensure_ascii=False), encoding="utf-8")


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
            max_len = max(
                len(str(column_name)),
                *(len(str(value)) for value in df.iloc[:, idx - 1].tolist())
            )
            worksheet.column_dimensions[get_column_letter(idx)].width = min(max_len + 2, 40)

        workbook.save(destination)


def write_latest_manifest(output_dir: Path, csv_path: Path, json_path: Path, xlsx_path: Path) -> None:
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
        write_latest_manifest(output_dir, csv_path, json_path, xlsx_path)

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
