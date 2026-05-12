"""Standardize activity information from the input workbook."""

from __future__ import annotations

from pathlib import Path

from cncc_exa_scale.context import AppContext
from cncc_exa_scale.modules.excel_normalizer import standardize_activity_workbook


def run(context: AppContext) -> Path:
    app_config = context.config.get("app", {})
    flow_config = context.config.get("flows", {}).get("standardize_activity_info", {})

    input_dir = context.resolve_path(app_config.get("input_path", "input"))
    output_dir = context.resolve_path(app_config.get("output_dir", "output"))
    input_file = flow_config.get("input_file", "活动信息.xlsx")
    input_path = input_dir / input_file

    if not input_path.exists():
        candidates = [path for path in input_dir.glob("*.xlsx") if not path.name.startswith("~$")]
        if len(candidates) == 1:
            input_path = candidates[0]
        else:
            raise FileNotFoundError(f"Input workbook not found: {input_path}")

    context.logger.info("Standardizing workbook: %s", input_path)
    output_path = standardize_activity_workbook(input_path, output_dir)
    context.logger.info("Standardized workbook written: %s", output_path)
    return output_path
