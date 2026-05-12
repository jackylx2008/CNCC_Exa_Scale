"""Run standardization, then calculate quarterly person-times."""

from __future__ import annotations

import shutil
from pathlib import Path

from cncc_exa_scale.context import AppContext
from cncc_exa_scale.flows.standardize_activity_info import run as run_standardization
from cncc_exa_scale.modules.person_time_stats import calculate_quarterly_person_times


def run(context: AppContext) -> Path:
    app_config = context.config.get("app", {})
    flow_config = context.config.get("flows", {}).get("quarterly_person_times", {})
    output_dir = context.resolve_path(app_config.get("output_dir", "output"))
    years = _parse_years(flow_config.get("years"))

    _clear_output_dir(output_dir, context.project_root)

    context.logger.info("Running prerequisite standardization workflow")
    standardized_path = run_standardization(context)

    context.logger.info("Calculating quarterly person-times")
    result = calculate_quarterly_person_times(
        standardized_path=standardized_path,
        output_dir=output_dir,
        years=years,
    )
    context.logger.info("Quarterly person-times written: %s", result.output_path)
    return result.output_path


def _parse_years(value: object) -> list[int] | None:
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    return [int(item.strip()) for item in text.split(",") if item.strip()]


def _clear_output_dir(output_dir: Path, project_root: Path) -> None:
    resolved_output = output_dir.resolve()
    resolved_project = project_root.resolve()
    if resolved_output == resolved_project:
        raise RuntimeError("拒绝清空项目根目录，请检查 app.output_dir 配置")

    if output_dir.exists() and not output_dir.is_dir():
        raise RuntimeError(f"输出路径不是目录，无法清空: {output_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)
    for child in output_dir.iterdir():
        try:
            if child.is_dir():
                shutil.rmtree(child)
            else:
                child.unlink()
        except OSError as exc:
            raise RuntimeError(
                f"无法清空 output 目录，请关闭被占用的文件后重试: {child}"
            ) from exc
