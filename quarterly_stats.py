from __future__ import annotations

import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))


def main() -> int:
    from cncc_exa_scale.config_loader import load_config
    from cncc_exa_scale.context import AppContext
    from cncc_exa_scale.flows.quarterly_person_times import run
    from cncc_exa_scale.logging_config import get_logger, setup_logger

    config = load_config(PROJECT_ROOT / "config.yaml")
    app_config = config.get("app", {})
    setup_logger(
        log_level=app_config.get("log_level", "INFO"),
        log_dir=PROJECT_ROOT / "log",
        log_name="quarterly_stats",
    )
    logger = get_logger(__name__)
    context = AppContext(project_root=PROJECT_ROOT, config=config, logger=logger)

    try:
        output_path = run(context)
    except Exception:
        logger.exception("Quarterly person-time statistics failed")
        return 1

    print(output_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
