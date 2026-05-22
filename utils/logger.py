import json
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path


def setup_bonus_logger(log_file: str = "logs/bonus_calc.log"):
    logger = logging.getLogger("bonus_calc")
    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)
    log_path = Path(log_file)
    log_path.parent.mkdir(parents=True, exist_ok=True)

    handler = RotatingFileHandler(log_path, maxBytes=10 * 1024 * 1024, backupCount=5, encoding="utf-8")
    formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    logger.propagate = False
    return logger


def audit_event(logger: logging.Logger, event: str, **payload):
    logger.info(json.dumps({"event": event, **payload}, ensure_ascii=False, default=str))
