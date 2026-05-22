import logging
from typing import Dict

from pydantic import BaseModel, ConfigDict, Field, field_validator

from config_loader import load_app_config


LOGGER = logging.getLogger("bonus_calc")
CONFIG = load_app_config()


class KPIMetricInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)

    metric_name: str
    plan: float = Field(ge=0)
    fact: float = Field(ge=0)

    @field_validator("fact")
    @classmethod
    def cap_fact(cls, value: float, info):
        plan = info.data.get("plan", 0)
        cap_ratio = CONFIG["thresholds"]["fact_cap_ratio"]
        if plan and value > plan * cap_ratio:
            LOGGER.warning(
                "Fact %.2f exceeds %.0f%% of plan %.2f for metric '%s'",
                value,
                cap_ratio * 100,
                plan,
                info.data.get("metric_name", ""),
            )
        return value


class BonusCalcRequest(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)

    dealer_center: str
    position: str
    month: str
    metrics: Dict[str, KPIMetricInput]
