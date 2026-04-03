from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Any, Dict, Optional

from utils.helpers import clean_text, parse_amount, parse_date


@dataclass
class MappedRow:
    business_date: Optional[datetime]
    supplier_name: str
    summary: str
    credit_amount: Optional[float]
    debit_amount: Optional[float]


class MappingService:
    """Map bank statement fields to journal template fields."""

    @staticmethod
    def map_record(record: Dict[str, Any]) -> MappedRow:
        business_date = parse_date(record.get("交易日期"))

        supplier_name = clean_text(record.get("对方户名"))

        remark = clean_text(record.get("备注(用途)"))
        summary_src = remark if remark else clean_text(record.get("摘要"))

        # 借贷方向对调
        credit_amount = parse_amount(record.get("借方金额"))
        debit_amount = parse_amount(record.get("贷方金额"))

        return MappedRow(
            business_date=business_date,
            supplier_name=supplier_name,
            summary=summary_src,
            credit_amount=credit_amount,
            debit_amount=debit_amount,
        )
