"""
recap_engine.models
===================
Domain models representing jewelry styles, metal specifications, stone details,
and presentation recap rows.
"""

from dataclasses import dataclass, field
from typing import Optional, Dict, Any


METAL_CELL_MAP: Dict[str, str] = {
    'G14Y': '14K\nYELLOW\nGOLD',
    'G14W': '14K\nWHITE\nGOLD',
    'G14W/Y': '14K\nWHITE/YELLOW\nGOLD',
    'G14WY': '14K\nWHITE/YELLOW\nGOLD',
    'G14W/DP': '14K\nWHITE / PINK\nGOLD',
    'G14WDP': '14K\nWHITE / PINK\nGOLD',
    'G14R': '14K\nROSE\nGOLD',
    'G10R': '10K\nROSE\nGOLD',
    'G10Y': '10K\nYELLOW\nGOLD',
    'G10W': '10K\nWHITE\nGOLD',
    'G18Y': '18K\nYELLOW\nGOLD',
    'G18W': '18K\nWHITE\nGOLD',
    'G18R': '18K\nROSE\nGOLD',
    'STGSIL': 'STERLING\nSILVER',
    'STG': 'STERLING\nSILVER',
    'SIL': 'STERLING\nSILVER',
    'SS': 'STERLING\nSILVER',
    '925': 'STERLING\nSILVER',
    'PLT': 'PLATINUM',
}

METAL_DESC_MAP: Dict[str, str] = {
    'G14Y': '14K Yellow Gold',
    'G14W': '14K White Gold',
    'G14W/Y': '14K White/Yellow Gold',
    'G14WY': '14K White/Yellow Gold',
    'G14W/DP': '14K White & Pink Gold',
    'G14WDP': '14K White & Pink Gold',
    'G14R': '14K Rose Gold',
    'G10R': '10K Rose Gold',
    'G10Y': '10K Yellow Gold',
    'G10W': '10K White Gold',
    'G18Y': '18K Yellow Gold',
    'G18W': '18K White Gold',
    'G18R': '18K Rose Gold',
    'STGSIL': 'Silver',
    'STG': 'Silver',
    'SIL': 'Sterling Silver',
    'SS': 'Silver',
    '925': 'Sterling Silver',
    'PLT': 'Platinum',
}


@dataclass
class MetalDetails:
    """Represents metal specifications and weights for a style item."""
    code: str = ""
    gold_weight: float = 0.0
    silver_weight: float = 0.0
    gold_alloy: str = ""
    total_value: float = 0.0

    @property
    def display_weight(self) -> Any:
        """Returns the primary metal weight (Gold if present, else Silver)."""
        if self.gold_weight > 0:
            return round(self.gold_weight, 2)
        if self.silver_weight > 0:
            return round(self.silver_weight, 2)
        return ""

    @property
    def cell_text(self) -> str:
        """Multi-line metal text for the Recap METAL column cell."""
        clean = self.code.strip().upper()
        return METAL_CELL_MAP.get(clean, self.code.strip())

    @property
    def description_text(self) -> str:
        """Single-line metal text for the Recap DESCRIPTION string."""
        clean = self.code.strip().upper()
        return METAL_DESC_MAP.get(clean, self.code.strip())


@dataclass
class StoneDetails:
    """Represents details for a stone component (Diamond 1, Diamond 2, or Gemstone)."""
    label: str = ""
    shape_size: str = ""
    range_code: str = ""
    dimension_mm: str = ""
    count: float = 0.0
    quality_code: str = ""
    carat_weight: float = 0.0
    rate: float = 0.0
    value: float = 0.0


@dataclass
class StyleItem:
    """
    Core domain model representing a single jewelry style extracted from a pricing sheet.
    """
    style_no: str
    sr_no: str = ""
    ref: str = ""
    size: str = ""
    metal: MetalDetails = field(default_factory=MetalDetails)
    dia1: StoneDetails = field(default_factory=StoneDetails)
    dia2: StoneDetails = field(default_factory=StoneDetails)
    gem: StoneDetails = field(default_factory=StoneDetails)

    memo_price: Optional[float] = None
    ny_sell_coop: Optional[float] = None
    box_price: Optional[float] = None

    is_bridal_set: bool = False
    components_info: str = ""
    character_name: str = ""
    sku: str = ""
    image_file: str = ""
    raw_dict: Dict[str, Any] = field(default_factory=dict)

    @property
    def total_carat_weight(self) -> float:
        """Sum of all diamond and gemstone carat weights."""
        return round(self.dia1.carat_weight + self.dia2.carat_weight + self.gem.carat_weight, 4)

    @property
    def total_diamond_weight(self) -> float:
        """Sum of diamond carat weights strictly (excludes gemstones)."""
        return round(self.dia1.carat_weight + self.dia2.carat_weight, 4)

    @property
    def has_diamonds(self) -> bool:
        """True if style item contains diamonds."""
        return (self.total_diamond_weight > 0 or
                (bool(self.dia1.quality_code) and self.dia1.quality_code.upper() not in ('NA', '', '-')))

    @property
    def effective_unit_price(self) -> Any:
        """
        Resolves unit price following business priority:
          1. Bridal Set: Box Set price if present
          2. Memo Price (from MEMO COST or WITH TARIFF MEMO COST)
          3. NY Selling Incl. Coop
        """
        if self.is_bridal_set and self.box_price is not None:
            return round(self.box_price, 2)
        if self.memo_price is not None:
            return round(self.memo_price, 2)
        if self.ny_sell_coop is not None:
            return round(self.ny_sell_coop, 2)
        return "TBD"


@dataclass
class RecapRow:
    """
    Represents a formatted row prepared for insertion into the Client Recap sheet.
    """
    sr: int
    style_no: str
    image: str = ""
    description: str = ""
    metal: str = ""
    metal_wt: Any = ""
    dia_qly: str = ""
    cttw: str = ""
    gem_info: str = ""
    unit_price: Any = ""
    comments: str = ""
