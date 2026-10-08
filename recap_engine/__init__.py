"""
recap_engine — Professional, Object-Oriented Jewelry Pricing & Recap Automation Engine
=======================================================================================
A modular, cross-platform Python package for parsing jewelry pricing sheets and
updating client Recap presentation workbooks.
"""

from .models import StyleItem, RecapRow, MetalDetails, StoneDetails
from .resolvers import QualityCodeResolver, ToleranceResolver, DescriptionHelper
from .parsers import ExcelMatrixReader, PricingBlockParser
from .builder import RecapWorkbookBuilder
from .pipeline import RecapPipeline

__all__ = [
    "StyleItem",
    "RecapRow",
    "MetalDetails",
    "StoneDetails",
    "QualityCodeResolver",
    "ToleranceResolver",
    "DescriptionHelper",
    "ExcelMatrixReader",
    "PricingBlockParser",
    "RecapWorkbookBuilder",
    "RecapPipeline",
]
