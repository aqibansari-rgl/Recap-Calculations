"""
recap_engine.resolvers
======================
Resolvers and lookups for mapping quality codes to human-readable gemological names,
converting carat weights to trade fractions, and formatting style descriptions.
"""

import os
import re
import csv
from typing import Dict, List, Tuple, Optional
import openpyxl


class QualityCodeResolver:
    """
    Resolves cryptic diamond, lab-grown, and gemstone quality codes into
    descriptive trade names using Code_Mapping.xlsx and compound token parsing.
    """

    SHAPE_LGD: Dict[str, str] = {
        'M': 'Lab Marquise Diamond',
        'R': 'Lab Round Diamond',
        'OW': 'Lab Oval Diamond',
        'OV': 'Lab Oval Diamond',
        'CU': 'Lab Cushion Diamond',
        'CUL': 'Lab Cushion Diamond',
        'PR': 'Lab Princess Diamond',
        'PER': 'Lab Pear Diamond',
        'EM': 'Lab Emerald Cut Diamond',
        'RA': 'Lab Radiant Diamond',
        'A': 'Lab Pear Diamond',
        'AL': 'Lab Pear Diamond',
        'OWL': 'Lab Oval Diamond',
    }

    SHAPE_NDA: Dict[str, str] = {
        'RB': 'Round Black Diamond',
        'R': 'Round Diamond',
        'M': 'Marquise Diamond',
        'OV': 'Oval Diamond',
        'CU': 'Cushion Diamond',
        'PR': 'Princess Diamond',
    }

    def __init__(self, mapping_file: Optional[str] = None):
        self.dia_map: Dict[str, str] = {}
        self.lgd_map: Dict[str, str] = {}
        self.stone_map: Dict[str, str] = {}
        if mapping_file and os.path.exists(mapping_file):
            self.load_mapping_file(mapping_file)

    def load_mapping_file(self, filepath: str) -> None:
        """Loads sheets 'Dia', 'LGD', and 'Stones' from Code_Mapping.xlsx."""
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        for sheet_name, target_dict in [('Dia', self.dia_map), ('LGD', self.lgd_map), ('Stones', self.stone_map)]:
            if sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                first = True
                for row in ws.iter_rows(values_only=True):
                    if first:
                        first = False
                        continue
                    if row and row[0] is not None:
                        key = str(row[0]).strip().upper()
                        val = str(row[1]).strip() if len(row) > 1 and row[1] else key
                        target_dict[key] = val
        wb.close()

    @staticmethod
    def _parse_quality_grade(suffix: str) -> str:
        """Splits quality grades like 'FVS2' -> 'F VS2', 'GSI1' -> 'G SI1'."""
        m = re.match(r'^([A-Z]{1,2})(VS2|VS1|SI1|SI2|I1|IF|VVS1|VVS2|FL)$', suffix, re.IGNORECASE)
        if m:
            return f"{m.group(1).upper()} {m.group(2).upper()}"
        return suffix

    def parse_compound_code(self, code: str) -> Optional[str]:
        """
        Parses compound codes combining shape, diamond type, and quality (e.g. MLGDFVS2).
        """
        c = code.strip().upper()

        # Pattern 1: Lab-grown: {shape}LGD{quality}
        if 'LGD' in c:
            idx = c.index('LGD')
            shape = c[:idx]
            qual = c[idx + 3:]
            shape_desc = None
            for k in sorted(self.SHAPE_LGD.keys(), key=len, reverse=True):
                if shape == k:
                    shape_desc = self.SHAPE_LGD[k]
                    break
            shape_desc = shape_desc or 'Lab Diamond'
            qual_desc = self._parse_quality_grade(qual) if qual else ''
            return f"{shape_desc} – {qual_desc}" if qual_desc else shape_desc

        # Pattern 2: Natural / treated: {shape}WNDA | {shape}NDA | {shape}BDA
        for suffix in ('WNDA', 'NDA', 'BDA'):
            if c.endswith(suffix):
                prefix = c[: len(c) - len(suffix)]
                shape_desc = self.SHAPE_NDA.get(prefix, f'Diamond ({prefix})')
                label = {'WNDA': 'White Natural', 'NDA': 'Natural', 'BDA': 'Black'}[suffix]
                return f"{shape_desc} ({label})"

        return None

    def resolve_code(self, code: str, *maps: Dict[str, str]) -> str:
        """Looks up code through prioritized dictionaries with fallback to raw code."""
        if not code:
            return ""
        key = code.strip().upper()
        for m in maps:
            if key in m:
                return m[key]
        return code.strip()

    def resolve_diamond_quality(self, code: str) -> str:
        """
        Resolves diamond quality code.
        Priority: Compound parser -> LGD map -> Dia map -> raw code.
        """
        if not code:
            return ""
        compound = self.parse_compound_code(code)
        if compound:
            return compound
        return self.resolve_code(code, self.lgd_map, self.dia_map, self.stone_map)

    def resolve_gem_quality(self, code: str) -> str:
        """
        Resolves gemstone quality code.
        Priority: Stones map -> Dia map -> Compound parser -> raw code.
        """
        if not code:
            return ""
        resolved = self.resolve_code(code, self.stone_map, self.dia_map, self.lgd_map)
        if resolved == code.strip():
            compound = self.parse_compound_code(code)
            if compound:
                return compound
        return resolved


class ToleranceResolver:
    """
    Maps continuous carat weights (e.g. 2.18 cttw) to commercial trade fraction ranges
    (e.g. 2 1/5 cttw) using tollerance.csv.
    """

    def __init__(self, tolerance_csv: Optional[str] = None):
        self.ranges: List[Tuple[float, float, str]] = []
        if tolerance_csv and os.path.exists(tolerance_csv):
            self.load_tolerance_file(tolerance_csv)

    def load_tolerance_file(self, filepath: str) -> None:
        """Parses tolerance ranges from tollerance.csv."""
        with open(filepath, encoding='utf-8') as f:
            for row in csv.reader(f):
                if not row or not row[0].strip():
                    continue
                rng_str = row[0].strip()
                if 'fraction' in rng_str.lower():
                    continue
                m = re.match(r'([\d\.]+)\s*-\s*([\d\.]+)', rng_str)
                if m and len(row) > 2 and row[2].strip():
                    self.ranges.append((
                        float(m.group(1)),
                        float(m.group(2)),
                        re.sub(r'\s+', ' ', row[2].strip())
                    ))

    def get_fraction(self, weight: float) -> str:
        """Returns the commercial fraction representation for a carat weight."""
        if not weight or weight <= 0:
            return ""
        for low, high, fraction in self.ranges:
            if low <= weight <= high:
                return fraction
        return f"{weight:.2f}"


class DescriptionHelper:
    """Helper utilities for formatting center stone descriptions and item text."""

    @staticmethod
    def get_center_stone_desc(dia_size: str, dia_wt: float) -> str:
        """Builds center diamond annotation if center stone is >= 1.0ct."""
        s = str(dia_size or '').upper()
        wt_int = round(float(dia_wt)) if dia_wt else 2
        shape_names = {
            'MRQ': 'Marqise',
            'OVAL': 'Oval',
            'CUS': 'Cushion',
            'PER': 'Pear',
            'RND': 'Round'
        }
        for code, name in shape_names.items():
            if code in s:
                return f"Center {name} {wt_int}ct"
        return ""

    @staticmethod
    def build_item_description(
        metal_desc: str,
        ctw_desc: str = "",
        dia_desc: str = "",
        gem_desc: str = "",
        style_no: str = "",
        dia_val: Optional[str] = None,
        character_name: str = ""
    ) -> str:
        """Assembles a full descriptive string: Metal  Fraction ctw  Dia Qly  Gem Qly  [Character Name]."""
        parts = [metal_desc]
        token = ctw_desc or dia_val or ""
        if token:
            parts.append(token)
        if dia_desc:
            parts.append(dia_desc)
        if gem_desc:
            parts.append(gem_desc)
        if character_name:
            parts.append(character_name)
        # style_no is omitted from all descriptions as per user request
        return '  '.join(p for p in parts if p)
