"""
recap_engine.pipeline
=====================
Orchestrates the complete pricing-to-recap transformation pipeline.
Connects matrix parsing, code resolution, tolerance mapping, and presentation building.
"""

import os
import time
from dataclasses import dataclass
from typing import List, Dict, Any, Optional

from .parsers import PricingBlockParser
from .resolvers import QualityCodeResolver, ToleranceResolver
from .builder import RecapWorkbookBuilder


@dataclass
class PipelineResult:
    """Represents the execution outcome of processing a project pair."""
    pair_id: int
    name: str
    status: str
    styles_count: int = 0
    output_recap: str = ""
    elapsed_seconds: float = 0.0
    error_message: Optional[str] = None


class RecapPipeline:
    """
    Main orchestration engine managing shared datasets, parsing, and recap generation.
    """

    def __init__(self, mapping_file: str, tolerance_csv: str):
        print("Initializing RecapPipeline (loading shared lookups)...")
        self.code_resolver = QualityCodeResolver(mapping_file)
        self.tolerance_resolver = ToleranceResolver(tolerance_csv)
        self.parser = PricingBlockParser(stone_codes=set(self.code_resolver.stone_map.keys()))
        self.builder = RecapWorkbookBuilder()
        print("Shared lookups initialized successfully.")

    def process_project(
        self,
        pricing_file: str,
        recap_template: str,
        output_dir: str,
        output_recap_name: str,
        pair_id: int = 1,
        project_name: str = "Project",
        customer_name: Optional[str] = None
    ) -> PipelineResult:
        """
        Executes end-to-end transformation for a single project pair.
        """
        t0 = time.time()
        print(f"\n[{pair_id}] Starting transformation: {project_name}")
        print(f"    Pricing Sheet : {os.path.basename(pricing_file)}")
        print(f"    Recap Template: {os.path.basename(recap_template)}")
        if customer_name:
            print(f"    Customer Name : {customer_name}")

        if not os.path.exists(pricing_file):
            return PipelineResult(
                pair_id=pair_id, name=project_name, status="FAILED",
                error_message=f"Pricing file missing: {pricing_file}"
            )
        if not os.path.exists(recap_template):
            return PipelineResult(
                pair_id=pair_id, name=project_name, status="FAILED",
                error_message=f"Recap template missing: {recap_template}"
            )

        os.makedirs(output_dir, exist_ok=True)
        csv_path = os.path.join(output_dir, "structured_styles.csv")
        recap_out_path = os.path.join(output_dir, output_recap_name)

        try:
            # 1. Parse pricing blocks
            styles = self.parser.parse_file(pricing_file)
            print(f"    -> Parsed {len(styles)} unique styles (production duplicate quotes selected).")

            # 2. Export structured CSV for audit trail
            self.parser.export_to_csv(styles, csv_path)

            # 3. Load template metadata & build recap presentation rows
            metadata = self.builder.load_template_metadata(recap_template)
            recap_rows = self.builder.prepare_recap_rows(
                items=styles,
                code_resolver=self.code_resolver,
                tolerance_resolver=self.tolerance_resolver,
                template_metadata=metadata
            )

            # 4. Update Recap workbook with preserved styling
            saved_path = self.builder.update_recap(
                template_path=recap_template,
                rows=recap_rows,
                output_path=recap_out_path,
                customer_name=customer_name,
                gold_lock_rate=getattr(self.parser, 'gold_lock_rate', 4250.0),
                silver_lock_rate=getattr(self.parser, 'silver_lock_rate', 65.0)
            )

            elapsed = round(time.time() - t0, 2)
            print(f"    -> Saved updated recap: {os.path.basename(saved_path)} ({elapsed}s)")

            return PipelineResult(
                pair_id=pair_id,
                name=project_name,
                status="SUCCESS",
                styles_count=len(styles),
                output_recap=saved_path,
                elapsed_seconds=elapsed
            )

        except Exception as e:
            return PipelineResult(
                pair_id=pair_id,
                name=project_name,
                status="FAILED",
                error_message=str(e),
                elapsed_seconds=round(time.time() - t0, 2)
            )

    def process_all(
        self,
        pairs: List[Dict[str, Any]],
        root_dir: str,
        default_template: Optional[str] = None
    ) -> List[PipelineResult]:
        """
        Processes a list of configured project dictionaries across the pipeline.
        Defaults to static/Recap-Template.xlsx as the base template.
        """
        results: List[PipelineResult] = []
        files_bt_dir = os.path.join(root_dir, "files_BT")
        outputs_dir = os.path.join(root_dir, "outputs")
        base_template_fallback = default_template or os.path.join(root_dir, "static", "Recap-Template.xlsx")

        for p in pairs:
            pricing_p = os.path.join(files_bt_dir, p['pricing_file'])
            
            # Use specific recap template if provided, else use base template
            if p.get('recap_file'):
                recap_p = os.path.join(files_bt_dir, p['recap_file'])
                if not os.path.exists(recap_p):
                    recap_p = base_template_fallback
            else:
                recap_p = base_template_fallback

            out_d = os.path.join(outputs_dir, p['folder'])

            res = self.process_project(
                pricing_file=pricing_p,
                recap_template=recap_p,
                output_dir=out_d,
                output_recap_name=p['recap_out_name'],
                pair_id=p['id'],
                project_name=p['name'],
                customer_name=p.get('customer')
            )
            results.append(res)

        return results
