#!/usr/bin/env python3
"""Synchronize regenerated package figures with the five upload files."""

from pathlib import Path
from shutil import copy2


ROOT = Path(__file__).resolve().parents[2]
PACKAGE_FIGURES = ROOT / "Zenodo_Package" / "figures"
SUBMISSION_FIGURES = ROOT / "Submission_Figures"

MAPPING = {
    "Figure_1_evidence_gates.png": "Figure_1.png",
    "Figure_2_cost_trajectories_and_rank_correlations.png": "Figure_2.png",
    "Figure_3_monte_carlo_cost_envelope.png": "Figure_3.png",
    "Figure_4_integrated_assessment_chain.png": "Figure_4.png",
    "Figure_5_tis_rubric_and_provenance.png": "Figure_5.png",
}


def main() -> None:
    SUBMISSION_FIGURES.mkdir(parents=True, exist_ok=True)
    for source_name, target_name in MAPPING.items():
        source = PACKAGE_FIGURES / source_name
        if not source.exists():
            raise FileNotFoundError(source)
        copy2(source, SUBMISSION_FIGURES / target_name)
    print("Synchronized five regenerated submission figures.")


if __name__ == "__main__":
    main()
