# Proposed tBLG-coated hard-carbon SIB: evidence-gated revision package

This package accompanies the revised manuscript *An Evidence-Gated Assessment
of Proposed Twisted Bilayer Graphene Coatings for Hard-Carbon Sodium-Ion
Anodes*. It supersedes the repository content associated with Zenodo v1.0.1.

## Scope and limits

The package contains archived **screening inputs and outputs**, not measured
tBLG-coated sodium-ion battery data, process data, or validated forecasts. The
analysis cannot establish coating feasibility, cell performance, scale-up cost,
commercial competitiveness, or a policy recommendation.

The legacy TIS score files do not contain raw publication/patent/company
records, weights, independent coding, or expert elicitation. They are retained
only to document the earlier workflow and are not used for quantitative or
comparative inference. `TIS_Evidence_Audit.csv` is the evidence-provenance table
used in the revised manuscript and Figure 5.

## Contents

```
data/
  TEA_Cost_Trajectories.csv          Archived illustrative cost inputs
  Monte_Carlo_TEA_2035_Cost.csv      10,000 conditional screening runs
  Monte_Carlo_TEA_Summary.json       Archived summary metadata
  TIS_Evidence_Audit.csv             Evidence-availability/provenance audit
  TIS_Scores.csv                     Legacy author-coded labels; not inferential
  Monte_Carlo_TIS_10000.csv          Legacy resampling file; not inferential
  recomputed_*.csv                   Created when the script is run
figures/
  Figure_1_*.png ... Figure_5_*.png  Recomputed archival renderings
  Graphical_Abstract.pptx/.png/.pdf  Editable source and upload-ready copies
src/generate_figures.py              Recomputes summaries and draws figures
src/sync_submission_figures.py       Copies regenerated figures to upload names
src/build_graphical_abstract.py      Rebuilds the editable graphical abstract
src/extract_submission_assets.py     Exports manuscript-embedded figure files
src/revise_submission.py             Applies the evidence-gated manuscript revision
src/revise_cover_letter.py           Updates the revision cover letter
src/build_response_to_reviewers.py   Builds the point-by-point response
src/embed_submission_figures.py      Replaces selected Word figure binaries
src/update_figure_designs.py         Updates captions after artwork redesign
src/finalize_manuscript_layout.py    Applies idempotent final layout-only fixes
src/validate_submission.py           Runs structural and cross-file checks
requirements.txt                     Minimal runtime dependencies
```

## Reproduce

```bash
python -m pip install -r requirements.txt
python src/generate_figures.py
python src/sync_submission_figures.py
python src/revise_submission.py
python src/revise_cover_letter.py
python src/build_response_to_reviewers.py
python src/update_figure_designs.py
python src/embed_submission_figures.py --numbers 1 2 3 4 5
python src/finalize_manuscript_layout.py
python src/validate_submission.py
```

`revise_submission.py` is a full rebuild script intended for the preserved
pre-revision manuscript, not for repeated execution on the final revised file.
For layout-only updates to an already revised manuscript, use
`finalize_manuscript_layout.py` instead.

The figure script writes recomputed cost summaries and Spearman rank
associations to `data/` and regenerates five 600-dpi PNG figures. Figure titles
are supplied only in the manuscript captions, not inside the image files.
Figures 1, 4, and 5 use non-sequential scientific schematics rather than
flowcharts or tables. Figure 5 is a pictorial evidence-provenance audit; its
legacy filename is retained for package compatibility. None of the figures
adds experimental evidence.

In Figure 1, the two planar lattice drawings are clipped to their physical
sheet polygons and remain inside the aqueous vessel. Figure 1 and Figure 5
use generous bottom safe areas; their permitted-inference conclusions are
carried by the manuscript captions rather than title-like text inside the
artwork.
The `EVIDENCE BOUNDARY` label sits directly above its dashed separator and
outside the supported-system panel. The `PROPOSED • NOT TESTED` label is
centered over the complete untested domain, including the particle and the
downstream cell, process, commercial, and policy contexts.

The separately uploaded figures in `../Submission_Figures/` are exact copies of
the regenerated package figures and are the binaries embedded in the final
manuscript. Their aspect ratios match the existing Word figure frames.

## Zenodo release instruction

Publish this directory as a **new version** of the existing Zenodo concept
record `10.5281/zenodo.19362805`. Do not overwrite or cite v1.0.1 as if it
contained the revised data, figures, and scripts. After publishing the new
version, update the manuscript and submission metadata with its version DOI;
the all-versions concept DOI remains unchanged.

## License

MIT License. The authors remain responsible for verifying that the content and
metadata of the published Zenodo release match the manuscript submitted.
