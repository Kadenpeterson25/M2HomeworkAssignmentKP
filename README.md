# M2HomeworkAssignmentKP

This repository contains a deterministic workflow to rank required MAcc core courses using the 2024 graduate exit survey.

## Research question
Rank required MAcc courses based on their average overall course rating for the 2024 survey year.

## Data input
Primary expected input path:

- `data/grad_exit_survey_2024.xlsx`

For compatibility with the original repository file name, the script and CI job can fall back to:

- `Grad Program Exit Survey Data 2024 (1).xlsx`

## What the workflow does
`python scripts/rank_required_courses.py` performs the following steps:

1. Reads the Excel workbook directly (no third-party Python dependencies).
2. Detects required core course columns from the survey question metadata.
3. Converts rating cells to numeric values and removes blank/missing values.
4. Aggregates by course:
   - mean rating
   - response count
5. Ranks courses by:
   - descending mean rating
   - ascending alphabetical course name (tie-break)
6. Writes outputs:
   - `outputs/rank_order.csv`
   - `outputs/rank_order.png` (horizontal bar chart ordered highest-to-lowest)

## Reproducibility in GitHub Actions
On every push, GitHub Actions runs `.github/workflows/rank-required-courses.yml`, which:

1. Checks out the repository
2. Sets up Python 3.11
3. Ensures `data/grad_exit_survey_2024.xlsx` exists (copies from the original filename if needed)
4. Runs the ranking script
5. Uploads the generated files as workflow artifacts

## Note on generated files
`outputs/rank_order.csv` and `outputs/rank_order.png` are generated artifacts and are gitignored to keep pull requests text-only and reviewable.
