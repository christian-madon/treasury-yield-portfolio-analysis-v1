# Progress and Codex workflow

Snapshot: November 1, 2026; P013 charting milestone.

## Workflow demonstrated

Christian chose the analytical questions and worked through bounded prompts with a planning assistant. Codex in VS Code implemented each requested step. Saved source, captured outputs, unsuccessful attempts and checkpoints were retained in the private development repository. Output review and independent checks were kept distinct from source review.

The synthetic holdings were prepared manually in Excel. The official Treasury data was preserved as a raw snapshot. Implementation progressed through CSV loading, a list of holding dictionaries, semiannual cash flows/present value, yield-scenario repricing, changes, portfolio weights, duration/DV01, approximation comparison and charts.

A typical prompt specified: the exact input and units; the function inputs and returned data structure; validation rules; what existing behavior to preserve; the execution boundary; and evidence to capture. For example, valuation changes used already-calculated PVs and ID-based matching, rather than repricing again or assuming matching list positions.

## Concepts explored

- CSV column-name access and Excel rows mapped to dictionary keys.
- Lists of dictionaries, nested dictionaries and functions that return new structures.
- Distinguishing annual percentage yields from semiannual decimal discount rates.
- Signed dollar changes, percentage changes and percentage-point weight differences.
- Market-value-weighted duration, summed dollar DV01 and first-order sensitivity estimates.
- Comparing an approximation with the same model's full repricing.
- Separating portfolio composition from attribution of gains/losses.
- Reproducible checkpoints and honest recorded execution/verification status.

These are topics explored with assistance. They are not a certification of independent mastery. Later projects can document which parts Christian implements and explains independently.

## Scope kept outside this publication

This is a curated publication copy, not the complete development transcript/history. Local machine paths, editor settings, broader personal planning, raw command logs and unfinished later tasks are omitted. The full records remain local. The financial source and both input files are unchanged from the identified checkpoint.

The Excel exporter, sourced CIO narrative, real-bond market-price comparison and later enhancements are outside this snapshot.
