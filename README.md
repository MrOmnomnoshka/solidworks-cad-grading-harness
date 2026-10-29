# SolidWorks CAD Grading Harness

A geometry-based evaluation harness for the PlayStation controller CAD benchmark.

This project implements a continuous scoring system for candidate SolidWorks parts. It compares measured geometry against frozen baseline measurements and evaluates the requested design transformations.

## Overview

The harness uses the SolidWorks COM API through `pywin32` to inspect a candidate part and calculate a weighted score from 0 to 4.

Body matching and geometric scoring use properties such as volume, surface area, centroids, bounding boxes, and face-level geometry. The implementation does not depend on feature or body names to establish geometric roles.

## Scoring criteria

| Component                    |   Weight |
|------------------------------|---------:|
| Feature-tree rebuild health  |     0.45 |
| Widening                     |     1.20 |
| Button-cluster placement     |     1.10 |
| Centre components            |     0.35 |
| Shoulder marking orientation |     0.55 |
| Button-symbol geometry       |     0.25 |
| Intersections                |     0.10 |
| **Total**                    | **4.00** |

The scores are continuous within each component, and the overall score is the weighted sum of the component scores.

### Evaluation details

- **Widening:** measures the increase in separation between mirrored pairs of sticks, triggers, and bumpers.
- **Button clusters:** compares the relative layout of the D-pad and face-button groups, as well as their target positions.
- **Centre components:** evaluates the placement of the START/SELECT group and central button relative to the expected mirrored positions.
- **Shoulder marking:** uses an asymmetric face-area signature of the shoulder marking as a geometric indicator of an incorrect mirror operation. It does not perform OCR.
- **Button symbols:** evaluates face-level geometric detail as a proxy for modeled 3D symbol geometry. It does not identify individual symbols or verify their chirality.
- **Feature-tree health:** uses SolidWorks rebuild diagnostics to measure model integrity, rather than matching feature names or tree topology.
- **Intersections:** checks mapped reference-role bodies. If the check cannot be completed, the component is marked unavailable and a warning is reported.

See [SCORING_NOTES.md](SCORING_NOTES.md) for scoring decisions, dataset observations, and known limitations.

## Requirements

- Windows with a licensed SolidWorks installation
- Python with `pywin32` installed
- The task assets and frozen baseline measurements

SolidWorks is controlled through its COM API. The harness attaches to an existing SolidWorks instance when available, or attempts to launch new headless SolidWorks when no instance is running. If a model path is supplied, the harness opens that document; otherwise, it evaluates the active document.


## Setup

This section provides a replacement for the original `harness.py` template file from the repository: 
https://github.com/arwinai/test_example/tree/main/test_example.

To run and verify the script, replace the original `harness.py` file in the test assignment directory with this new version. Additionally, complete the required setup steps: download the dependencies, as well as the CAD models and reference files.

## Execution

From the `SolidWorks\1_playstation_controller` directory, run:

```bat
python tests\task\harness\harness.py solution\solution.SLDPRT
```

To run the script using an example designed to test robustness:

```bat
python tests\task\harness\harness.py examples\adversarial_widened_by_30mm\adversarial_widened_by_30mm.SLDPRT
```

## Output

Each invocation returns a JSON score envelope containing:

- `score`: weighted total score
- `max_score`: maximum possible score
- `passed`: whether the configured acceptance threshold is met
- `subscores`: normalized component scores, weights, and awarded points

The harness reserves standard output for this machine-readable result. Verbose diagnostics can be enabled with `--debug`; diagnostic logs are also written to standard output.

## Scope

This is a task-specific grading harness. Its role is to evaluate the requested geometric changes reproducibly and provide interpretable component scores and diagnostic warnings.