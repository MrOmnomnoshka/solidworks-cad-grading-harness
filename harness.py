#!/usr/bin/env python3
"""Grading harness for 1_playstation_controller.

Usage (Windows, with (or without) SolidWorks running):
    python tests/task/harness/harness.py [candidate path]

Prints a JSON score envelope to stdout:
    {"score": ..., "max_score": 4, "passed": ..., "subscores": {...}}
"""

import json
import math
import statistics
import sys
from itertools import combinations, permutations
from pathlib import Path

import pythoncom
import win32com.client
from win32com.client import VARIANT

# import threading
# NOTE: This task cannot be parallelized with threading inside a single SolidWorks process.
#  The SolidWorks API is strictly single-threaded, even when using pythoncom.CoInitialize().
#  Multiprocessing is possible, but each process must launch its own SLDWORKS.exe instance via DispatchEx.


MAX_SCORE = 4  # keep equal to max_score in task.toml

# CHANGEABLE
DEBUG_FULL_OUTPUT = False
CLOSE_OUR_SW_IN_THE_END = True
# /CHANGEABLE

# SolidWorks constants
SW_SOLID_BODY = 0  # swSolidBody type in the SolidWorks API
SW_BODY_INTERSECT = 15901  # Boolean intersection operation
SW_DOC_PART = 1  # Document type: '.SLDPRT'

# Task root: the directory containing the harness, tests, environment, and other project files.
PROJECT_DIR = Path(__file__).resolve().parents[1]

# Baseline measurements captured from the reference model.
BASELINE_PATH = PROJECT_DIR / "prompt" / "input.json"

# Shared program constants
SMALL_FACE_AREA_M2 = 15e-6  # Face-area threshold: 15 mm²
TARGET_WIDENING_MM = 15.0
HALF_TARGET_WIDENING_M = TARGET_WIDENING_MM / 1000 / 2


# ========================= Baseline =========================
def load_baseline():
    """
    Loads the baseline measurements and validates the annotations.
    Body names are used only within the baseline JSON;
    geometric matching is performed independently.
    """

    if not BASELINE_PATH.is_file():
        raise FileNotFoundError(f"Baseline file not found: {BASELINE_PATH}")

    with BASELINE_PATH.open("r", encoding="utf-8") as file:
        baseline = json.load(file)

    if baseline.get("schema") != "ps-annotation-baseline/1":
        raise ValueError("Unrecognized input.json schema.")

    bodies = baseline.get("bodies", [])
    roles = baseline.get("roles", {})

    body_ids = [body["id"] for body in bodies]

    if len(body_ids) != len(set(body_ids)):
        raise ValueError("The baseline contains duplicate body IDs.")

    known_ids = set(body_ids)
    assigned_ids = [body_id for group_ids in roles.values() for body_id in group_ids]

    if len(assigned_ids) != len(set(assigned_ids)):
        raise ValueError("A body is assigned to more than one role.")

    unknown_ids = set(assigned_ids) - known_ids

    if unknown_ids:
        raise ValueError(f"Roles reference unknown IDs: {sorted(unknown_ids)}")

    unassigned_ids = known_ids - set(assigned_ids)

    return baseline, unassigned_ids


def print_baseline(baseline, unassigned_ids):
    print_if_debug("\nBaseline:")
    print_if_debug(f"File: {BASELINE_PATH}")
    print_if_debug(f"Schema: {baseline['schema']}")
    print_if_debug(f"Bodies in baseline: {len(baseline['bodies'])}")
    print_if_debug(f"X symmetry plane: {baseline['plane_x_m'] * 1000:.3f} mm")

    print_if_debug("\nBaseline roles:")
    for role, body_ids in baseline["roles"].items():
        print_if_debug(f"  {role}: {len(body_ids)} bodies")

    print_if_debug(f"\nBodies without an assigned role: {sorted(unassigned_ids)}")
...  # ==================== /Baseline =========================


# ========================= Global Matcher =========================
def match_bodies_global(baseline, records, missing_cost=2.5, min_margin=0.08, position_scale_mm=25.0):
    """
    Matches baseline bodies to bodies in the current model.

    For each baseline body, identifies several of the most
    suitable current-model candidates across the allowed transformation hypotheses.

    Constraints:
    - a current Solid Body cannot be assigned to more than one baseline body;
    - D-pad and face_buttons are searched for only within recognized clusters;
    - position is compared against the permitted transformation states;
    - overly costly or ambiguous matches are rejected;
    - additional bodies remain unmatched.

    The missing_cost and min_margin parameters are preliminary
    settings and will be tuned using the test models.
    """

    if not records:
        raise ValueError("The current model contains no Solid Bodies.")

    baseline_by_id = {body["id"]: body for body in baseline["bodies"]}
    role_by_id = {body_id: role for role, ids in baseline["roles"].items() for body_id in ids}

    small_face_counts = baseline.get("small_face_counts", {})

    # Identify the housing separately by volume (V), so changes to geometric features do not cause it to be confused with smaller parts.
    # Use the largest body as an anchor; this is not guaranteed to be the exact housing body from input.json.
    housing_anchor = max(records, key=body_volume_mm3)

    housing_id = baseline["roles"]["housing"][0]
    matches = {housing_id: housing_anchor}
    costs = {}
    details = {
        housing_id: {
            "state": "anchor_only",
            "matched": True,
            "match_cost": None,
            "shape_cost": None,
            "position_error_mm": None,
            "margin": None,
            "s_faces": (small_face_counts.get(housing_id), housing_anchor['small_face_count']),
            "already_classified": False,
        }
    }
    used_indices = {housing_anchor["index"]}

    # First, identify the physical clusters and classify each cluster.
    clusters = find_button_clusters(records, baseline["plane_x_m"])
    classification = classify_button_clusters(baseline, clusters)
    cluster_mapping = classification["mapping"] if classification["confident"] else None

    # When classification is confident, use the detected button matches as a soft prior during
    # global matching. This prior does not lock the assignment; stronger geometric evidence can still override it.
    button_matches = (classification.get("button_matches", {}) if classification["confident"] else {})

    button_layout_penalty = 0.05  # Small penalty favoring the button already identified above

    cluster_indices = {side: {body["index"] for body in bodies} for side, bodies in clusters.items()}
    all_button_indices = cluster_indices["left"] | cluster_indices["right"]

    # Exclude the housing
    references = [body for body in baseline["bodies"] if body["id"] != housing_id]
    candidates = [body for body in records if body["index"] not in used_indices]

    # The cost of an invalid assignment is set far above the cost of leaving a body unmatched.
    forbidden_cost = 1e6
    matrix = []
    candidate_details = {}
    row_info = []

    for row_index, reference in enumerate(references):
        reference_id = reference["id"]
        role = role_by_id.get(reference_id, "unclassified")
        hypotheses = get_transform_hypotheses(reference, role, baseline)
        candidate_scores = {}
        row = []

        for candidate in candidates:
            current_index = candidate["index"]

            # The likely button bodies have already been identified above.
            # When matching buttons, ignore all other bodies, and vice versa. This constrains the physical
            # groups and prevents buttons from being misclassified as central components, or vice versa.
            allowed = True

            if role in ("dpad", "face_buttons"):
                allowed = (cluster_mapping is not None and current_index in cluster_indices[cluster_mapping[role]])
            elif current_index in all_button_indices:
                allowed = False

            if not allowed or not hypotheses:
                row.append(forbidden_cost)
                continue

            # Compare object shapes independently of their positions.
            shape_cost = body_shape_distance(reference, candidate)

            best = None
            # Evaluate every permitted permutation hypothesis.
            for hypothesis in hypotheses:
                # Euclidean distance, in millimeters, between the expected and current centroids.
                position_error_mm = _centroid_distance_mm(candidate["centroid_m"], hypothesis["centroid_m"])

                # Normalization: a 25 mm positional error adds 1.0 to the cost.
                position_cost = position_error_mm / position_scale_mm

                small_face_cost = 0.0
                reference_small_faces = small_face_counts.get(reference_id)
                candidate_small_faces = candidate.get("small_face_count", 0)
                if reference_small_faces is not None:
                    abs_diff = abs(candidate_small_faces - reference_small_faces)
                    small_face_cost = 0.003 * min(abs_diff, 100)

                # Slightly favor the match identified from the cluster’s internal layout.
                # If the geometry strongly favors another body, the global matcher can still select it.
                layout_penalty = 0.0
                if role in ("dpad", "face_buttons"):
                    preferred = button_matches.get(reference_id)

                    if preferred is not None and int(preferred["index"]) != current_index:
                        # The likely button has already been identified, so it receives no penalty; other candidates receive a 0.05 penalty.
                        layout_penalty = button_layout_penalty

                total_cost = shape_cost + position_cost + small_face_cost + layout_penalty

                if best is None or total_cost < best["match_cost"]:
                    best = {
                        "state": hypothesis["state"],
                        "match_cost": total_cost,
                        "shape_cost": shape_cost,
                        "position_error_mm": position_error_mm,
                        "s_faces": (reference_small_faces, candidate_small_faces),
                        "already_classified": bool(layout_penalty),
                    }

            candidate_scores[current_index] = best
            row.append(best["match_cost"])

        # Check for ambiguity: if two different bodies are nearly equally
        # suitable for the same role, do not assign the role unless the best candidate has a clear advantage.
        ranked = sorted(candidate_scores.items(), key=lambda item: item[1]["match_cost"])
        # The gap between the two best candidates indicates how decisively the best candidate outranks the runner-up.
        margin = ranked[1][1]["match_cost"] - ranked[0][1]["match_cost"] if len(ranked) > 1 else float("inf")
        second_name = records[ranked[1][0]]['name_debug'] if len(ranked) > 1 else "N/A"
        ambiguous = len(ranked) > 1 and margin < min_margin
        best_cost = ranked[0][1]["match_cost"] if ranked else float("inf")

        if ambiguous:
            row = [forbidden_cost] * len(candidates)

        # Dummy columns allow baseline bodies to remain unmatched.
        row.extend([missing_cost] * len(references))
        matrix.append(row)
        candidate_details[row_index] = candidate_scores
        row_info.append({
            "reference_id": reference_id,
            "role": role,
            "best_cost": best_cost,
            "margin": margin,
            "ambiguous": ambiguous,
            "candidate_count": len(ranked),
            "second_name": second_name,
        })

    assignment = _hungarian_assignment(matrix)

    for row_index, column_index in assignment.items():
        info = row_info[row_index]
        reference_id = info["reference_id"]

        if column_index >= len(candidates):
            if info["ambiguous"]:
                reason = "ambiguous_candidates"
            elif info["role"] in ("dpad", "face_buttons") and cluster_mapping is None:
                reason = "button_cluster_classification_uncertain"
            elif info["candidate_count"] == 0:
                reason = "no_allowed_candidates"
            else:
                reason = "missing_cost_preferred"

            details[reference_id] = {**info, "matched": False, "reason": reason}
            continue

        candidate = candidates[column_index]
        current_index = candidate["index"]
        selected = candidate_details[row_index].get(current_index)

        if selected is None or matrix[row_index][column_index] >= missing_cost:
            details[reference_id] = {**info, "matched": False, "reason": "match_cost_too_high"}
            continue

        matches[reference_id] = candidate
        costs[reference_id] = selected["match_cost"]
        used_indices.add(current_index)
        details[reference_id] = {**info, **selected, "matched": True}

    missing_reference_ids = set(baseline_by_id) - set(matches)
    unmatched = [body for body in records if body["index"] not in used_indices]
    unclassified_ids = set(baseline_by_id) - set(role_by_id)

    return {
        "matches": matches,
        "costs": costs,
        "unmatched": unmatched,
        "unclassified_ids": unclassified_ids,
        "missing_reference_ids": missing_reference_ids,
        "details": details,
        "button_classification": classification,
    }


def print_global_match_report(baseline, result):
    """Prints the global matching results grouped by role."""

    matches = result["matches"]
    details = result["details"]
    unmatched = result["unmatched"]

    role_by_id = {body_id: role for role, ids in baseline["roles"].items() for body_id in ids}
    reference_by_id = {body["id"]: body for body in baseline["bodies"]}

    print_if_debug("\n" + "=" * 90)
    print_if_debug("GLOBAL BODY MATCHING")
    print_if_debug("=" * 90)

    for role in baseline["roles"]:
        print_if_debug(f"\n{role}:")

        if role == "housing":
            reference_id = baseline["roles"]["housing"][0]
            anchor = matches.get(reference_id)

            if anchor is None:
                print_if_debug("  The largest housing component was not found.")
            else:
                print_if_debug(
                    f"  housing_anchor -> #{anchor['index']:02d} "
                    f"{body_name(anchor)[:22]:22} | "
                    f"V={body_volume_mm3(anchor):.1f} mm³"
                )

            continue

        for reference_id in baseline["roles"][role]:
            body = matches.get(reference_id)
            info = details.get(reference_id, {})

            if body is None:
                print_if_debug(f"  {reference_id}: NO MATCH | reason={info.get('reason', 'unknown')}")
                continue

            margin = info.get("margin")
            margin_text = f"{margin:.3f}" if margin is not None and math.isfinite(margin) else "n/a"

            print_if_debug(
                f"  {reference_id} -> #{body['index']:02d} {body_name(body)[:22]:22} | "
                f"state={info.get('state', 'n/a')} | "
                f"cost={info.get('match_cost', 0.0):.3f} | "
                f"shape={info.get('shape_cost', 0.0):.3f} | "
                f"pos={info.get('position_error_mm', 0.0):.2f} mm | "
                f"s_face={"r:{} h:{}".format(*info.get('s_faces'))} | "
                f"margin={margin_text} ({info.get('second_name')})"
            )

    print_if_debug("\nunclassified:")
    for reference_id in sorted(result["unclassified_ids"]):
        body = matches.get(reference_id)
        info = details.get(reference_id, {})

        if body is None:
            print_if_debug(f"  {reference_id}: NO MATCH | reason={info.get('reason', 'unknown')}")
        else:
            print_if_debug(
                f"  {reference_id} -> #{body['index']:02d} {body_name(body)[:22]:22} | cost={info.get('match_cost', 0.0):.3f}")

    classification = result["button_classification"]
    print_if_debug("\nButton clusters:")
    if classification["confident"]:
        for role, side in classification["mapping"].items():
            print_if_debug(f"  {role} -> {side}")
    else:
        print_if_debug("  Classification was not sufficiently confident.")

    # Compare button volume and surface area with the original baseline.
    # This diagnostic helps determine whether the button bodies change
    # when their engraved symbols are removed from the model.
    baseline_by_id = {body["id"]: body for body in baseline["bodies"]}

    print_if_debug("\nButton geometry relative to the baseline:")
    for role in ("dpad", "face_buttons"):
        print_if_debug(f"  {role}:")
        for ref_id in baseline["roles"][role]:
            reference = baseline_by_id[ref_id]
            candidate = matches.get(ref_id)

            if candidate is None:
                print_if_debug(f"    {ref_id}: no match")
                continue

            vol_delta = 100.0 * (candidate["volume_m3"] / reference["volume_m3"] - 1.0)
            area_delta = 100.0 * (candidate["area_m2"] / reference["area_m2"] - 1.0)
            print_if_debug(f"    {ref_id} -> #{candidate['index']:02d}: "
                           f"V={vol_delta:+.2f}%, "
                           f"A={area_delta:+.2f}%, "
                           f"faces={candidate.get('face_count', '?')}, "
                           f"small_faces={candidate.get('small_face_count', '?')}")

    print_if_debug(f"\nUnmatched baseline bodies: {len(result['missing_reference_ids'])}")
    print_if_debug(f"Additional bodies in the current model: {len(unmatched)}")

    for body in unmatched:
        print_if_debug(f"  #{body['index']:02d} {body_name(body)[:22]:22} | "
                       f"V={body_volume_mm3(body):.1f} mm³ | "
                       f"COM=({" | ".join(f'{value * 1000:.3f}' for value in body["centroid_m"])}) (mm)")
...  # ==================== /Global Matcher =========================


# ========================= Scoring =========================
def diagnose_feature_tree_health(document):
    """Reports the overall feature-tree health and errors in individual features."""
    print_if_debug("\n" + "=" * 90)
    print_if_debug("FEATURE-TREE DIAGNOSTICS")
    print_if_debug("=" * 90)

    whats_wrong = {"success": False, "problem_count": None, "items": []}
    try:
        whats_wrong = get_whats_wrong(document)
        print_if_debug(f"What's Wrong API call succeeded: {whats_wrong['success']}")
        print_if_debug(f"Issues reported by What's Wrong: {whats_wrong['problem_count']}")

        if whats_wrong["success"]:
            errors = [item for item in whats_wrong["items"] if not item["is_warning"]]
            warnings = [item for item in whats_wrong["items"] if item["is_warning"]]

            print_if_debug(f"Errors: {len(errors)}; warnings: {len(warnings)}")

            for item in whats_wrong["items"]:
                kind = "WARNING" if item["is_warning"] else "ERROR"
                print_if_debug(f"  [{kind}] {item['name']} [{item['type']}] code={item['error_code']}")

    except Exception as exc:
        print_if_debug(f"Failed to retrieve What's Wrong: {type(exc).__name__}: {exc}")

    additional_feature_errors = False  # NOTE: Finds one additional warning, but makes the full traversal much slower
    if additional_feature_errors:
        feature_errors = []
        scan_failures = 0
        scanned_count = 0

        for feature, depth, name, type_name in iter_feature_tree(document):
            scanned_count += 1

            try:
                error_code, is_warning = get_feature_error(feature)
            except Exception:
                scan_failures += 1
                continue

            if error_code:
                feature_errors.append({
                    "name": name,
                    "type": type_name,
                    "depth": depth,
                    "error_code": error_code,
                    "is_warning": is_warning,
                })

        print_if_debug(f"\nFeature-tree nodes checked: {scanned_count}")
        print_if_debug(f"Nodes that could not be checked: {scan_failures}")
        print_if_debug(f"Non-zero GetErrorCode2 codes: {len(feature_errors)}")

        for item in feature_errors:
            kind = "WARNING" if item["is_warning"] else "ERROR"
            indent = "  " * item["depth"]
            print_if_debug(f"{indent}[{kind}] {item['name']} [{item['type']}] code={item['error_code']}")

    return whats_wrong


def score_feature_tree_health(whats_wrong):
    """Scores feature-tree health from the What's Wrong issue list."""

    if not whats_wrong["success"]:
        return {"score": None, "status": "unavailable", "errors": 0, "warnings": 0}

    items = whats_wrong["items"]
    errors = sum(not item["is_warning"] for item in items)
    warnings = sum(item["is_warning"] for item in items)

    score = math.exp(-0.18 * errors - 0.08 * warnings)

    return {"score": score, "status": "ok", "errors": errors, "warnings": warnings}


def print_feature_tree_score(result):
    """Prints the numeric feature-tree health score."""

    print_if_debug("\nFeature-tree health score:")

    if result["status"] != "ok":
        print_if_debug("Could not calculate the score: the What's Wrong API is unavailable.")
        return

    print_if_debug(f"Errors: {result['errors']}; warnings: {result['warnings']}")
    print_if_debug(f"Feature-tree score: {result['score']:.3f} / 1.000")


def diagnose_widening_pairs(baseline, matches):
    """Shows the actual coordinates of baseline and matched bodies in each pair."""

    print_if_debug("\n" + "=" * 90)
    print_if_debug("WIDENING-PAIR DIAGNOSTICS")
    print_if_debug("=" * 90)

    print_if_debug("\nWidening-pair coordinates:")

    reference_by_id = {body["id"]: body for body in baseline["bodies"]}

    pair_definitions = get_widening_pairs(baseline)
    measurements = []

    for pair_name, (left_id, right_id) in pair_definitions.items():
        print_if_debug(f"\n{pair_name}:")

        # Do not fabricate coordinates when one of the paired bodies is missing.
        # Missing bodies are handled separately by the completeness check.
        if left_id not in matches or right_id not in matches:
            measurements.append({"pair": pair_name, "missing": True})
            print_if_debug(f"N/A:")
            continue

        reference_left = reference_by_id[left_id]
        reference_right = reference_by_id[right_id]

        current_left = matches[left_id]
        current_right = matches[right_id]

        # Measure separation along the X axis; the symmetry plane itself does not have to be at X = 0.
        x_left_r = reference_left["centroid_m"][0]
        x_right_r = reference_right["centroid_m"][0]
        baseline_separation_m = abs(x_left_r - x_right_r)

        x_left_c = current_left["centroid_m"][0]
        x_right_c = current_right["centroid_m"][0]
        current_separation_m = abs(x_left_c - x_right_c)

        growth_mm = (current_separation_m - baseline_separation_m) * 1000.0

        midpoint_mm = (current_left["centroid_m"][0] + current_right["centroid_m"][0]) * 1000.0 / 2

        print_if_debug(f"  {left_id} -> #{current_left['index']:02d} {body_name(current_left):22} | "
                       f"X: {x_left_r * 1000:.2f} -> {x_left_c * 1000:.2f} mm")
        print_if_debug(f"  {right_id} -> #{current_right['index']:02d} {body_name(current_right):22} | "
                       f"X: {x_right_r * 1000:.2f} -> {x_right_c * 1000:.2f} mm")

        print_if_debug(f"  Separation: {baseline_separation_m * 1000.0:.2f} -> "
                       f"{current_separation_m * 1000.0:.2f} mm | "
                       f"increase={growth_mm:+.2f} mm | "
                       f"midpoint X={midpoint_mm:.2f} mm")

        measurements.append({
            "pair": pair_name,
            "missing": False,
            "baseline_separation_mm": baseline_separation_m * 1000.0,
            "current_separation_mm": current_separation_m * 1000.0,
            "growth_mm": growth_mm,
            "midpoint_x_mm": midpoint_mm,
        })
    return measurements


def score_widening_pairs(measurements):
    """
    Computes a continuous widening score in the range 0..1.

    Use the geometric matching results and the body centroids.
    A full score corresponds to a 15 mm increase in the mirrored-pair separation.
    """

    valid_growths = [item["growth_mm"] for item in measurements if not item["missing"]]

    # The median remains a diagnostic metric but is no longer used directly in the final score.
    median_growth = (statistics.median(valid_growths) if valid_growths else None)

    # Score each mirrored pair independently. A pair receives 0 points if either body is missing.
    for item in measurements:
        if item["missing"]:
            item["score"] = 0.0
            continue

        deviation = abs(item["growth_mm"] - TARGET_WIDENING_MM)

        # Use a linear score based on deviation from the target increase of +15 mm.
        # Increases of 0 mm or 30 mm score 0; an increase of exactly 15 mm scores 1.
        item["score"] = max(0.0, min(1.0, 1.0 - deviation / TARGET_WIDENING_MM))

    # The final score is the mean across all expected pairs. A pair modified incorrectly lowers the overall result.
    pair_scores = [item["score"] for item in measurements]

    score = (statistics.fmean(pair_scores) if pair_scores else 0.0)

    return {
        "score": score,
        "target_growth_mm": TARGET_WIDENING_MM,
        "median_growth_mm": median_growth,
        "pairs": measurements,
    }


def print_widening_pairs_score(widening):
    print_if_debug("\nHousing-widening criterion:")
    for item in widening["pairs"]:
        if item["missing"]:
            print_if_debug(f"  {item['pair']}: one of the paired bodies is unmatched")
            continue

        print_if_debug(f"  {item['pair']}: (score={item['score']:.3f})")

    if widening["median_growth_mm"] is None:
        print_if_debug("Not enough pairs to compute the median increase.")
    else:
        print_if_debug(f"Median increase: {widening['median_growth_mm']:+.2f} mm")

    print_if_debug(f"Target increase: +{widening['target_growth_mm']:.1f} mm")
    print_if_debug(f"Widening score: {widening['score']:.3f} / 1.000")


def diagnose_button_clusters(baseline, records):
    """
    For diagnostics, identifies two clusters based on body positions
    and compares their relative geometry with the baseline groups.

    Body names are printed for human inspection only.
    Body names are not used by the scoring logic.
    """

    plane_x_mm = baseline["plane_x_m"] * 1000.0
    clusters = find_button_clusters(records, baseline["plane_x_m"])

    print_if_debug("\n" + "=" * 80)
    print_if_debug("BUTTON-CLUSTER DIAGNOSTICS")
    print_if_debug("=" * 80)

    for side, bodies in clusters.items():
        print_if_debug(f"\nCluster {side}: {len(bodies)} bodies")

        if not bodies:
            continue

        center_x = sum(body["centroid_m"][0] for body in bodies) / len(bodies)

        center_z = sum(body["centroid_m"][2] for body in bodies) / len(bodies)

        for body in sorted(bodies, key=lambda item: (item["centroid_m"][0], item["centroid_m"][2]),
                           ):
            x, y, z = body["centroid_m"]

            print_if_debug(
                f"  #{body['index']:02d} "
                f"{body['name_debug'][:22]:22} | "
                f"X={x * 1000:8.2f} mm | "
                f"Z={z * 1000:8.2f} mm | "
                f"dX={(x - center_x) * 1000:+7.2f} mm | "
                f"dZ={(z - center_z) * 1000:+7.2f} mm | "
                f"V={body['volume_m3'] * 1e9:9.1f} mm³ | "
                f"small_faces={body['small_face_count']}"
            )

    reference_by_id = {body["id"]: body for body in baseline["bodies"]}

    print_if_debug("\nRelative cluster-geometry error (mm):")

    for role in ("dpad", "face_buttons"):
        reference_group = [reference_by_id[body_id] for body_id in baseline["roles"][role]]

        for side, current_group in clusters.items():
            error = cluster_layout_rms(reference_group, current_group)

            if error is None:
                print_if_debug(f"  {role} -> {side}: cannot compare; found {len(current_group)} bodies")
            else:
                print_if_debug(f"  {role} -> {side}: RMS={error:.3f} mm")

    classification = classify_button_clusters(baseline, clusters)

    print_if_debug("\nAutomatic cluster classification:")

    if not classification["confident"]:
        print_if_debug("  No confident assignment was found.")
        print_if_debug(f"  Reason: {classification.get('reason', 'insufficient confidence')}")

        if "best_total_rms_mm" in classification:
            print_if_debug(
                f"  Best assignment: "
                f"{classification['best_total_rms_mm']:.3f} mm"
            )
            print_if_debug(
                f"  Alternative assignment: "
                f"{classification['alternative_total_rms_mm']:.3f} mm"
            )
    else:
        for role, side in classification["mapping"].items():
            print_if_debug(f"  {role} -> {side}")

        print_if_debug(
            f"  Best-assignment RMS: "
            f"{classification['best_total_rms_mm']:.3f} mm"
        )
        print_if_debug(
            f"  Margin over the alternative: "
            f"{classification['margin_mm']:.3f} mm"
        )


def score_button_clusters(baseline, records, cluster_classification=None):
    """
    Scores the swapping and placement of the D-pad and face buttons.

    Finds two four-body clusters, matches them to the baseline
    groups by relative layout, and scores the position of each
    group against its expected position after swapping and widening.

    Returns an overall score from 0 to 1 and per-group details.
    """

    plane_x = baseline["plane_x_m"]
    plane_x_mm = plane_x * 1000.0

    # Select button bodies by geometry, without using their names.
    candidates = find_bodies_by_feature(
        records,
        y_range_mm=(-12.0, 12.0),
        z_range_mm=(-35.0, 35.0),
        volume_range_mm3=(500.0, 10000.0),
        predicate=lambda body: abs(body_centroid_mm(body)[0] - plane_x_mm) >= 30.0,
    )

    # Exhaustively testing all partitions becomes prohibitively expensive when
    # there are too many candidates. The current models
    # are expected to contain 8 bodies; allow up to 12 to accommodate extras.
    if not 8 <= len(candidates) <= 12:
        return {
            "score": 0.0,
            "status": "candidate_count_invalid",
            "candidate_count": len(candidates),
            "candidates": candidates,
            "reason": f"Expected 8–12 candidates, found {len(candidates)}.",
        }

    reference_groups = {role: get_reference_bodies(baseline, role) for role in ("dpad", "face_buttons")}

    def center_of_group(bodies):
        """Returns the mean centroid of a group, in meters."""
        return tuple(sum(body["centroid_m"][axis] for body in bodies) / len(bodies) for axis in range(3))

    # Compute target centers by reflecting each group across the symmetry plane
    # and then moving it another 7.5 mm outward.
    targets = {}

    for role, bodies in reference_groups.items():
        original_center = center_of_group(bodies)
        direction = 1.0 if original_center[0] < plane_x else -1.0
        target_x = 2.0 * plane_x - original_center[0] + direction * HALF_TARGET_WIDENING_M

        targets[role] = (target_x, original_center[1], original_center[2])

    # Find two disjoint sets of four bodies that best match the
    # baseline D-pad and face_buttons. If there are more than eight candidates,
    # the remaining bodies are excluded from the button clusters.
    best = None

    button_matches = (
        cluster_classification.get("button_matches", {})
        if cluster_classification and cluster_classification.get("confident")
        else {}
    )

    dpad_ids = baseline["roles"]["dpad"]
    face_ids = baseline["roles"]["face_buttons"]

    # If the classifier has already matched all eight buttons,
    # use its result instead of enumerating four-body combinations.
    if all(ref_id in button_matches for ref_id in dpad_ids + face_ids):
        dpad_group = [button_matches[ref_id] for ref_id in dpad_ids]
        face_group = [button_matches[ref_id] for ref_id in face_ids]

        dpad_rms = cluster_layout_rms(reference_groups["dpad"], dpad_group)
        face_rms = cluster_layout_rms(reference_groups["face_buttons"], face_group)

        if dpad_rms is not None and face_rms is not None:
            dpad_center = center_of_group(dpad_group)
            face_center = center_of_group(face_group)

            dpad_position_error = _centroid_distance_xz_mm(dpad_center, targets["dpad"])
            face_position_error = _centroid_distance_xz_mm(face_center, targets["face_buttons"])

            best = {
                "cost": dpad_rms + face_rms + 0.05 * (dpad_position_error + face_position_error) / TARGET_WIDENING_MM,
                "dpad_group": dpad_group,
                "face_group": face_group,
                "dpad_rms": dpad_rms,
                "face_rms": face_rms,
                "dpad_center": dpad_center,
                "face_center": face_center,
                "dpad_position_error": dpad_position_error,
                "face_position_error": face_position_error,
            }

    # If classification is uncertain, fall back to the existing exhaustive search.
    candidate_indices = range(len(candidates)) if best is None else range(0)

    for dpad_indices in combinations(candidate_indices, 4):
        remaining = [i for i in candidate_indices if i not in dpad_indices]

        for face_indices in combinations(remaining, 4):
            dpad_group = [candidates[i] for i in dpad_indices]
            face_group = [candidates[i] for i in face_indices]

            dpad_rms = cluster_layout_rms(reference_groups["dpad"], dpad_group)
            face_rms = cluster_layout_rms(reference_groups["face_buttons"], face_group)

            if dpad_rms is None or face_rms is None:
                continue

            dpad_center = center_of_group(dpad_group)
            face_center = center_of_group(face_group)

            dpad_position_error = _centroid_distance_xz_mm(dpad_center, targets["dpad"])
            face_position_error = _centroid_distance_xz_mm(face_center, targets["face_buttons"])

            # Shape determines the cluster role; position is used
            # as a secondary feature to distinguish between placement alternatives.
            cost = dpad_rms + face_rms + 0.05 * (dpad_position_error + face_position_error) / TARGET_WIDENING_MM

            if best is None or cost < best["cost"]:
                best = {
                    "cost": cost,
                    "dpad_group": dpad_group,
                    "face_group": face_group,
                    "dpad_rms": dpad_rms,
                    "face_rms": face_rms,
                    "dpad_center": dpad_center,
                    "face_center": face_center,
                    "dpad_position_error": dpad_position_error,
                    "face_position_error": face_position_error,
                }

    if best is None:
        return {
            "score": 0.0,
            "status": "no_valid_partition",
            "candidate_count": len(candidates),
            "candidates": candidates,
            "reason": "Could not find two suitable four-body groups.",
        }

    def shape_score(rms_mm):
        """Full score for RMS up to 0.5 mm; zero score for RMS of 3 mm or more."""
        if rms_mm <= 0.5:
            return 1.0
        return max(0.0, 1.0 - (rms_mm - 0.5) / 2.5)

    def placement_score(error_mm):
        """Allow 1 mm; beyond that, reduce the score linearly to zero at 15 mm."""
        if error_mm <= 1.0:
            return 1.0
        return max(0.0, 1.0 - (error_mm - 1.0) / (TARGET_WIDENING_MM - 1))

    details = {}

    for role, group_key, rms_key, error_key, center_key in (
            ("dpad", "dpad_group", "dpad_rms", "dpad_position_error", "dpad_center"),
            ("face_buttons", "face_group", "face_rms", "face_position_error", "face_center"),
    ):
        rms = best[rms_key]
        position_error = best[error_key]
        group_shape_score = shape_score(rms)
        group_placement_score = placement_score(position_error)

        details[role] = {
            "bodies": best[group_key],
            "layout_rms_mm": rms,
            "position_error_mm": position_error,
            "shape_score": group_shape_score,
            "placement_score": group_placement_score,
            "score": group_shape_score * group_placement_score,
            "target_center_mm": tuple(value * 1000.0 for value in targets[role]),
            "current_center_mm": tuple(value * 1000.0 for value in best[center_key]),
        }

    return {
        "score": (details["dpad"]["score"] + details["face_buttons"]["score"]) / 2.0,
        "status": "ok",
        "candidate_count": len(candidates),
        "details": details,
        "unused_candidates": [
            body for body in candidates
            if body not in best["dpad_group"] and body not in best["face_group"]
        ],
    }


def print_button_clusters_score(result):
    """Prints the button-cluster matches and scores."""

    print_if_debug("\n" + "=" * 80)
    print_if_debug("BUTTON-LAYOUT SCORE")
    print_if_debug("=" * 80)

    if result["status"] != "ok":
        print_if_debug(f"Score unavailable: {result['reason']}")
        print_if_debug(f"Candidates: {result['candidate_count']}")
        return

    for role, details in result["details"].items():
        print_if_debug(f"\n{role}:")

        for body in details["bodies"]:
            print_if_debug(f"  #{body['index']:02d} {body_name(body)}")

        print_if_debug(f"  Internal-layout RMS: {details['layout_rms_mm']:.3f} mm")
        print_if_debug(f"  Position error: {details['position_error_mm']:.3f} mm")
        print_if_debug(f"  Shape score: {details['shape_score']:.3f}")
        print_if_debug(f"  Placement score: {details['placement_score']:.3f}")
        print_if_debug(f"  Group score: {details['score']:.3f}")

        target = details["target_center_mm"]
        current = details["current_center_mm"]
        print_if_debug(f"  Target centroid: ({target[0]:.2f}, {target[1]:.2f}, {target[2]:.2f}) mm")
        print_if_debug(f"  Current centroid: ({current[0]:.2f}, {current[1]:.2f}, {current[2]:.2f}) mm")

    print_if_debug(f"\nOverall button-layout score: {result['score']:.3f} / 1.000")
    print_if_debug(f"Unused candidates: {len(result['unused_candidates'])}")


def score_centre_positions(baseline, matches, full_error_mm=1.0, zero_error_mm=15.0):
    """Scores the mirrored placement of the side center components (START/SELECT) and the position of the central PS button."""
    reference_by_id = {body["id"]: body for body in baseline["bodies"]}
    plane_x = baseline["plane_x_m"]
    centre_ids = baseline["roles"].get("centre", [])

    if not centre_ids:
        return {"score": 0.0, "status": "no_reference", "elements": []}

    scores = []
    elements = []

    print_if_debug("\nCenter-component score:")

    for body_id in centre_ids:
        reference = reference_by_id.get(body_id)
        current = matches.get(body_id)

        if reference is None or current is None:
            scores.append(0.0)
            elements.append({"id": body_id, "score": 0.0, "status": "missing"})
            print_if_debug(f"  {body_id}: matching body not found | score=0.000")
            continue

        x, y, z = reference["centroid_m"]

        # The side components must move to the opposite side.
        # The central button must remain on the symmetry axis (plane_x).
        is_central_button = abs(x - plane_x) <= 0.010
        target_x = plane_x if is_central_button else 2.0 * plane_x - x

        # Distance from the current centroid to the mirrored target.
        cx, cy, cz = current["centroid_m"]

        # This criterion checks only X; it does not require an exact match
        # in Y and Z, as the task does not specify those constraints separately.
        error_mm = abs(cx - target_x) * 1000.0

        if error_mm <= full_error_mm:
            body_score = 1.0
        elif error_mm >= zero_error_mm:
            body_score = 0.0
        else:
            body_score = (zero_error_mm - error_mm) / (zero_error_mm - full_error_mm)

        state = "stay" if is_central_button else "mirrored"
        scores.append(body_score)
        elements.append({
            "id": body_id,
            "state": state,
            "position_error_mm": error_mm,
            "score": body_score,
            "status": "ok",
        })

        print_if_debug(f"  {body_id}: target={state} | deviation={error_mm:.2f} mm | score={body_score:.3f}")

    score = statistics.fmean(scores) if scores else 0.0
    status = "partial" if any(item["status"] == "missing" for item in elements) else "ok"

    print_if_debug(f"Center-component score: {score:.3f} / 1.000")

    return {"score": score, "status": status, "elements": elements}


def diagnose_light_cluster(baseline, records, radius_mm=25.0, bodies_amount=1):
    """Searches the three largest bodies for the indicator cluster near the expected
    or mirrored position, using the baseline face-area signature."""

    reference = baseline["light_cluster"]
    target_areas = reference["areas_m2"]
    target_centroid = reference["centroid_m"]
    plane_x = baseline["plane_x_m"]
    radius_m = radius_mm / 1000.0

    # The cluster is part of the housing. Instead of scanning every component,
    # rank solid bodies by volume and inspect only the N largest.
    candidate_bodies = sorted(records, key=body_volume_mm3, reverse=True)[:bodies_amount]

    # Check both the cluster’s original and mirrored positions.
    # The second search region helps detect an incorrect mirror operation.
    search_regions = [
        ("expected", target_centroid),
        ("mirrored", (2.0 * plane_x - target_centroid[0], target_centroid[1], target_centroid[2])),
    ]

    best = None

    for body_rank, candidate_info in enumerate(candidate_bodies, start=1):
        body_index = candidate_info['index']

        # Retrieve the face geometry of each selected body only once.
        faces = []

        for face_index, face in enumerate(candidate_info['raw_faces']):
            area = face.GetArea
            box = face.GetBox

            if area <= 0 or box is None or len(box) < 6:
                continue

            # box = [float(value) for value in box]
            center = ((box[0] + box[3]) / 2, (box[1] + box[4]) / 2, (box[2] + box[5]) / 2)
            faces.append({"body_index": body_index, "face_index": face_index, "area": area, "center": center})

        # Each body requires only two matching attempts, rather than one
        # attempt for every possible face center.
        for region_name, region_center in search_regions:
            # For each possible center, find a local region and match
            # its face areas one-to-one against the baseline areas.
            local = [face for face in faces
                     if all(abs(face["center"][axis] - region_center[axis]) <= radius_m
                            for axis in range(3))]

            if len(local) < len(target_areas):
                # print_if_debug(f"\nlight_cluster: too few faces in the search region ({len(local)} of {len(target_areas)}).")
                continue

            # Match the baseline and measured face areas.
            # The logarithmic error measures relative changes in area.
            matrix = [[abs(math.log(face["area"] / area)) for face in local] for area in target_areas]

            assignment = _hungarian_assignment(matrix)

            if len(assignment) != len(target_areas):
                continue

            selected = [local[assignment[i]] for i in range(len(target_areas))]
            errors = [matrix[i][assignment[i]] for i in range(len(target_areas))]
            mean_error = statistics.fmean(errors)  # sum / len

            weights = [face["area"] for face in selected]
            total_area = sum(weights)

            if total_area <= 0:
                continue

            centroid = tuple(
                sum(face["center"][axis] * weight for face, weight in zip(selected, weights)) / total_area
                for axis in range(3)
            )

            centroid_error_m = math.sqrt(sum((centroid[axis] - region_center[axis]) ** 2 for axis in range(3)))

            # First, select the best match by face-area similarity.
            # Break ties in favor of the centroid closest to the target position.
            candidate_key = (mean_error, centroid_error_m)

            if best is None or candidate_key < best["candidate_key"]:
                sorted_errors = sorted(errors)
                p90_index = math.ceil(0.90 * len(sorted_errors)) - 1

                best = {
                    "found": True,
                    "mean_log_error": mean_error,
                    "p90_log_error": sorted_errors[p90_index],
                    "centroid_m": centroid,
                    "side": 1 if centroid[0] > plane_x else -1,
                    "matched_faces": selected,
                    "matched_count": len(selected),
                    "candidate_key": candidate_key,
                    "candidate_body_index": body_index,
                    "candidate_body_rank": body_rank,
                    "candidate_region": region_name,
                    "local_face_count": len(local),
                }

    if best is None:
        print_if_debug("\nlight_cluster: no cluster found near either target position in the three largest bodies.")
        return {"found": False, "reason": "no_matching_cluster_in_top_bodies"}

    cx, cy, cz = (value * 1000 for value in best["centroid_m"])
    tx, ty, tz = (value * 1000 for value in target_centroid)
    mean_relative_error = math.expm1(best["mean_log_error"]) * 100.0

    print_if_debug("\n" + "=" * 80)
    print_if_debug("LIGHT-CLUSTER DIAGNOSTICS")
    print_if_debug("=" * 80)
    print_if_debug(f"Largest bodies inspected: {len(candidate_bodies)}")
    print_if_debug(f"Best body: #{best['candidate_body_index']} (volume rank {best['candidate_body_rank']})")
    print_if_debug(f"Search region: {best['candidate_region']}")
    print_if_debug(f"Local candidate faces: {best['local_face_count']}")
    print_if_debug(f"Matched faces: {best['matched_count']} / {len(target_areas)}")
    print_if_debug(f"Mean multiplicative face-area deviation: {mean_relative_error:.2f}%")
    print_if_debug(f"P90 logarithmic error: {best['p90_log_error']:.4f}")
    print_if_debug(f"Baseline centroid: ({tx:.2f}, {ty:.2f}, {tz:.2f}) mm")
    print_if_debug(f"Detected centroid: ({cx:.2f}, {cy:.2f}, {cz:.2f}) mm")
    print_if_debug(f"Side relative to X={plane_x * 1000:.2f} mm: {best['side']:+d}")
    print_if_debug(f"Expected side: {reference['side']:+d}")

    best.pop("candidate_key", None)
    return best


def score_light_cluster(result, baseline, full_error=0.03, zero_error=0.15,
                        full_position_error_mm=5.0, zero_position_error_mm=25.0):
    """Scores the indicator cluster’s position and which side of the symmetry plane it occupies."""

    if not result.get("found"):
        return {
            "score": 0.0,
            "status": "not_found",
            "side_ok": False,
            "position_score": 0.0,
        }

    reference = baseline["light_cluster"]
    expected_side = reference["side"]
    side_ok = result["side"] == expected_side

    # Compare the detected cluster position with its original center.
    target = reference["centroid_m"]
    current = result["centroid_m"]

    position_error_mm = math.sqrt(sum((current[axis] - target[axis]) ** 2 for axis in range(3))) * 1000.0

    if position_error_mm <= full_position_error_mm:
        position_score = 1.0
    elif position_error_mm >= zero_position_error_mm:
        position_score = 0.0
    else:
        position_score = ((zero_position_error_mm - position_error_mm)
                          / (zero_position_error_mm - full_position_error_mm))

    # Face-area error remains a diagnostic measure of match quality.
    # It does not zero the score when the cluster is on the correct side and in the correct position.
    mean_error = math.expm1(result["mean_log_error"])

    if mean_error <= full_error:
        shape_score = 1.0
    elif mean_error >= zero_error:
        shape_score = 0.0
    else:
        shape_score = (zero_error - mean_error) / (zero_error - full_error)

    score = position_score if side_ok else 0.0

    if not side_ok:
        status = "wrong_side"
    elif position_score == 0.0:
        status = "position_error"
    elif mean_error > zero_error:
        status = "warning_signature_mismatch"
    else:
        status = "ok"

    return {
        "score": score,
        "status": status,
        "side_ok": side_ok,
        "position_error_mm": position_error_mm,
        "position_score": position_score,
        "shape_score": shape_score,
        "mean_relative_error": mean_error,
    }


def print_light_cluster_score(light_score):
    print_if_debug(f"\nHousing marking score: {light_score['score']:.3f} / 1.000")
    print_if_debug(f"Status: {light_score['status']}")


def score_glyph_preservation(baseline, matches):
    """
    Scores whether geometric detail is preserved on the four face buttons.

    In the tested models, buttons without symbols have 4 faces, while buttons with symbols have at least 8.
    """
    scores = []

    print_if_debug("\nFace-button symbol preservation:")

    for ref_id in baseline["roles"]["face_buttons"]:
        candidate = matches.get(ref_id)

        if candidate is None:
            print_if_debug(f"  {ref_id}: body not found | score=0.000")
            scores.append(0.0)
            continue

        face_count = candidate.get("face_count")

        if face_count is None:
            print_if_debug(f"  {ref_id}: face-count data unavailable | score=0.000")
            scores.append(0.0)
            continue

        # 4 faces = no symbol detail.
        # 8 or more faces = full credit for this geometric feature.
        # The score increases linearly between 4 and 8 faces.
        body_score = max(0.0, min(1.0, (face_count - 4) / 4.0))
        scores.append(body_score)

        print_if_debug(
            f"  {ref_id} -> #{candidate['index']:02d} {candidate['name_debug']:22}: "
            f"faces={face_count} | score={body_score:.3f}"
        )

    score = sum(scores) / len(scores) if scores else 0.0
    print_if_debug(f"Overall symbol-preservation score: {score:.3f}")

    return score


def diagnose_control_interference(baseline, records, sheet_bodies, matches):
    """Checks pairwise intersections among the 18 primary bodies assigned baseline roles."""
    bodies = [r["raw_body"] for r in records]

    print_if_debug("\nHousing-body check:")
    print_if_debug(f"  Solid bodies: {len(bodies)}")
    print_if_debug(f"  Sheet bodies: {len(sheet_bodies)}")

    solid_summary = []
    for index, body in enumerate(bodies):
        props = [float(v) for v in body.GetMassProperties(1.0)]
        box = [float(v) for v in body.GetBodyBox()]
        dims_mm = [(box[i + 3] - box[i]) * 1000 for i in range(3)]
        solid_summary.append((props[3], index, props[4], dims_mm, body.GetFaceCount()))

    for volume_m3, index, area_m2, dims_mm, face_count in sorted(solid_summary, reverse=True)[:8]:
        print_if_debug(f"  solid #{index:02d}: V={volume_m3 * 1e9:.1f} mm³, A={area_m2 * 1e6:.1f} mm², "
                       f"bbox={dims_mm[0]:.1f}*{dims_mm[1]:.1f}*{dims_mm[2]:.1f} mm, faces={face_count}")

    for index, body in enumerate(sheet_bodies):
        props = [float(v) for v in body.GetMassProperties(1.0)]
        box = [float(v) for v in body.GetBodyBox()]
        dims_mm = [(box[i + 3] - box[i]) * 1000 for i in range(3)]
        print_if_debug(f"  sheet #{index:02d}: A={props[3] * 1e6:.1f} mm², "
                       f"bbox={dims_mm[0]:.1f}*{dims_mm[1]:.1f}*{dims_mm[2]:.1f} mm, "
                       f"faces={body.GetFaceCount()}")

    selected = {}
    ref_by_index = {}
    missing = []

    # Collect the housing and all bodies assigned task roles; unclassified bodies are excluded.
    reference_ids = [ref_id for ids in baseline["roles"].values() for ref_id in ids]

    for ref_id in reference_ids:
        candidate = matches.get(ref_id)

        if candidate is None:
            missing.append(ref_id)
            continue

        index = int(candidate["index"])
        if not 0 <= index < len(bodies):
            missing.append(ref_id)
            continue

        selected[index] = bodies[index]
        ref_by_index[index] = ref_id

    overlaps = []
    errors = []

    for (index_a, body_a), (index_b, body_b) in combinations(sorted(selected.items()), 2):
        try:
            volume_m3 = body_intersection_volume_m3(body_a, body_b)
        except Exception as exc:
            errors.append((index_a, index_b, str(exc)))
            continue

        # Discard only numerically negligible volumes; this is not a geometric tolerance in mm.
        if volume_m3 > 1e-15:
            overlaps.append((ref_by_index[index_a], ref_by_index[index_b], volume_m3))

    tested_pairs = len(selected) * (len(selected) - 1) // 2
    total_volume_m3 = sum(item[2] for item in overlaps)

    print_if_debug("\nPrimary-body intersections:")
    print_if_debug(f"  Bodies matched: {len(selected)} of {len(reference_ids)}")
    print_if_debug(f"  Pairs checked: {tested_pairs} (baseline: {baseline['interference']['tested_pairs']})")
    print_if_debug(f"  Intersecting pairs: {len(overlaps)}")
    print_if_debug(f"  Sum of pairwise intersection volumes: {total_volume_m3 * 1e9:.6f} mm³")

    for ref_a, ref_b, volume_m3 in overlaps:
        print_if_debug(f"    {ref_a} / {ref_b}: {volume_m3 * 1e9:.6f} mm³")

    for index_a, index_b, message in errors:
        print_if_debug(f"    ERROR #{index_a}/#{index_b}: {message}")

    if missing:
        print_if_debug(f"  Could not match bodies for: {', '.join(missing)}")

    return {
        "tested_pairs": tested_pairs,
        "overlaps": overlaps,
        "total_volume_m3": total_volume_m3,
        "errors": errors,
        "missing": missing,
    }


def score_control_interference(result, baseline, warning_volume_mm3=5.0, decay_volume_mm3=1000.0):
    """Scores intersections among the primary bodies; small contacts trigger a warning and a minor penalty."""
    expected_pairs = int(baseline["interference"]["tested_pairs"])

    # If not all expected pairs were checked or Boolean operations failed,
    # the model cannot reliably be declared free of intersections.
    incomplete = (
            result.get("tested_pairs") != expected_pairs
            or bool(result.get("errors"))
            or bool(result.get("missing"))
    )

    if incomplete:
        print_if_debug("\nIntersection score unavailable: the check was incomplete.")
        print_if_debug(f"  Expected pairs: {expected_pairs}; checked: {result.get('tested_pairs')}")
        print_if_debug(f"  Operation errors: {len(result.get('errors', []))}")
        print_if_debug(f"  Unmatched bodies: {len(result.get('missing', []))}")
        return {"score": None, "status": "incomplete", "warnings": []}

    minor_overlaps = []
    significant_overlaps = []

    for ref_a, ref_b, volume_m3 in result["overlaps"]:
        volume_mm3 = max(0.0, float(volume_m3) * 1e9)
        item = {"pair": (ref_a, ref_b), "volume_mm3": volume_mm3}

        if volume_mm3 <= warning_volume_mm3:
            minor_overlaps.append(item)
        else:
            significant_overlaps.append(item)

    # Intersections above the threshold progressively reduce the score.
    # Small intersections incur a penalty of only 2% of this criterion per pair.
    significant_volume = sum(item["volume_mm3"] - warning_volume_mm3 for item in significant_overlaps)

    score = math.exp(-significant_volume / decay_volume_mm3) * (0.98 ** len(minor_overlaps))

    warnings = []

    for item in minor_overlaps:
        message = (f"Minor intersection {item['pair'][0]} / {item['pair'][1]}: {item['volume_mm3']:.3f} mm³")
        warnings.append(message)
        print_if_debug(f"  WARNING: {message}")

    for item in significant_overlaps:
        message = (f"Intersection {item['pair'][0]} / {item['pair'][1]}: {item['volume_mm3']:.3f} mm³")
        warnings.append(message)
        print_if_debug(f"  WARNING: {message}")

    status = "ok" if not warnings else "warning"
    print_if_debug(f"Intersection score: {score:.3f} / 1.000")

    return {
        "score": score,
        "status": status,
        "minor_count": len(minor_overlaps),
        "significant_count": len(significant_overlaps),
        "significant_volume_mm3": significant_volume,
        "warnings": warnings,
    }


# Unused for now
def diagnose_all_solid_intersections(records):
    """Checks intersections among all solid bodies, including extra components and fasteners."""
    bodies = [r["raw_body"] for r in records]
    record_by_index = {int(record["index"]): record for record in records}

    def label(index):
        record = record_by_index.get(index, {})
        name = record.get("name_debug", "unknown")
        return f"#{index:02d} {name}"

    overlaps = []
    errors = []
    total_pairs = len(bodies) * (len(bodies) - 1) // 2
    bbox_candidates = 0

    for (index_a, body_a), (index_b, body_b) in combinations(enumerate(bodies), 2):
        # Run the Boolean operation only for bodies whose bounding boxes overlap.
        if not _body_boxes_may_intersect(body_a, body_b):
            continue

        bbox_candidates += 1

        try:
            volume_m3 = body_intersection_volume_m3(body_a, body_b)
        except Exception as exc:
            errors.append((label(index_a), label(index_b), str(exc)))
            continue

        if volume_m3 > 1e-15:
            overlaps.append((label(index_a), label(index_b), volume_m3))

    print_if_debug("\nIntersections among all solid bodies:")
    print_if_debug(f"  Solid bodies: {len(bodies)}")
    print_if_debug(f"  Total pairs: {total_pairs}")
    print_if_debug(f"  Pairs passing the bounding-box filter: {bbox_candidates}")
    print_if_debug(f"  Intersecting pairs: {len(overlaps)}")
    print_if_debug(f"  Calculation errors: {len(errors)}")

    for name_a, name_b, volume_m3 in overlaps:
        print_if_debug(f"  WARNING: {name_a} / {name_b}: {volume_m3 * 1e9:.6f} mm³")

    for name_a, name_b, message in errors:
        print_if_debug(f"  ERROR: {name_a} / {name_b}: {message}")

    return {
        "body_count": len(bodies),
        "tested_pairs": total_pairs,
        "bbox_candidates": bbox_candidates,
        "overlaps": overlaps,
        "errors": errors,
    }


# Unused for now
def score_all_solid_intersections(result, warning_volume_mm3=5.0, decay_volume_mm3=1000.0):
    """Applies a small penalty for minor intersections and a larger one for significant intersections."""
    if result.get("errors"):
        return {
            "score": None,
            "status": "incomplete",
            "warnings": [],
            "errors": result["errors"],
        }

    minor_count = 0
    significant_volume_mm3 = 0.0
    warnings = []

    for name_a, name_b, volume_m3 in result["overlaps"]:
        volume_mm3 = volume_m3 * 1e9

        if volume_mm3 <= warning_volume_mm3:
            minor_count += 1
            warnings.append(f"Minor intersection {name_a} / {name_b}: {volume_mm3:.3f} mm³")
        else:
            significant_volume_mm3 += volume_mm3 - warning_volume_mm3
            warnings.append(f"Intersection {name_a} / {name_b}: {volume_mm3:.3f} mm³")

    score = (0.98 ** minor_count) * math.exp(-significant_volume_mm3 / decay_volume_mm3)

    for warning in warnings:
        print_if_debug(f"  WARNING: {warning}")

    print_if_debug(f"All-body intersection score: {score:.3f} / 1.000")

    return {
        "score": score,
        "status": "warning" if warnings else "ok",
        "minor_count": minor_count,
        "significant_volume_mm3": significant_volume_mm3,
        "warnings": warnings,
        "errors": [],
    }


# Unscored for now
def diagnose_unmatched_housing_sized_bodies(baseline, records, matches):
    """Checks for intersections involving additional bodies roughly the size of the housing."""
    bodies = [r["raw_body"] for r in records]
    assigned_indices = {
        int(candidate["index"])
        for candidate in matches.values()
        if candidate is not None
    }

    matched_refs = [
        (role, ref_id, matches.get(ref_id))
        for role, ref_ids in baseline["roles"].items()
        for ref_id in ref_ids
    ]

    candidates = []

    for index, body in enumerate(bodies):
        if index in assigned_indices:
            continue

        props = [float(value) for value in body.GetMassProperties(1.0)]
        box = [float(value) for value in body.GetBodyBox()]
        dx = (box[3] - box[0]) * 1000
        dz = (box[5] - box[2]) * 1000
        volume_mm3 = props[3] * 1e9

        if dx >= 200 and dz >= 120 and volume_mm3 >= 50000:
            candidates.append((index, body, volume_mm3, dx, dz))

    print_if_debug("\nIntersections involving additional large bodies:")

    if not candidates:
        print_if_debug("  No unassigned bodies comparable in size to the housing were found.")
        return

    for index, body, volume_mm3, dx, dz in candidates:
        print_if_debug(f"  #{index:02d}: V={volume_mm3:.1f} mm³, X={dx:.1f} mm, Z={dz:.1f} mm")

        tested = 0
        overlaps = 0

        for role, ref_id, candidate in matched_refs:
            if candidate is None:
                continue

            other_index = int(candidate["index"])
            if other_index == index:
                continue

            try:
                volume = body_intersection_volume_m3(body, bodies[other_index])
                tested += 1
            except Exception as exc:
                print_if_debug(f"    ERROR for {ref_id}: {exc}")
                continue

            if volume > 1e-15:
                overlaps += 1
                print_if_debug(f"    Intersection with {role}/{ref_id}: {volume * 1e9:.6f} mm³")

        if overlaps == 0:
            print_if_debug(f"    No intersections among the checked pairs ({tested}).")


def calculate_total_score(tree, widening, button_layout, centre, light, glyph, interference):
    """Combines the individual criteria into a final score out of 4 points."""

    # The agreed weighting scheme; all criterion weights sum to 4 points.
    weights = {
        "Feature tree (⍋)": 0.45,
        "Widening (↔)": 1.20,
        "Button clusters (|)": 1.10,
        "Center components (|)": 0.35,
        "Housing marking (R)": 0.55,
        "Button symbols (✗○☐△)": 0.25,
        "Intersections (∩)": 0.10,
    }

    components = [
        ("Feature tree (⍋)", tree.get("score")),
        ("Widening (↔)", widening.get("score")),
        ("Button clusters (|)", button_layout.get("score")),
        ("Center components (|)", centre.get("score")),
        ("Housing marking (R)", light.get("score")),
        ("Button symbols (✗○☐△)", glyph),
        ("Intersections (∩)", interference.get("score")),
    ]

    maximum_points = sum(weights.values())
    earned_points = 0.0
    unavailable = []
    details = {}

    print_if_debug("\n" + "=" * 64)
    print_if_debug("FINAL SCORE")
    print_if_debug("=" * 64)

    for name, raw_score in components:
        weight = weights[name]

        if raw_score is None or not math.isfinite(float(raw_score)):
            unavailable.append({"name": name, "weight": weight})

            # If intersections could not be checked because the model is damaged,
            # assign 0 points for this criterion while retaining the warning.
            points = 0.0 if name == "Intersections (∩)" else None
            details[name] = {"score": None, "weight": weight, "points": points}

            if name == "Intersections (∩)":
                print_if_debug(f"  {name:<26} N/A -> 0.000 (conservative fallback, weight {weight:.2f})")
            else:
                print_if_debug(f"  {name:<26} N/A   (weight {weight:.2f})")

            continue

        normalized_score = max(0.0, min(1.0, float(raw_score)))
        points = normalized_score * weight
        earned_points += points

        details[name] = {"score": normalized_score, "weight": weight, "points": points}
        print_if_debug(f"  {name:<26} {normalized_score:.3f} * {weight:.2f} = {points:.3f}")

    if unavailable and all(item["name"] == "Intersections (∩)" for item in unavailable):
        final_score = earned_points
        status = "conservative_incomplete"
        print_if_debug("\nWARNING: intersections could not be checked completely.")
        print_if_debug(f"TOTAL (conservative): {final_score:.3f} / {maximum_points:.2f}")
    elif unavailable:
        unavailable_weight = sum(item["weight"] for item in unavailable)
        print_if_debug(f"\nPartial total: {earned_points:.3f} / {maximum_points:.2f}")
        print_if_debug(f"Final score not finalized: criteria worth {unavailable_weight:.2f} points are unavailable.")
        final_score = None
        status = "incomplete"
    else:
        final_score = earned_points
        status = "ok"
        print_if_debug(f"\nTOTAL: {final_score:.3f} / {maximum_points:.2f}")

    passed = final_score is not None and final_score >= 3.90

    return {
        "score": final_score,
        "earned_points": earned_points,
        "max_points": maximum_points,
        "status": status,
        "passed": passed,
        "unavailable": unavailable,
        "components": details,
    }
...  # ==================== /Scoring =========================


# ========================= Helpers =========================
def print_if_debug(*text, end="\n"):
    if DEBUG_FULL_OUTPUT:
        print(*text, end=end)


def get_widening_pairs(baseline):
    """
    Identifies mirrored pairs for the widening criterion.

    Read the stick, trigger, and bumper pairs directly from the baseline geometric annotations.
    """

    pairs = {}

    for role in ("sticks", "triggers", "bumpers"):
        body_ids = baseline["roles"].get(role, [])

        if len(body_ids) == 2:
            pairs[role] = tuple(body_ids)

    return pairs


def _relative_log_error(actual, expected):
    """
    Computes the relative difference between two positive values.
    The logarithm of the ratio allows bodies of different sizes to be compared.
    """
    actual = max(float(actual), 1e-30)
    expected = max(float(expected), 1e-30)
    return abs(math.log(actual / expected))


def _body_dimensions(body):
    """
    Returns the three dimensions of the bounding box.
    Sorting makes the result independent of axis order.
    The body centroid is intentionally ignored here.
    """
    box = body["bbox_m"]
    return sorted([
        box[3] - box[0],
        box[4] - box[1],
        box[5] - box[2],
    ])


def body_shape_distance(reference, candidate):
    """
    Scores the shape difference between a baseline body and a candidate body.
    Uses body volume, surface area, and dimensions.
    """

    volume_error = _relative_log_error(candidate["volume_m3"], reference["volume_m3"])

    area_error = _relative_log_error(candidate["area_m2"], reference["area_m2"])

    ref_dims = _body_dimensions(reference)
    cand_dims = _body_dimensions(candidate)

    dimensions_error = sum(_relative_log_error(current, expected) for current, expected in zip(cand_dims, ref_dims)) / 3

    return volume_error + 0.7 * area_error + 0.7 * dimensions_error


def get_transform_hypotheses(reference, role, baseline):
    """
    Generates position hypotheses for baseline bodies during global matching.

    The hypotheses affect body assignments and indirectly affect the final
    geometric scores, but are not a standalone scoring criterion.
    """

    x, y, z = reference["centroid_m"]
    plane_x = baseline["plane_x_m"]

    def make_hypothesis(state, target_x, target_y=y, target_z=z):
        return {"state": state, "centroid_m": [target_x, target_y, target_z]}

    # Treat unclassified bodies separately as bodies that remain
    # in their original positions. The housing is handled by a different method.
    if role == "housing":
        return []

    if role == "unclassified":
        return [make_hypothesis("original", x)]

    # The D-pad and face buttons move as complete clusters.
    if role in ("dpad", "face_buttons"):
        reference_by_id = {body["id"]: body for body in baseline["bodies"]}

        group_ids = baseline["roles"][role]
        group_center_x = sum(
            reference_by_id[body_id]["centroid_m"][0]
            for body_id in group_ids
        ) / len(group_ids)

        # This body’s offset from the center of its cluster.
        # The offset is preserved when the entire group is moved.
        local_offset_x = x - group_center_x

        source_direction = (-1.0 if group_center_x < plane_x else 1.0)

        # Move the cluster to the opposite side.
        swapped_center_x = 2.0 * plane_x - group_center_x

        swapped_direction = (1.0 if swapped_center_x > plane_x else -1.0)

        return [
            make_hypothesis("original", x),
            make_hypothesis("widened", x + source_direction * HALF_TARGET_WIDENING_M),
            make_hypothesis("swapped", swapped_center_x + local_offset_x),
            make_hypothesis("swapped+widened", swapped_center_x + local_offset_x +
                            swapped_direction * HALF_TARGET_WIDENING_M),
        ]

    # START/SELECT may swap sides.
    # The central PS button must remain on the symmetry axis.
    if role == "centre":
        if abs(x - plane_x) <= 0.010:
            return [make_hypothesis("original", x)]

        return [make_hypothesis("original", x), make_hypothesis("mirrored", 2.0 * plane_x - x)]

    # The sticks move outward with their respective sides.
    if role == "sticks":
        direction = 1.0 if x > plane_x else -1.0

        return [
            make_hypothesis("original", x),
            make_hypothesis("widened", x + direction * HALF_TARGET_WIDENING_M),
        ]

    # For triggers and bumpers, check not only widening
    # but also whether their assignments have been swapped on each side.
    if role in ("triggers", "bumpers"):
        direction = 1.0 if x > plane_x else -1.0

        other_role = ("bumpers" if role == "triggers" else "triggers")

        reference_by_id = {body["id"]: body for body in baseline["bodies"]}

        # Find the corresponding element in the other group
        # on the same side of the housing.
        counterpart_ids = baseline["roles"].get(other_role, [])

        if counterpart_ids:
            counterpart_id = min(
                counterpart_ids,
                key=lambda body_id: abs(
                    reference_by_id[body_id]["centroid_m"][0] - x
                ),
            )

            counterpart_xyz = reference_by_id[counterpart_id]["centroid_m"]

            cx, cy, cz = counterpart_xyz

            return [
                make_hypothesis("original", x),
                make_hypothesis("widened", x + direction * HALF_TARGET_WIDENING_M),
                make_hypothesis("assignment_swapped", cx, cy, cz),
                make_hypothesis("assignment_swapped+widened", cx + direction * HALF_TARGET_WIDENING_M, cy, cz),
            ]

        return [make_hypothesis("original", x)]

    return [make_hypothesis("original", x)]


def cluster_layout_rms(reference_bodies, current_bodies, return_details=False):
    """
    Compares the relative arrangement of the four bodies in a cluster.

    Only the X and Z centroid coordinates are used. Before comparison,
    subtract the cluster center so translation within the model does not affect the result.

    Evaluate all 4! permutations of the bodies and select the one for which
    the root-mean-square distance between relative coordinates
    is minimized. Return distances in millimeters.
    """

    if len(reference_bodies) != 4 or len(current_bodies) != 4:
        return None

    ref_points = [(body["centroid_m"][0], body["centroid_m"][2]) for body in reference_bodies]

    cur_points = [(body["centroid_m"][0], body["centroid_m"][2]) for body in current_bodies]

    ref_center = (sum(point[0] for point in ref_points) / 4, sum(point[1] for point in ref_points) / 4)

    cur_center = (sum(point[0] for point in cur_points) / 4, sum(point[1] for point in cur_points) / 4)

    ref_local = [(x - ref_center[0], z - ref_center[1]) for x, z in ref_points]

    cur_local = [(x - cur_center[0], z - cur_center[1]) for x, z in cur_points]

    best_rms = float("inf")
    best_perm = None

    for perm in permutations(range(4)):
        squared_errors = [(ref_local[i][0] - cur_local[perm[i]][0]) ** 2
                          + (ref_local[i][1] - cur_local[perm[i]][1]) ** 2
                          for i in range(4)]

        rms_m = (sum(squared_errors) / 4) ** 0.5
        if rms_m < best_rms:
            best_rms = rms_m
            best_perm = perm

    if return_details:
        return {
            "rms_mm": best_rms * 1000.0,
            "permutation": best_perm,
            "matched_bodies": [current_bodies[i] for i in best_perm],
        }

    return best_rms * 1000.0


def find_button_clusters(records, plane_x_m, y_range=12.0, z_range=35.0, v_min=500.0, v_max=10_000.0):
    """
    Identifies two four-body button clusters based on geometry.
    Splits the result into left and right clusters relative to the model’s center plane.

    Filters were chosen manually based on dimensions and measurements.
    """

    # Find bodies in the button region: small solid bodies
    # near the Y = 0 plane and within a limited Z range.
    plane_x_mm = plane_x_m * 1000.0
    candidates = []

    for body in records:
        x, y, z = (value * 1000.0 for value in body["centroid_m"])
        volume_mm3 = body["volume_m3"] * 1e9

        # Select small bodies in the button region.
        if not (-y_range <= y <= y_range):
            continue
        if not (-z_range <= z <= z_range):
            continue
        if not (v_min <= volume_mm3 <= v_max):
            continue

        # Exclude the central housing region and center buttons.
        if abs(x - plane_x_mm) < 30.0:
            continue

        candidates.append(body)

    clusters = {"left": [body for body in candidates if body["centroid_m"][0] * 1000.0 < plane_x_mm],
                "right": [body for body in candidates if body["centroid_m"][0] * 1000.0 > plane_x_mm]}

    return clusters


def classify_button_clusters(baseline, clusters):
    """
    Determines which current cluster corresponds to the baseline D-pad and which corresponds to the face buttons.

    Compare the bodies’ internal spatial arrangement.
    Choose assignments jointly for both clusters; the same cluster cannot be assigned to both roles.
    """

    if len(clusters["left"]) != 4 or len(clusters["right"]) != 4:
        return {
            "confident": False,
            "reason": (
                "Expected four bodies in each cluster; found:"
                f" left={len(clusters['left'])},"
                f" right={len(clusters['right'])}"
            ),
        }

    reference_by_id = {body["id"]: body for body in baseline["bodies"]}

    reference_groups = {role: [reference_by_id[body_id] for body_id in baseline["roles"][role]]
                        for role in ("dpad", "face_buttons")}

    # Identify the candidate roles and store their indices for later use.
    layout_details = {
        role: {
            # For each role, measure how closely its baseline layout matches the left and right current clusters.
            side: cluster_layout_rms(reference_groups[role], clusters[side], return_details=True)
            for side in ("left", "right")
        }
        for role in ("dpad", "face_buttons")
    }

    errors = {
        role: {
            side: layout_details[role][side]["rms_mm"]
            for side in ("left", "right")
        }
        for role in ("dpad", "face_buttons")
    }

    # Evaluate both possible assignments.
    assignments = [
        {
            "mapping": {"dpad": "left", "face_buttons": "right"},
            "total_rms_mm": (errors["dpad"]["left"] + errors["face_buttons"]["right"]),
        },
        {
            "mapping": {"dpad": "right", "face_buttons": "left"},
            "total_rms_mm": (errors["dpad"]["right"] + errors["face_buttons"]["left"]),
        },
    ]

    assignments.sort(key=lambda item: item["total_rms_mm"])

    # Keep the runner-up so the difference can be calculated and confidence can be assessed
    best = assignments[0]
    second = assignments[1]

    # This threshold is provisional and should be validated against negative test models.
    margin_mm = second["total_rms_mm"] - best["total_rms_mm"]
    confident = (best["total_rms_mm"] <= 2.0 and margin_mm >= 1.0)

    # When classification is confident, record the current body matched to each baseline button.
    button_matches = {}
    if confident:
        for role, side in best["mapping"].items():
            permutation = layout_details[role][side]["permutation"]
            current_cluster = clusters[side]

            for reference_body, current_index in zip(reference_groups[role], permutation):
                button_matches[reference_body["id"]] = current_cluster[current_index]

    return {
        "confident": confident,
        "mapping": best["mapping"] if confident else None,
        "errors": errors,
        "best_total_rms_mm": best["total_rms_mm"],
        "alternative_total_rms_mm": second["total_rms_mm"],
        "margin_mm": margin_mm,
        "assignments": assignments,
        "layout_details": layout_details,
        "button_matches": button_matches if confident else {},
    }


def _hungarian_assignment(cost_matrix):
    """Finds the assignment with the minimum total cost. The matrix must have no more rows than columns."""

    n = len(cost_matrix)
    if not n:
        return {}

    m = len(cost_matrix[0])
    if n > m or any(len(row) != m for row in cost_matrix):
        raise ValueError("Invalid cost-matrix dimensions.")

    u, v, p, way = [0.0] * (n + 1), [0.0] * (m + 1), [0] * (m + 1), [0] * (m + 1)

    for i in range(1, n + 1):
        p[0] = i
        j0 = 0
        min_values = [float("inf")] * (m + 1)
        used = [False] * (m + 1)

        while True:
            used[j0] = True
            i0 = p[j0]
            delta, j1 = float("inf"), 0

            for j in range(1, m + 1):
                if used[j]:
                    continue

                current = cost_matrix[i0 - 1][j - 1] - u[i0] - v[j]

                if current < min_values[j]:
                    min_values[j], way[j] = current, j0

                if min_values[j] < delta:
                    delta, j1 = min_values[j], j

            for j in range(m + 1):
                if used[j]:
                    u[p[j]] += delta
                    v[j] -= delta
                else:
                    min_values[j] -= delta

            j0 = j1
            if p[j0] == 0:
                break

        while j0:
            j1 = way[j0]
            p[j0] = p[j1]
            j0 = j1

    return {p[j] - 1: j - 1 for j in range(1, m + 1) if p[j]}


def _centroid_distance_mm(point_a, point_b):
    """Euclidean distance between two centroids expressed in meters."""
    return math.sqrt(sum((a - b) ** 2 for a, b in zip(point_a, point_b))) * 1000.0


def _centroid_distance_xz_mm(point_a, point_b):
    """Distance between centroids in the X-Z plane, in millimeters."""
    return math.hypot(point_a[0] - point_b[0], point_a[2] - point_b[2]) * 1000.0


def _body_boxes_may_intersect(body_a, body_b, padding_m=1e-5):
    """Quickly rejects body pairs whose bounding boxes do not overlap."""
    box_a = [float(v) for v in body_a.GetBodyBox()]
    box_b = [float(v) for v in body_b.GetBodyBox()]

    for axis in range(3):
        overlap = min(box_a[axis + 3], box_b[axis + 3]) - max(box_a[axis], box_b[axis])
        if overlap < -padding_m:
            return False

    return True


def body_intersection_volume_m3(body_a, body_b):
    """Returns the intersection volume of two bodies in m³ without modifying the original geometry."""
    if not _body_boxes_may_intersect(body_a, body_b):
        return 0.0

    # The Boolean operation is performed on temporary body copies.
    temp_a = body_a.Copy()
    temp_b = body_b.Copy()

    # PyWin32 returns the error_code out parameter as the second tuple element.
    intersection_bodies, error_code = temp_a.Operations2(SW_BODY_INTERSECT, temp_b)
    error_code = int(error_code)
    # Known codes indicating no intersection are not treated as calculation failures.
    if error_code in (5, 6, 1067):  # - NO INTERSECTION
        # Codes: 5 = bodies do not intersect; 6 = empty body; 1067 = no intersection.
        return 0.0

    if error_code != 0:
        # Code 0 indicates success; any other code indicates an error.
        raise RuntimeError(f"SolidWorks Operations2: {error_code=}")

    if intersection_bodies is None:
        return 0.0

    if not isinstance(intersection_bodies, (tuple, list)):
        intersection_bodies = (intersection_bodies,)

    total_volume_m3 = 0.0

    for raw_body in intersection_bodies:
        if raw_body is None:
            continue

        intersection_body = win32com.client.Dispatch(raw_body)
        mass_properties = intersection_body.GetMassProperties(1.0)
        total_volume_m3 += max(0.0, float(mass_properties[3]))

    return total_volume_m3


def get_role_ids(baseline, role):
    """Returns baseline body IDs for a role, including unclassified bodies."""
    if role == "unclassified":
        assigned = {body_id for ids in baseline["roles"].values() for body_id in ids}
        return [body["id"] for body in baseline["bodies"] if body["id"] not in assigned]
    if role not in baseline["roles"]:
        raise KeyError(f"Unknown role: {role}")
    return list(baseline["roles"][role])


def get_reference_bodies(baseline, role=None):
    """Returns all baseline bodies or only the bodies assigned to the specified role."""
    bodies_by_id = {body["id"]: body for body in baseline["bodies"]}
    if role is None:
        return list(baseline["bodies"])
    return [bodies_by_id[body_id] for body_id in get_role_ids(baseline, role)]


def get_matched_bodies_by_role(baseline, matches, role, include_missing=False):
    """Returns current bodies matched to a baseline role, keyed by baseline body ID."""
    result = {}
    for body_id in get_role_ids(baseline, role):
        if body_id in matches:
            result[body_id] = matches[body_id]
        elif include_missing:
            result[body_id] = None
    return result


def find_bodies_by_feature(records, *, centroid_mm=None, centroid_tolerance_mm=5.0, x_range_mm=None, y_range_mm=None,
                           z_range_mm=None, volume_range_mm3=None, area_range_mm2=None, dimensions_range_mm=None,
                           face_count_range=None, small_face_count_range=None, predicate=None):
    """Filters bodies by geometry. Ranges are inclusive; centroid constraints use Euclidean distance."""
    criteria = (centroid_mm, x_range_mm, y_range_mm, z_range_mm, volume_range_mm3, area_range_mm2, dimensions_range_mm,
                face_count_range, small_face_count_range, predicate)
    if not any(value is not None for value in criteria):
        raise ValueError("At least one geometric feature must be specified.")

    def in_range(value, limits):
        return limits is None or limits[0] <= value <= limits[1]

    result = []

    for body in records:
        x, y, z = body_centroid_mm(body)

        if centroid_mm is not None:
            distance = math.sqrt(sum((a - b) ** 2 for a, b in zip((x, y, z), centroid_mm)))
            if distance > centroid_tolerance_mm:
                continue

        if not in_range(x, x_range_mm) or not in_range(y, y_range_mm) or not in_range(z, z_range_mm):
            continue
        if not in_range(body_volume_mm3(body), volume_range_mm3):
            continue
        if not in_range(body_area_mm2(body), area_range_mm2):
            continue

        if dimensions_range_mm is not None:
            dimensions = body_dimensions_mm(body)
            if not all(in_range(value, dimensions_range_mm) for value in dimensions):
                continue

        if face_count_range is not None:
            if "face_count" not in body or not in_range(body["face_count"], face_count_range):
                continue

        if small_face_count_range is not None:
            if "small_face_count" not in body or not in_range(body["small_face_count"], small_face_count_range):
                continue

        if predicate is not None and not predicate(body):
            continue

        result.append(body)

    return result
...  # ==================== /Helpers =========================


# ========================= Dict helpers =========================
def body_name(body):
    """Returns a diagnostic body name for both input.json and the current model."""
    return str(body.get("name_debug", body.get("name", "")))


def body_centroid_mm(body):
    """Body centroid in millimeters."""
    return tuple(value * 1000.0 for value in body["centroid_m"])


def body_volume_mm3(body):
    """Body volume in mm³."""
    return body["volume_m3"] * 1e9


def body_area_mm2(body):
    """Body surface area in mm²."""
    return body["area_m2"] * 1e6


def body_bbox_mm(body):
    """Bounding box in the order xmin, ymin, zmin, xmax, ymax, zmax."""
    return tuple(value * 1000.0 for value in body["bbox_m"])


def body_dimensions_mm(body):
    """The three sorted bounding-box dimensions, in millimeters."""
    bbox = body_bbox_mm(body)
    return tuple(sorted((bbox[3] - bbox[0], bbox[4] - bbox[1], bbox[5] - bbox[2])))
...  # ========================= /Dict helpers =========================


# ========================= Attempts to analyze glyphs =========================
def get_feature_sketch(feature):
    """Attempts to retrieve an ISketch object from a feature or its data."""

    def looks_like_sketch(obj):
        return obj is not None and sw_value(obj, "GetSketchBlockInstanceCount") is not None

    for source in (sw_value(feature, "GetSpecificFeature2"), sw_value(feature, "GetDefinition")):  # if not none
        if source is None:
            return None

        if looks_like_sketch(source):
            return source

        sketch = sw_value(source, "GetSketch")
        if looks_like_sketch(sketch):
            return sketch

        profile = sw_value(source, "ProfileFeature")
        sketch = sw_value(profile, "GetSpecificFeature2")
        if looks_like_sketch(sketch):
            return sketch

    return None


def inspect_sketch_blocks(sketch):
    """Prints sketch blocks, their definitions, notes, and geometry entities."""

    segments = sw_as_list(sw_value(sketch, "GetSketchSegments"))
    blocks = sw_as_list(sw_value(sketch, "GetSketchBlockInstances"))

    print_if_debug(f"    Sketch segments: {len(segments)}; blocks: {len(blocks)}")

    for block in blocks:
        definition = sw_value(block, "Definition")
        block_name = sw_value(block, "Name", "<unnamed>")
        definition_name = sw_value(definition, "Name", "<no definition>")

        counts = {}
        for label, method in (("lines", "GetLineCount"), ("arcs", "GetArcCount"), ("ellipses", "GetEllipseCount"),
                              ("splines", "GetSplineCount"), ("notes", "GetNoteCount")):
            counts[label] = sw_value(definition, method)

        point_count_var = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        counts["splines"] = definition.GetSplineCount(point_count_var)

        note_texts = []
        for note in sw_as_list(sw_value(definition, "GetNotes")):
            text = sw_value(note, "GetText")
            if text:
                note_texts.append(str(text))

        print_if_debug(f"      Block {block_name!r}, definition {definition_name!r}")
        print_if_debug(f"        Definition geometry: {counts}")
        print_if_debug(f"        Text notes: {note_texts or 'none found'}")


def inspect_feature_faces(feature):
    """Displays the areas of faces created by a feature."""

    faces = sw_as_list(sw_value(feature, "GetFaces"))
    areas = []

    for face in faces:
        area = sw_value(face, "GetArea")
        if area is not None:
            areas.append(float(area) * 1e6)

    areas.sort()

    print_if_debug(f"    Feature faces: {len(faces)}")
    print_if_debug(f"    Face areas (mm²): {[round(a, 4) for a in areas[:30]]}")

    face_boxes = []

    for face in faces:
        box = face.GetBox
        if box is not None and len(box) >= 6:
            x1, y1, z1, x2, y2, z2 = map(float, box[:6])
            face_boxes.append((
                min(x1, x2), min(y1, y2), min(z1, z2),
                max(x1, x2), max(y1, y2), max(z1, z2),
            ))

    if face_boxes:
        xmin = min(box[0] for box in face_boxes)
        ymin = min(box[1] for box in face_boxes)
        zmin = min(box[2] for box in face_boxes)
        xmax = max(box[3] for box in face_boxes)
        ymax = max(box[4] for box in face_boxes)
        zmax = max(box[5] for box in face_boxes)

        print_if_debug(
            f"    Feature-face bounds (mm): "
            f"X=[{xmin * 1000:.2f}, {xmax * 1000:.2f}], "
            f"Y=[{ymin * 1000:.2f}, {ymax * 1000:.2f}], "
            f"Z=[{zmin * 1000:.2f}, {zmax * 1000:.2f}]"
        )


def diagnose_glyph_features(document):
    """
    Finds features and sketches potentially associated with the L/R letters,
    START/SELECT labels, and sketch blocks.
    """

    print_if_debug("\n" + "=" * 90)
    print_if_debug("LABEL AND SKETCH DIAGNOSTICS")
    print_if_debug("=" * 90)

    for feature, depth, name, type_name in iter_feature_tree(document):
        lower_name = name.casefold()
        lower_type = type_name.casefold()

        is_extrude = "boss" in lower_name or "extrude" in lower_name
        is_split_line = "split line" in lower_name or "splitline" in lower_type
        is_target = any(token in lower_name for token in
                        ("Split Line18", "sketch99", "sketch100", "sketch101", "boss-extrude11", "boss-extrude12"))

        if not is_target:
            continue

        is_sketch_feature = "profilefeature" in lower_type or "3dprofile" in lower_type

        sketch = get_feature_sketch(feature) if is_sketch_feature or is_target else None
        block_count = sw_value(sketch, "GetSketchBlockInstanceCount") if sketch is not None else 0

        # Show all extrusion features, split lines,
        # relevant sketches, and any sketches that contain blocks.
        if not (is_extrude or is_split_line or is_target or block_count):
            continue

        print_if_debug(f"\n{'  ' * depth}{name} [{type_name}]")

        if sketch is not None:
            inspect_sketch_blocks(sketch)

        if is_extrude or is_split_line:
            inspect_feature_faces(feature)
...  # ==================== /Attempts to analyze glyphs =========================


# ========================= SolidWorks API =========================
def get_sw_and_doc(document_path):
    """Starts or attaches to SolidWorks and opens the requested document."""
    print_if_debug("Starting / locating SolidWorks...")

    # =============== SolidWorks ===============
    try:
        # Attach to an existing SolidWorks instance.
        sw_app = win32com.client.GetActiveObject("SldWorks.Application")
        if sw_app.Visible:
            print_if_debug("An existing visible SolidWorks instance was found. Closing it will improve performance :)")
    except Exception as exc:
        print_if_debug("No running SolidWorks instance found. Attempting a new headless launch...")
        try:
            # If no instance is running, start a new process
            sw_app = win32com.client.Dispatch("SldWorks.Application")
            print_if_debug("SolidWorks started successfully (or an existing headless instance was found).")
        except Exception as e:
            print_if_debug(f"Error: failed to start SolidWorks. {e}")
            raise RuntimeError("Could not connect to SolidWorks. Verify that it is available on this machine.") from exc

    # =============== Document ===============
    if document_path:
        model_path = Path(document_path).expanduser().resolve()

        if not model_path.is_file():
            raise FileNotFoundError(f"Model file not found: {model_path}")

        if model_path.suffix.lower() != ".sldprt":
            raise ValueError("Expected a SolidWorks part file (*.SLDPRT).")

        # If the requested file is already active, it does not need to be reopened.
        active_doc = sw_app.ActiveDoc
        active_path = (active_doc.GetPathName if active_doc is not None else "")

        # Check whether the file paths match.
        same_path = (active_path and str(Path(active_path).resolve()).casefold() == str(model_path).casefold())

        if same_path:
            document = active_doc
        else:
            # NOTE: 'OpenDoc7' accepts a single document specification object and is easier to use from Python.

            # Create a document specification object.
            open_spec = sw_app.GetOpenDocSpec(str(model_path))

            # Configure the properties on the object.
            open_spec.DocumentType = SW_DOC_PART
            # open_spec.Silent = False

            # Pass the specification object to OpenDoc7 and open the model without dialogs.
            document = sw_app.OpenDoc7(open_spec)

            if document is not None:
                print_if_debug(f"SolidWorks opened the file successfully: '{document.GetTitle}'")
            else:
                # Error details are stored on the specification object.
                raise RuntimeError(f"SolidWorks could not open {model_path}. error code={open_spec.Error}")

            # NOTE: If SolidWorks was launched headlessly, its window can be made visible again:
            # document.Visible = True  # OR sw_app.Visible
    else:
        document = sw_app.ActiveDoc

        if document is None:
            raise RuntimeError("No active document in SolidWorks.")

    print_if_debug(f"Working with document: {document.GetTitle}")

    return sw_app, document


def close_sw(sw_app):
    # Close the current document (the .SLDPRT file name can also be supplied).
    sw_app.CloseDoc(sw_app.ActiveDoc.GetTitle)

    # Close the application itself if this harness launched it headlessly.
    if CLOSE_OUR_SW_IN_THE_END and not sw_app.Visible:
        print_if_debug("\nClosing the headless SolidWorks window...\n")
        sw_app.ExitApp()


def sw_value(obj, prop, default=None):
    """Safely reads a COM property."""
    try:
        value = getattr(obj, prop)
        if callable(value):
            try:
                return value()
            except Exception:
                return value
        return value
    except Exception:
        return default


def sw_as_list(value):
    """Converts a COM array result to a regular Python list."""
    if value is None:
        return []
    if isinstance(value, (tuple, list)):
        return list(value)
    try:
        return list(value)
    except TypeError:
        return [value]


def iter_feature_tree(document):
    """
    Traverses the SolidWorks feature tree, including subfeatures.
    Names are used for diagnostics only, not for scoring.
    """

    seen = set()

    def walk(first, is_sub=False, depth=0):
        feature = first

        while feature is not None:
            name = str(sw_value(feature, "Name", ""))
            type_name = str(sw_value(feature, "GetTypeName2") or "")
            key = (name, type_name)

            if key not in seen:
                seen.add(key)
                yield feature, depth, name, type_name

                sub = sw_value(feature, "GetFirstSubFeature")
                if sub is not None:
                    yield from walk(sub, True, depth + 1)

            next_method = "GetNextSubFeature" if is_sub else "GetNextFeature"
            feature = sw_value(feature, next_method)

    yield from walk(sw_value(document, "FirstFeature"))


def com_array_to_list(value):
    """Converts a SAFEARRAY or array-valued VARIANT to a regular Python list."""
    if hasattr(value, "value"):
        value = value.value
    if value is None:
        return []
    if isinstance(value, (tuple, list)):
        return list(value)
    return [value]


def get_feature_error(feature):
    """Returns a feature error code and whether it is a warning."""
    is_warning = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_BOOL, False)
    error_code = int(feature.GetErrorCode2(is_warning) or 0)
    return error_code, bool(is_warning.value) if error_code else False


def get_whats_wrong(document):
    """Reads the issue list from SolidWorks' What's Wrong pane."""
    extension = document.Extension
    problem_count = int(extension.GetWhatsWrongCount)

    features_out = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_VARIANT, None)
    codes_out = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_VARIANT, None)
    warnings_out = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_VARIANT, None)

    success = extension.GetWhatsWrong(features_out, codes_out, warnings_out)

    if not success:
        return {"success": False, "problem_count": problem_count, "items": []}

    features = com_array_to_list(features_out)
    codes = com_array_to_list(codes_out)
    warnings = com_array_to_list(warnings_out)

    items = []
    for index, feature in enumerate(features):
        error_code = int(codes[index]) if index < len(codes) else 0
        is_warning = bool(warnings[index]) if index < len(warnings) else False

        items.append({
            "name": str(feature.Name),
            "type": str(feature.GetTypeName2),
            "error_code": error_code,
            "is_warning": is_warning,
        })

    return {"success": True, "problem_count": problem_count, "items": items}


def extract_body_geometry(body, index):
    """
    Extracts the geometric properties of a single solid body.

    All values are kept in the units used by the SolidWorks API:
    coordinates in meters, volume in m³, and area in m².
    This allows direct comparison with input.json.
    """
    print_if_debug(f"{index + 1} ", end="")

    # Keep all dimensions in the original units returned by the API.
    # GetMassProperties returns an array of geometric properties.
    #
    # The array element order is documented in the SolidWorks API:
    # 0..2 = XYZ centroid, 3 = volume, 4 = surface area,
    # 5 = mass, 6..11 = moments of inertia.
    props = body.GetMassProperties(1.0)

    if props is None or len(props) < 12:
        raise RuntimeError(f"\nCould not retrieve properties for body index {index}")

    # The bounding box in document coordinates.
    # These are approximate bounds, not the exact surface geometry.
    # xmin, ymin, zmin, xmax, ymax, zmax.
    bbox = body.GetBodyBox()

    if bbox is None or len(bbox) != 6:
        raise RuntimeError(f"\nCould not retrieve the bounding box for body index {index}")

    # Retrieve all body faces so that, in addition to the total face count,
    # small faces can be counted as a potential indicator of embossed symbols.
    faces = sw_as_list(body.GetFaces())
    small_face_count = sum(1 for face in faces if float(face.GetArea) < SMALL_FACE_AREA_M2)

    # Keep the body name for debugging only.
    name = str(body.Name)

    return {
        "index": index,
        "name_debug": name,
        "centroid_m": props[0:3],
        "volume_m3": props[3],
        "area_m2": props[4],
        "inertia_com": props[6:12],
        "bbox_m": bbox,
        "face_count": len(faces),
        "small_face_count": small_face_count,

        "raw_faces": faces,
        "raw_body": body,
    }


def extract_document_geometry(document):
    """Extracts the geometry of all Solid Bodies in the open part."""

    if int(document.GetType) != SW_DOC_PART:
        raise RuntimeError("Expected a SolidWorks Part document (*.SLDPRT).")

    # NOTE: the second argument specifies whether only visible bodies are requested
    bodies = sw_as_list(document.GetBodies2(SW_SOLID_BODY, False))
    # Also retrieve all sheet bodies.
    sheet_bodies = com_array_to_list(document.GetBodies2(1, False))

    if not bodies:
        raise RuntimeError("No Solid Bodies were found in the document.")

    print_if_debug(f"\nAnalyzing document bodies (total bodies={len(bodies)}):")
    result = [extract_body_geometry(body, index) for index, body in enumerate(bodies)]
    print_if_debug("")
    return result, sheet_bodies


def print_document_geometry(records):
    print_if_debug(f"\nNumber of Solid Bodies: {len(records)}")
    print_if_debug("-" * 100)

    # Sort only the displayed report to make the geometry easier to inspect.
    # The original indices in the data remain unchanged.
    for rec in sorted(records, key=lambda item: item["centroid_m"][0]):
        # Convert COM coordinates and bounds from meters to millimeters for readability only.
        x, y, z = (v * 1000 for v in rec["centroid_m"])

        # Volume: m³ -> mm³; surface area: m² -> mm².
        volume = rec["volume_m3"] * 1e9
        area = rec["area_m2"] * 1e6

        print_if_debug(
            f"#{rec['index']:02d} "
            f"{rec['name_debug'][:22]:22} | "
            f"COM=({x:8.2f}, {y:8.2f}, {z:8.2f}) mm | "
            f"V={volume:12.1f} mm³ | "
            f"A={area:10.1f} mm² | "
            f"faces={rec['face_count']} | "
            f"s_faces={rec['small_face_count']}"
        )

    # Use the maximum volume for diagnostics only.
    largest = max(records, key=lambda item: item["volume_m3"])

    print_if_debug("-" * 100)
    print_if_debug(f"Largest body by volume: {largest['name_debug']}; V={largest['volume_m3'] * 1e9:.1f} mm³")
...  # ========================= /SolidWorks API =========================


def score_model(document):
    # diagnose_glyph_features(document); return  # In theory, this could be used to inspect the labels and their mirroring in more detail.

    # ============= Retrieve all solid bodies =============
    records, sheet_bodies = extract_document_geometry(document)
    print_document_geometry(records)

    # ============= Load the baseline =============
    baseline, unassigned_ids = load_baseline()
    print_baseline(baseline, unassigned_ids)

    # ============= Match the components =============
    global_result = match_bodies_global(baseline, records)
    print_global_match_report(baseline, global_result)

    # ========================== RUN THE CHECKS ==========================
    # --------- Feature tree ---------
    tree_report = diagnose_feature_tree_health(document)
    tree_score = score_feature_tree_health(tree_report)
    print_feature_tree_score(tree_score)

    # --------- Housing widening (by pairs) ---------
    widening_report = diagnose_widening_pairs(baseline, global_result["matches"])
    widening_score = score_widening_pairs(widening_report)
    print_widening_pairs_score(widening_score)

    # --------- Button-cluster placement ---------
    diagnose_button_clusters(baseline, records)
    button_layout_score = score_button_clusters(baseline, records, global_result["button_classification"])
    print_button_clusters_score(button_layout_score)

    # --------- Center buttons (|) ---------
    centre_score = score_centre_positions(baseline, global_result["matches"])

    # --------- Mirroring of labels (L      - .... R) ---------
    light_cluster_result = diagnose_light_cluster(baseline, records)
    light_score = score_light_cluster(light_cluster_result, baseline)
    print_light_cluster_score(light_score)

    # --------- Presence of button symbols ---------
    glyph_score = score_glyph_preservation(baseline, global_result["matches"])

    # --------- Component intersections ---------
    interference_result = diagnose_control_interference(baseline, records, sheet_bodies, global_result["matches"])
    interference_score = score_control_interference(interference_result, baseline)

    # Including every internal part adds many fasteners and internal components; the resulting intersection contribution is 0.874 x 0.10 = 0.087.
    # all_intersections_result = diagnose_all_solid_intersections(records)
    # interference_score = score_all_solid_intersections(all_intersections_result)

    # Intersections with other large bodies can also be checked if this turns out to matter.
    if DEBUG_FULL_OUTPUT:
        diagnose_unmatched_housing_sized_bodies(baseline, records, global_result["matches"])

    # ============= Calculate the final score =============
    total_result = calculate_total_score(
        tree_score,
        widening_score,
        button_layout_score,
        centre_score,
        light_score,
        glyph_score,
        interference_score,
    )

    return total_result


def inspect_candidate_or_active(document_path=None):
    # Start SolidWorks and open the document
    sw_app, document = get_sw_and_doc(document_path)

    try:
        # Attempt to score the model
        return score_model(document)
    except Exception as e:
        # Report the error at the point where it occurred
        raise e
    finally:
        # Close the current document
        close_sw(sw_app)


def main(argv):
    candidate = argv[1] if len(argv) > 1 else None

    result = inspect_candidate_or_active(candidate)

    envelope = {
        "candidate": candidate,
        "note": "",
        "score": result['score'],
        "max_score": MAX_SCORE,
        "passed": result['passed'],
        "subscores": result['components']
    }

    print(json.dumps(envelope, ensure_ascii=False, indent=2))

    return 0


if __name__ == "__main__":
    argv = sys.argv

    if "--debug" in argv or "-d" in argv:
        DEBUG_FULL_OUTPUT = True
        print_if_debug("DEBUG MODE ENABLED.")
        argv.remove("--debug") if "--debug" in argv else argv.remove("-d")

    sys.exit(main(argv))
