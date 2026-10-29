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
# NOTE: Распараллелить эту задачу через threading внутри одного процесса SolidWorks не получится.
#  Сам API SolidWorks спроектирован как строго однопоточный. (Даже используя pythoncom.CoInitialize())
#  Можно через multiprocessing, но создавая для каждого процесса свой экземпляр SLDWORKS.exe (через DispatchEx)


MAX_SCORE = 4  # keep equal to max_score in task.toml

# CHANGEABLE
DEBUG_FULL_OUTPUT = False
CLOSE_OUR_SW_IN_THE_END = True
# /CHANGEABLE

# Константы SolidWorks
SW_SOLID_BODY = 0  # Тип тела swSolidBody в API SolidWorks
SW_BODY_INTERSECT = 15901  # Операция пересечения
SW_DOC_PART = 1  # Тип документа - '.SLDPRT'

# Корень задания: каталог, внутри которого лежат harness, tests, environment и остальные файлы проекта.
PROJECT_DIR = Path(__file__).resolve().parents[1]

# Исходные измерения, собранные для эталонной модели.
BASELINE_PATH = PROJECT_DIR / "prompt" / "input.json"

# Общие константы для программы
SMALL_FACE_AREA_M2 = 15e-6  # Порог площади грани: 15 мм²
TARGET_WIDENING_MM = 15.0
HALF_TARGET_WIDENING_M = TARGET_WIDENING_MM / 1000 / 2


# ========================= Baseline =========================
def load_baseline():
    """
    Загружает исходные измерения и проверяет корректность разметки.
    Имена тел используются только внутри эталонного JSON;
    геометрическое сопоставление реализуем отдельно.
    """

    if not BASELINE_PATH.is_file():
        raise FileNotFoundError(f"Не найден файл эталона: {BASELINE_PATH}")

    with BASELINE_PATH.open("r", encoding="utf-8") as file:
        baseline = json.load(file)

    if baseline.get("schema") != "ps-annotation-baseline/1":
        raise ValueError("Неизвестная схема input.json.")

    bodies = baseline.get("bodies", [])
    roles = baseline.get("roles", {})

    body_ids = [body["id"] for body in bodies]

    if len(body_ids) != len(set(body_ids)):
        raise ValueError("В эталоне есть повторяющиеся ID тел.")

    known_ids = set(body_ids)
    assigned_ids = [body_id for group_ids in roles.values() for body_id in group_ids]

    if len(assigned_ids) != len(set(assigned_ids)):
        raise ValueError("Одно тело включено в несколько ролей.")

    unknown_ids = set(assigned_ids) - known_ids

    if unknown_ids:
        raise ValueError(f"Роли ссылаются на неизвестные ID: {sorted(unknown_ids)}")

    unassigned_ids = known_ids - set(assigned_ids)

    return baseline, unassigned_ids


def print_baseline(baseline, unassigned_ids):
    print_if_debug("\nЭталон:")
    print_if_debug(f"Файл: {BASELINE_PATH}")
    print_if_debug(f"Схема: {baseline['schema']}")
    print_if_debug(f"Тел в эталоне: {len(baseline['bodies'])}")
    print_if_debug(f"Плоскость симметрии X: {baseline['plane_x_m'] * 1000:.3f} мм")

    print_if_debug("\nГруппы эталона:")
    for role, body_ids in baseline["roles"].items():
        print_if_debug(f"  {role}: {len(body_ids)} тел")

    print_if_debug(f"\nТел без назначенной роли: {sorted(unassigned_ids)}")
...  # ==================== /Baseline =========================


# ========================= Global Matcher =========================
def match_bodies_global(baseline, records, missing_cost=2.5, min_margin=0.08, position_scale_mm=25.0):
    """
    Сопоставляет эталонные тела с текущей моделью.

    Для каждого эталонного тела находит несколько наиболее
    подходящих текущих тел с учётом вариантов преобразования.

    Условия:
    - один текущий Solid Body нельзя назначить двум эталонным телам;
    - D-pad и face_buttons ищутся только в распознанных кластерах;
    - положение сравнивается с допустимыми состояниями преобразования;
    - слишком дорогие или неоднозначные совпадения пропускаются;
    - дополнительные тела остаются в unmatched.

    Параметры missing_cost и min_margin пока являются начальными
    настройками. Их будем корректировать по результатам тестовых моделей.
    """

    if not records:
        raise ValueError("Текущая модель не содержит Solid Bodies.")

    baseline_by_id = {body["id"]: body for body in baseline["bodies"]}
    role_by_id = {body_id: role for role, ids in baseline["roles"].items() for body_id in ids}

    small_face_counts = baseline.get("small_face_counts", {})

    # Корпус определяем отдельно по V, чтобы не перепутать с мелкими деталями из-за изменений геометрических признаков.
    # И как ориентир используем крупнейшее тело. Но это не точное соответствие исходному housing из input.json.
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

    # Сначала определяем физические кластеры и назначение каждого кластера.
    clusters = find_button_clusters(records, baseline["plane_x_m"])
    classification = classify_button_clusters(baseline, clusters)
    cluster_mapping = classification["mapping"] if classification["confident"] else None

    # Если классификация уверенная, используем найденные соответствия кнопок как мягкий  приоритет при
    # глобальном сопоставлении. Приоритет не фиксирует назначение: геометрия по-прежнему может его переиграть.
    button_matches = (classification.get("button_matches", {}) if classification["confident"] else {})

    button_layout_penalty = 0.05  # Мини-штраф в пользу уже найденной кнопки выше

    cluster_indices = {side: {body["index"] for body in bodies} for side, bodies in clusters.items()}
    all_button_indices = cluster_indices["left"] | cluster_indices["right"]

    # Исключаем корпус
    references = [body for body in baseline["bodies"] if body["id"] != housing_id]
    candidates = [body for body in records if body["index"] not in used_indices]

    # Стоимость заведомо недопустимых назначений значительно выше стоимости отсутствующего тела.
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

            # Мы выше уже (скорее всего) определили какие тела могут быть кнопками.
            # Так что при поиске кнопок - пропускаем всё остальное и наоборот. Тем самым ограничиваем физические
            # группы, чтобы кнопки не могли случайно стать центральными элементами или наоборот.
            allowed = True

            if role in ("dpad", "face_buttons"):
                allowed = (cluster_mapping is not None and current_index in cluster_indices[cluster_mapping[role]])
            elif current_index in all_button_indices:
                allowed = False

            if not allowed or not hypotheses:
                row.append(forbidden_cost)
                continue

            # Сравниваем форму объектов, где бы они ни были
            shape_cost = body_shape_distance(reference, candidate)

            best = None
            # Проверяем все возможные гипотезы перестановок
            for hypothesis in hypotheses:
                # Евклидово расстояние между ожидаемым и текущим центром масс, в миллиметрах.
                position_error_mm = _centroid_distance_mm(candidate["centroid_m"], hypothesis["centroid_m"])

                # Нормировка: ошибка позиции 25 мм даёт +1 к стоимости.
                position_cost = position_error_mm / position_scale_mm

                small_face_cost = 0.0
                reference_small_faces = small_face_counts.get(reference_id)
                candidate_small_faces = candidate.get("small_face_count", 0)
                if reference_small_faces is not None:
                    abs_diff = abs(candidate_small_faces - reference_small_faces)
                    small_face_cost = 0.003 * min(abs_diff, 100)

                # Слегка предпочитаем соответствие, найденное при анализе внутренней структуры кластера.
                # Если геометрия убедительно указывает на другое тело, глобальный matcher всё равно сможет выбрать его.
                layout_penalty = 0.0
                if role in ("dpad", "face_buttons"):
                    preferred = button_matches.get(reference_id)

                    if preferred is not None and int(preferred["index"]) != current_index:
                        # Мы уже (скорее всего) нашли нужную кнопку и не даём ей штраф, а другим даём - 0.05
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

        # Проверяем неоднозначность: если два разных тела почти одинаково
        # подходят одной роли, не назначаем её без уверенного преимущества.
        ranked = sorted(candidate_scores.items(), key=lambda item: item[1]["match_cost"])
        # Разница между двумя лучшими кандидатами показывает, насколько уверенно лучший вариант выигрывает у второго.
        margin = ranked[1][1]["match_cost"] - ranked[0][1]["match_cost"] if len(ranked) > 1 else float("inf")
        second_name = records[ranked[1][0]]['name_debug'] if len(ranked) > 1 else "N/A"
        ambiguous = len(ranked) > 1 and margin < min_margin
        best_cost = ranked[0][1]["match_cost"] if ranked else float("inf")

        if ambiguous:
            row = [forbidden_cost] * len(candidates)

        # Фиктивные столбцы позволяют пропускать отсутствующие тела.
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
    """Выводит результаты глобального сопоставления по ролям."""

    matches = result["matches"]
    details = result["details"]
    unmatched = result["unmatched"]

    role_by_id = {body_id: role for role, ids in baseline["roles"].items() for body_id in ids}
    reference_by_id = {body["id"]: body for body in baseline["bodies"]}

    print_if_debug("\n" + "=" * 90)
    print_if_debug("ГЛОБАЛЬНОЕ СОПОСТАВЛЕНИЕ")
    print_if_debug("=" * 90)

    for role in baseline["roles"]:
        print_if_debug(f"\n{role}:")

        if role == "housing":
            reference_id = baseline["roles"]["housing"][0]
            anchor = matches.get(reference_id)

            if anchor is None:
                print_if_debug("  Крупнейший компонент корпуса не найден.")
            else:
                print_if_debug(
                    f"  housing_anchor -> #{anchor['index']:02d} "
                    f"{body_name(anchor)[:22]:22} | "
                    f"V={body_volume_mm3(anchor):.1f} мм³"
                )

            continue

        for reference_id in baseline["roles"][role]:
            body = matches.get(reference_id)
            info = details.get(reference_id, {})

            if body is None:
                print_if_debug(f"  {reference_id}: НЕТ СООТВЕТСТВИЯ | причина={info.get('reason', 'unknown')}")
                continue

            margin = info.get("margin")
            margin_text = f"{margin:.3f}" if margin is not None and math.isfinite(margin) else "n/a"

            print_if_debug(
                f"  {reference_id} -> #{body['index']:02d} {body_name(body)[:22]:22} | "
                f"state={info.get('state', 'n/a')} | "
                f"cost={info.get('match_cost', 0.0):.3f} | "
                f"shape={info.get('shape_cost', 0.0):.3f} | "
                f"pos={info.get('position_error_mm', 0.0):.2f} мм | "
                f"s_face={"r:{} h:{}".format(*info.get('s_faces'))} | "
                f"margin={margin_text} ({info.get('second_name')})"
            )

    print_if_debug("\nunclassified:")
    for reference_id in sorted(result["unclassified_ids"]):
        body = matches.get(reference_id)
        info = details.get(reference_id, {})

        if body is None:
            print_if_debug(f"  {reference_id}: НЕТ СООТВЕТСТВИЯ | причина={info.get('reason', 'unknown')}")
        else:
            print_if_debug(
                f"  {reference_id} -> #{body['index']:02d} {body_name(body)[:22]:22} | cost={info.get('match_cost', 0.0):.3f}")

    classification = result["button_classification"]
    print_if_debug("\nКластеры кнопок:")
    if classification["confident"]:
        for role, side in classification["mapping"].items():
            print_if_debug(f"  {role} -> {side}")
    else:
        print_if_debug("  Уверенная классификация не получена.")

    # Сравниваем объём и площадь кнопок с исходным baseline.
    # Это диагностика: она поможет понять, меняются ли тела кнопок,
    # когда из модели удаляют нанесённые на них символы.
    baseline_by_id = {body["id"]: body for body in baseline["bodies"]}

    print_if_debug("\nГеометрия кнопок относительно baseline:")
    for role in ("dpad", "face_buttons"):
        print_if_debug(f"  {role}:")
        for ref_id in baseline["roles"][role]:
            reference = baseline_by_id[ref_id]
            candidate = matches.get(ref_id)

            if candidate is None:
                print_if_debug(f"    {ref_id}: нет соответствия")
                continue

            vol_delta = 100.0 * (candidate["volume_m3"] / reference["volume_m3"] - 1.0)
            area_delta = 100.0 * (candidate["area_m2"] / reference["area_m2"] - 1.0)
            print_if_debug(f"    {ref_id} -> #{candidate['index']:02d}: "
                           f"V={vol_delta:+.2f}%, "
                           f"A={area_delta:+.2f}%, "
                           f"faces={candidate.get('face_count', '?')}, "
                           f"small_faces={candidate.get('small_face_count', '?')}")

    print_if_debug(f"\nЭталонных тел без соответствия: {len(result['missing_reference_ids'])}")
    print_if_debug(f"Дополнительных текущих тел: {len(unmatched)}")

    for body in unmatched:
        print_if_debug(f"  #{body['index']:02d} {body_name(body)[:22]:22} | "
                       f"V={body_volume_mm3(body):.1f} мм³ | "
                       f"COM=({" | ".join(f'{value * 1000:.3f}' for value in body["centroid_m"])}) (мм)")
...  # ==================== /Global Matcher =========================


# ========================= Scoring =========================
def diagnose_feature_tree_health(document):
    """Выводит общее состояние дерева построения и ошибки отдельных операций."""
    print_if_debug("\n" + "=" * 90)
    print_if_debug("ДИАГНОСТИКА ДЕРЕВА ПОСТРОЕНИЯ")
    print_if_debug("=" * 90)

    whats_wrong = {"success": False, "problem_count": None, "items": []}
    try:
        whats_wrong = get_whats_wrong(document)
        print_if_debug(f"What's Wrong API успешно отработал: {whats_wrong['success']}")
        print_if_debug(f"Проблем в What's Wrong: {whats_wrong['problem_count']}")

        if whats_wrong["success"]:
            errors = [item for item in whats_wrong["items"] if not item["is_warning"]]
            warnings = [item for item in whats_wrong["items"] if item["is_warning"]]

            print_if_debug(f"Ошибок: {len(errors)}; предупреждений: {len(warnings)}")

            for item in whats_wrong["items"]:
                kind = "WARNING" if item["is_warning"] else "ERROR"
                print_if_debug(f"  [{kind}] {item['name']} [{item['type']}] code={item['error_code']}")

    except Exception as exc:
        print_if_debug(f"Не удалось получить What's Wrong: {type(exc).__name__}: {exc}")

    additional_feature_errors = False  # NOTE: Даёт на 1 предупреждение больше, но тратит много времени на весь обход
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

        print_if_debug(f"\nУзлов дерева проверено: {scanned_count}")
        print_if_debug(f"Не удалось проверить узлов: {scan_failures}")
        print_if_debug(f"Ненулевых кодов GetErrorCode2: {len(feature_errors)}")

        for item in feature_errors:
            kind = "WARNING" if item["is_warning"] else "ERROR"
            indent = "  " * item["depth"]
            print_if_debug(f"{indent}[{kind}] {item['name']} [{item['type']}] code={item['error_code']}")

    return whats_wrong


def score_feature_tree_health(whats_wrong):
    """Оценивает состояние дерева по списку проблем What's Wrong."""

    if not whats_wrong["success"]:
        return {"score": None, "status": "unavailable", "errors": 0, "warnings": 0}

    items = whats_wrong["items"]
    errors = sum(not item["is_warning"] for item in items)
    warnings = sum(item["is_warning"] for item in items)

    score = math.exp(-0.18 * errors - 0.08 * warnings)

    return {"score": score, "status": "ok", "errors": errors, "warnings": warnings}


def print_feature_tree_score(result):
    """Выводит численный показатель состояния дерева."""

    print_if_debug("\nОценка здоровья дерева построения:")

    if result["status"] != "ok":
        print_if_debug("Оценку рассчитать не удалось: What's Wrong API недоступен.")
        return

    print_if_debug(f"Ошибки: {result['errors']}; предупреждения: {result['warnings']}")
    print_if_debug(f"Балл за дерево: {result['score']:.3f} / 1.000")


def diagnose_widening_pairs(baseline, matches):
    """Показывает реальные координаты эталонных и сопоставленных тел пар."""

    print_if_debug("\n" + "=" * 90)
    print_if_debug("ДИАГНОСТИКА ПАР РАСШИРЕИНЯ")
    print_if_debug("=" * 90)

    print_if_debug("\nКоординаты пар расширения:")

    reference_by_id = {body["id"]: body for body in baseline["bodies"]}

    pair_definitions = get_widening_pairs(baseline)
    measurements = []

    for pair_name, (left_id, right_id) in pair_definitions.items():
        print_if_debug(f"\n{pair_name}:")

        # Если одной из деталей нет, не придумываем её координаты.
        # Отсутствие тел будет отдельно учтено в проверке полноты.
        if left_id not in matches or right_id not in matches:
            measurements.append({"pair": pair_name, "missing": True})
            print_if_debug(f"N/A:")
            continue

        reference_left = reference_by_id[left_id]
        reference_right = reference_by_id[right_id]

        current_left = matches[left_id]
        current_right = matches[right_id]

        # Расстояния измеряем вдоль оси X. (Сама плоскость симметрии при этом не обязана быть X=0.)
        x_left_r = reference_left["centroid_m"][0]
        x_right_r = reference_right["centroid_m"][0]
        baseline_separation_m = abs(x_left_r - x_right_r)

        x_left_c = current_left["centroid_m"][0]
        x_right_c = current_right["centroid_m"][0]
        current_separation_m = abs(x_left_c - x_right_c)

        growth_mm = (current_separation_m - baseline_separation_m) * 1000.0

        midpoint_mm = (current_left["centroid_m"][0] + current_right["centroid_m"][0]) * 1000.0 / 2

        print_if_debug(f"  {left_id} -> #{current_left['index']:02d} {body_name(current_left):22} | "
                       f"X: {x_left_r * 1000:.2f} -> {x_left_c * 1000:.2f} мм")
        print_if_debug(f"  {right_id} -> #{current_right['index']:02d} {body_name(current_right):22} | "
                       f"X: {x_right_r * 1000:.2f} -> {x_right_c * 1000:.2f} мм")

        print_if_debug(f"  Расстояние: {baseline_separation_m * 1000.0:.2f} -> "
                       f"{current_separation_m * 1000.0:.2f} мм | "
                       f"прирост={growth_mm:+.2f} мм | "
                       f"середина X={midpoint_mm:.2f} мм")

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
    Вычисляет непрерывный показатель расширения в диапазоне 0..1.

    Берём  результаты геометрического сопоставления и координаты центров масс.
    Полный балл соответствует увеличению зеркальных расстояний на 15 мм.
    """

    valid_growths = [item["growth_mm"] for item in measurements if not item["missing"]]

    # Медиана остаётся диагностическим показателем, но больше не используется непосредственно для итогового балла.
    median_growth = (statistics.median(valid_growths) if valid_growths else None)

    # Сначала рассчитываем отдельный балл каждой зеркальной пары. Если тела отсутствуют, пара получает 0 баллов.
    for item in measurements:
        if item["missing"]:
            item["score"] = 0.0
            continue

        deviation = abs(item["growth_mm"] - TARGET_WIDENING_MM)

        # Линейная оценка отклонения от целевого прироста +15 мм.
        # 0 мм и 30 мм дают 0 баллов; ровно 15 мм - 1 балл.
        item["score"] = max(0.0, min(1.0, 1.0 - deviation / TARGET_WIDENING_MM))

    # Итоговый балл - среднее по всем ожидаемым парам. Одна неправильно изменённая пара снижает общий результат.
    pair_scores = [item["score"] for item in measurements]

    score = (statistics.fmean(pair_scores) if pair_scores else 0.0)

    return {
        "score": score,
        "target_growth_mm": TARGET_WIDENING_MM,
        "median_growth_mm": median_growth,
        "pairs": measurements,
    }


def print_widening_pairs_score(widening):
    print_if_debug("\nКритерий расширения корпуса:")
    for item in widening["pairs"]:
        if item["missing"]:
            print_if_debug(f"  {item['pair']}: нет соответствия одной из деталей")
            continue

        print_if_debug(f"  {item['pair']}: (score={item['score']:.3f})")

    if widening["median_growth_mm"] is None:
        print_if_debug("Недостаточно пар для вычисления медианного прироста.")
    else:
        print_if_debug(f"Медианный прирост: {widening['median_growth_mm']:+.2f} мм")

    print_if_debug(f"Целевой прирост: +{widening['target_growth_mm']:.1f} мм")
    print_if_debug(f"Балл за расширение: {widening['score']:.3f} / 1.000")


def diagnose_button_clusters(baseline, records):
    """
    Диагностически выделяет два кластера по расположению тел,
    а затем сравнивает их взаимную геометрию с эталонными группами.

    Имена тел выводятся только для проверки человеком.
    В критериях алгоритма имена не используются.
    """

    plane_x_mm = baseline["plane_x_m"] * 1000.0
    clusters = find_button_clusters(records, baseline["plane_x_m"])

    print_if_debug("\n" + "=" * 80)
    print_if_debug("ДИАГНОСТИКА КЛАСТЕРОВ КНОПОК")
    print_if_debug("=" * 80)

    for side, bodies in clusters.items():
        print_if_debug(f"\nКластер {side}: {len(bodies)} тел")

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
                f"X={x * 1000:8.2f} мм | "
                f"Z={z * 1000:8.2f} мм | "
                f"dX={(x - center_x) * 1000:+7.2f} мм | "
                f"dZ={(z - center_z) * 1000:+7.2f} мм | "
                f"V={body['volume_m3'] * 1e9:9.1f} мм³ | "
                f"small_faces={body['small_face_count']}"
            )

    reference_by_id = {body["id"]: body for body in baseline["bodies"]}

    print_if_debug("\nОшибка относительной геометрии кластеров (мм):")

    for role in ("dpad", "face_buttons"):
        reference_group = [reference_by_id[body_id] for body_id in baseline["roles"][role]]

        for side, current_group in clusters.items():
            error = cluster_layout_rms(reference_group, current_group)

            if error is None:
                print_if_debug(f"  {role} -> {side}: нельзя сравнить, найдено {len(current_group)} тел")
            else:
                print_if_debug(f"  {role} -> {side}: RMS={error:.3f} мм")

    classification = classify_button_clusters(baseline, clusters)

    print_if_debug("\nАвтоматическая классификация кластеров:")

    if not classification["confident"]:
        print_if_debug("  Уверенное назначение не найдено.")
        print_if_debug(f"  Причина: {classification.get('reason', 'недостаточная уверенность')}")

        if "best_total_rms_mm" in classification:
            print_if_debug(
                f"  Лучший вариант: "
                f"{classification['best_total_rms_mm']:.3f} мм"
            )
            print_if_debug(
                f"  Альтернативный вариант: "
                f"{classification['alternative_total_rms_mm']:.3f} мм"
            )
    else:
        for role, side in classification["mapping"].items():
            print_if_debug(f"  {role} -> {side}")

        print_if_debug(
            f"  RMS лучшего назначения: "
            f"{classification['best_total_rms_mm']:.3f} мм"
        )
        print_if_debug(
            f"  Отрыв от альтернативы: "
            f"{classification['margin_mm']:.3f} мм"
        )


def score_button_clusters(baseline, records, cluster_classification=None):
    """
    Оценивает перестановку и размещение D-pad и лицевых кнопок.

    Ищет два кластера по четыре тела, сопоставляет их с эталонными
    группами по относительному расположению и оценивает позицию каждой
    группы относительно ожидаемого места после перестановки и расширения.

    Возвращает общий балл 0..1 и подробности по каждой группе.
    """

    plane_x = baseline["plane_x_m"]
    plane_x_mm = plane_x * 1000.0

    # Отбираем тела кнопок по геометрии, не используя имена.
    candidates = find_bodies_by_feature(
        records,
        y_range_mm=(-12.0, 12.0),
        z_range_mm=(-35.0, 35.0),
        volume_range_mm3=(500.0, 10000.0),
        predicate=lambda body: abs(body_centroid_mm(body)[0] - plane_x_mm) >= 30.0,
    )

    # Перебор всех разбиений становится неоправданно дорогим при
    # слишком большом количестве кандидатов. В текущих моделях
    # ожидаем 8 тел; допускаем до 12 на случай дополнительных тел.
    if not 8 <= len(candidates) <= 12:
        return {
            "score": 0.0,
            "status": "candidate_count_invalid",
            "candidate_count": len(candidates),
            "candidates": candidates,
            "reason": f"Ожидалось 8–12 кандидатов, найдено {len(candidates)}.",
        }

    reference_groups = {role: get_reference_bodies(baseline, role) for role in ("dpad", "face_buttons")}

    def center_of_group(bodies):
        """Среднее положение центров масс группы в метрах."""
        return tuple(sum(body["centroid_m"][axis] for body in bodies) / len(bodies) for axis in range(3))

    # Вычисляем целевые центры: отражаем группу относительно плоскости,
    # а затем раздвигаем её ещё на 7,5 мм наружу.
    targets = {}

    for role, bodies in reference_groups.items():
        original_center = center_of_group(bodies)
        direction = 1.0 if original_center[0] < plane_x else -1.0
        target_x = 2.0 * plane_x - original_center[0] + direction * HALF_TARGET_WIDENING_M

        targets[role] = (target_x, original_center[1], original_center[2])

    # Ищем две непересекающиеся четвёрки тел, лучше всего соответствующие
    # эталонным D-pad и face_buttons. Если кандидатов больше восьми,
    # оставшиеся тела не включаются в кнопочные кластеры.
    best = None

    button_matches = (
        cluster_classification.get("button_matches", {})
        if cluster_classification and cluster_classification.get("confident")
        else {}
    )

    dpad_ids = baseline["roles"]["dpad"]
    face_ids = baseline["roles"]["face_buttons"]

    # Если классификатор уже сопоставил все восемь кнопок,
    # используем его результат вместо перебора комбинаций четвёрок.
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

    # Если классификация неуверенная, сохраняем прежний резервный поиск.
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

            # Форма определяет роль кластера; положение используется
            # как дополнительный признак, чтобы различать варианты размещения.
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
            "reason": "Не удалось найти две подходящие группы по четыре тела.",
        }

    def shape_score(rms_mm):
        """Полный балл при RMS до 0,5 мм, ноль при RMS от 3 мм."""
        if rms_mm <= 0.5:
            return 1.0
        return max(0.0, 1.0 - (rms_mm - 0.5) / 2.5)

    def placement_score(error_mm):
        """Допуск 1 мм, далее линейное снижение до нуля при 15 мм."""
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
    """Выводит баллы и найденные группы кнопок."""

    print_if_debug("\n" + "=" * 80)
    print_if_debug("ОЦЕНКА РАСКЛАДКИ КНОПОК")
    print_if_debug("=" * 80)

    if result["status"] != "ok":
        print_if_debug(f"Оценка не рассчитана: {result['reason']}")
        print_if_debug(f"Кандидатов: {result['candidate_count']}")
        return

    for role, details in result["details"].items():
        print_if_debug(f"\n{role}:")

        for body in details["bodies"]:
            print_if_debug(f"  #{body['index']:02d} {body_name(body)}")

        print_if_debug(f"  RMS внутренней раскладки: {details['layout_rms_mm']:.3f} мм")
        print_if_debug(f"  Ошибка положения: {details['position_error_mm']:.3f} мм")
        print_if_debug(f"  Оценка формы: {details['shape_score']:.3f}")
        print_if_debug(f"  Оценка положения: {details['placement_score']:.3f}")
        print_if_debug(f"  Балл группы: {details['score']:.3f}")

        target = details["target_center_mm"]
        current = details["current_center_mm"]
        print_if_debug(f"  Целевой центр: ({target[0]:.2f}, {target[1]:.2f}, {target[2]:.2f}) мм")
        print_if_debug(f"  Текущий центр: ({current[0]:.2f}, {current[1]:.2f}, {current[2]:.2f}) мм")

    print_if_debug(f"\nОбщий балл за раскладку кнопок: {result['score']:.3f} / 1.000")
    print_if_debug(f"Неиспользованных кандидатов: {len(result['unused_candidates'])}")


def score_centre_positions(baseline, matches, full_error_mm=1.0, zero_error_mm=15.0):
    """Оценивает зеркальное положение боковых центральных элементов (START/SELECT) и положение центральной кнопки PS."""
    reference_by_id = {body["id"]: body for body in baseline["bodies"]}
    plane_x = baseline["plane_x_m"]
    centre_ids = baseline["roles"].get("centre", [])

    if not centre_ids:
        return {"score": 0.0, "status": "no_reference", "elements": []}

    scores = []
    elements = []

    print_if_debug("\nОценка центральных элементов:")

    for body_id in centre_ids:
        reference = reference_by_id.get(body_id)
        current = matches.get(body_id)

        if reference is None or current is None:
            scores.append(0.0)
            elements.append({"id": body_id, "score": 0.0, "status": "missing"})
            print_if_debug(f"  {body_id}: соответствующее тело не найдено | score=0.000")
            continue

        x, y, z = reference["centroid_m"]

        # Боковые элементы должны перейти на противоположную сторону.
        # Центральная кнопка должна остаться на оси симметрии (plane_x).
        is_central_button = abs(x - plane_x) <= 0.010
        target_x = plane_x if is_central_button else 2.0 * plane_x - x

        # Расстояние от текущего центра масс до зеркальной цели.
        cx, cy, cz = current["centroid_m"]

        # Для этого критерия проверяем X, а не требуем точного совпадения
        # по Y и Z, которое отдельно не установлено условиями задания.
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

        print_if_debug(f"  {body_id}: target={state} | отклонение={error_mm:.2f} мм | score={body_score:.3f}")

    score = statistics.fmean(scores) if scores else 0.0
    status = "partial" if any(item["status"] == "missing" for item in elements) else "ok"

    print_if_debug(f"Итоговая оценка центра: {score:.3f} / 1.000")

    return {"score": score, "status": status, "elements": elements}


def diagnose_light_cluster(baseline, records, radius_mm=25.0, bodies_amount=1):
    """Ищет кластер индикаторов в трёх крупнейших телах около ожидаемой
    или зеркальной позиции по эталонному набору площадей граней."""

    reference = baseline["light_cluster"]
    target_areas = reference["areas_m2"]
    target_centroid = reference["centroid_m"]
    plane_x = baseline["plane_x_m"]
    radius_m = radius_mm / 1000.0

    # Кластер находится на корпусе. Вместо сканирования всех деталей
    # ранжируем твёрдые тела по объёму и проверяем только N крупнейших.
    candidate_bodies = sorted(records, key=body_volume_mm3, reverse=True)[:bodies_amount]

    # Проверяем исходное место кластера и его зеркальную позицию.
    # Второй регион нужен, чтобы диагностировать неправильное зеркальное отражение.
    search_regions = [
        ("expected", target_centroid),
        ("mirrored", (2.0 * plane_x - target_centroid[0], target_centroid[1], target_centroid[2])),
    ]

    best = None

    for body_rank, candidate_info in enumerate(candidate_bodies, start=1):
        body_index = candidate_info['index']

        # Получаем геометрию граней каждого выбранного тела только один раз.
        faces = []

        for face_index, face in enumerate(candidate_info['raw_faces']):
            area = face.GetArea
            box = face.GetBox

            if area <= 0 or box is None or len(box) < 6:
                continue

            # box = [float(value) for value in box]
            center = ((box[0] + box[3]) / 2, (box[1] + box[4]) / 2, (box[2] + box[5]) / 2)
            faces.append({"body_index": body_index, "face_index": face_index, "area": area, "center": center})

        # На каждое тело приходится только два сопоставления, а не одно
        # сопоставление на каждую возможную грань-центр.
        for region_name, region_center in search_regions:
            # Для каждого возможного центра ищем локальный участок и сопоставляем
            # его площади с эталонными площадями один к одному.
            local = [face for face in faces
                     if all(abs(face["center"][axis] - region_center[axis]) <= radius_m
                            for axis in range(3))]

            if len(local) < len(target_areas):
                # print_if_debug(f"\nlight_cluster: недостаточно граней в области поиска ({len(local)} из {len(target_areas)}).")
                continue

            # Сопоставляем эталонные и фактические площади граней.
            # Логарифмическая ошибка сравнивает относительные изменения площади.
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

            # Сначала выбираем лучшее совпадение площадей.
            # При равенстве предпочтение отдаём центру, ближайшему к целевой позиции.
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
        print_if_debug("\nlight_cluster: кластер не найден в трёх крупнейших телах рядом с целевыми позициями.")
        return {"found": False, "reason": "no_matching_cluster_in_top_bodies"}

    cx, cy, cz = (value * 1000 for value in best["centroid_m"])
    tx, ty, tz = (value * 1000 for value in target_centroid)
    mean_relative_error = math.expm1(best["mean_log_error"]) * 100.0

    print_if_debug("\n" + "=" * 80)
    print_if_debug("ДИАГНОСТИКА LIGHT CLUSTER")
    print_if_debug("=" * 80)
    print_if_debug(f"Проверено крупнейших тел: {len(candidate_bodies)}")
    print_if_debug(f"Лучшее тело: #{best['candidate_body_index']} (по объёму ранг {best['candidate_body_rank']})")
    print_if_debug(f"Область поиска: {best['candidate_region']}")
    print_if_debug(f"Локальных граней-кандидатов: {best['local_face_count']}")
    print_if_debug(f"Сопоставлено граней: {best['matched_count']} / {len(target_areas)}")
    print_if_debug(f"Среднее мультипликативное расхождение площадей: {mean_relative_error:.2f}%")
    print_if_debug(f"P90 логарифмической ошибки: {best['p90_log_error']:.4f}")
    print_if_debug(f"Эталонный центр: ({tx:.2f}, {ty:.2f}, {tz:.2f}) мм")
    print_if_debug(f"Найденный центр: ({cx:.2f}, {cy:.2f}, {cz:.2f}) мм")
    print_if_debug(f"Сторона относительно X={plane_x * 1000:.2f} мм: {best['side']:+d}")
    print_if_debug(f"Ожидаемая сторона: {reference['side']:+d}")

    best.pop("candidate_key", None)
    return best


def score_light_cluster(result, baseline, full_error=0.03, zero_error=0.15,
                        full_position_error_mm=5.0, zero_position_error_mm=25.0):
    """Оценивает сохранение положения и стороны кластера индикаторов."""

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

    # Проверяем положение найденного кластера относительно его исходного центра.
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

    # Ошибка площадей остаётся диагностикой качества найденного совпадения.
    # Она не обнуляет балл, если положение и сторона кластера правильные.
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
    print_if_debug(f"\nОценка маркировки корпуса: {light_score['score']:.3f} / 1.000")
    print_if_debug(f"Статус: {light_score['status']}")


def score_glyph_preservation(baseline, matches):
    """
    Оценивает сохранность геометрической детализации четырёх лицевых кнопок.

    По проверенным моделям кнопки без символов имеют 4 грани, а кнопки с символами - не менее 8 граней.
    """
    scores = []

    print_if_debug("\nСохранность символов на лицевых кнопках:")

    for ref_id in baseline["roles"]["face_buttons"]:
        candidate = matches.get(ref_id)

        if candidate is None:
            print_if_debug(f"  {ref_id}: тело не найдено | score=0.000")
            scores.append(0.0)
            continue

        face_count = candidate.get("face_count")

        if face_count is None:
            print_if_debug(f"  {ref_id}: нет данных о гранях | score=0.000")
            scores.append(0.0)
            continue

        # 4 грани = детализация символа отсутствует.
        # 8 и более граней = полная оценка по этому геометрическому признаку.
        # Между 4 и 8 гранями оценка растёт линейно.
        body_score = max(0.0, min(1.0, (face_count - 4) / 4.0))
        scores.append(body_score)

        print_if_debug(
            f"  {ref_id} -> #{candidate['index']:02d} {candidate['name_debug']:22}: "
            f"faces={face_count} | score={body_score:.3f}"
        )

    score = sum(scores) / len(scores) if scores else 0.0
    print_if_debug(f"Итоговая сохранность символов: {score:.3f}")

    return score


def diagnose_control_interference(baseline, records, sheet_bodies, matches):
    """Проверяет попарные пересечения 18 основных тел из эталонных ролей."""
    bodies = [r["raw_body"] for r in records]

    print_if_debug("\nПроверка тел корпуса:")
    print_if_debug(f"  Твёрдых тел: {len(bodies)}")
    print_if_debug(f"  Поверхностных тел: {len(sheet_bodies)}")

    solid_summary = []
    for index, body in enumerate(bodies):
        props = [float(v) for v in body.GetMassProperties(1.0)]
        box = [float(v) for v in body.GetBodyBox()]
        dims_mm = [(box[i + 3] - box[i]) * 1000 for i in range(3)]
        solid_summary.append((props[3], index, props[4], dims_mm, body.GetFaceCount()))

    for volume_m3, index, area_m2, dims_mm, face_count in sorted(solid_summary, reverse=True)[:8]:
        print_if_debug(f"  solid #{index:02d}: V={volume_m3 * 1e9:.1f} мм³, A={area_m2 * 1e6:.1f} мм², "
                       f"bbox={dims_mm[0]:.1f}*{dims_mm[1]:.1f}*{dims_mm[2]:.1f} мм, faces={face_count}")

    for index, body in enumerate(sheet_bodies):
        props = [float(v) for v in body.GetMassProperties(1.0)]
        box = [float(v) for v in body.GetBodyBox()]
        dims_mm = [(box[i + 3] - box[i]) * 1000 for i in range(3)]
        print_if_debug(f"  sheet #{index:02d}: A={props[3] * 1e6:.1f} мм², "
                       f"bbox={dims_mm[0]:.1f}*{dims_mm[1]:.1f}*{dims_mm[2]:.1f} мм, "
                       f"faces={body.GetFaceCount()}")

    selected = {}
    ref_by_index = {}
    missing = []

    # Собираем корпус и все тела из ролей задания; unclassified сюда не входят.
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

        # Отсекаем только численно ничтожные объёмы - это не геометрический допуск в мм.
        if volume_m3 > 1e-15:
            overlaps.append((ref_by_index[index_a], ref_by_index[index_b], volume_m3))

    tested_pairs = len(selected) * (len(selected) - 1) // 2
    total_volume_m3 = sum(item[2] for item in overlaps)

    print_if_debug("\nПересечения основных тел:")
    print_if_debug(f"  Тело найдено: {len(selected)} из {len(reference_ids)}")
    print_if_debug(f"  Рассмотрено пар: {tested_pairs} (эталон: {baseline['interference']['tested_pairs']})")
    print_if_debug(f"  Пар с пересечением: {len(overlaps)}")
    print_if_debug(f"  Сумма попарных объёмов: {total_volume_m3 * 1e9:.6f} мм³")

    for ref_a, ref_b, volume_m3 in overlaps:
        print_if_debug(f"    {ref_a} / {ref_b}: {volume_m3 * 1e9:.6f} мм³")

    for index_a, index_b, message in errors:
        print_if_debug(f"    ОШИБКА #{index_a}/#{index_b}: {message}")

    if missing:
        print_if_debug(f"  Не удалось подобрать тела для: {', '.join(missing)}")

    return {
        "tested_pairs": tested_pairs,
        "overlaps": overlaps,
        "total_volume_m3": total_volume_m3,
        "errors": errors,
        "missing": missing,
    }


def score_control_interference(result, baseline, warning_volume_mm3=5.0, decay_volume_mm3=1000.0):
    """Оценивает пересечения основных тел; мелкие контакты вызывают предупреждение и небольшой штраф."""
    expected_pairs = int(baseline["interference"]["tested_pairs"])

    # Если проверены не все ожидаемые пары или булевы операции завершились ошибками,
    # нельзя достоверно утверждать, что модель не содержит пересечений.
    incomplete = (
            result.get("tested_pairs") != expected_pairs
            or bool(result.get("errors"))
            or bool(result.get("missing"))
    )

    if incomplete:
        print_if_debug("\nОценка пересечений недоступна: проверка выполнена не полностью.")
        print_if_debug(f"  Ожидалось пар: {expected_pairs}; проверено: {result.get('tested_pairs')}")
        print_if_debug(f"  Ошибок операций: {len(result.get('errors', []))}")
        print_if_debug(f"  Несопоставленных тел: {len(result.get('missing', []))}")
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

    # Пересечения свыше порога постепенно уменьшают оценку.
    # Небольшие пересечения штрафуются лишь на 2% от критерия за каждую пару.
    significant_volume = sum(item["volume_mm3"] - warning_volume_mm3 for item in significant_overlaps)

    score = math.exp(-significant_volume / decay_volume_mm3) * (0.98 ** len(minor_overlaps))

    warnings = []

    for item in minor_overlaps:
        message = (f"Малое пересечение {item['pair'][0]} / {item['pair'][1]}: {item['volume_mm3']:.3f} мм³")
        warnings.append(message)
        print_if_debug(f"  WARNING: {message}")

    for item in significant_overlaps:
        message = (f"Пересечение {item['pair'][0]} / {item['pair'][1]}: {item['volume_mm3']:.3f} мм³")
        warnings.append(message)
        print_if_debug(f"  WARNING: {message}")

    status = "ok" if not warnings else "warning"
    print_if_debug(f"Оценка пересечений: {score:.3f} / 1.000")

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
    """Проверяет пересечения всех твёрдых тел, включая дополнительные детали и крепёж."""
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
        # Булевую операцию выполняем только для тел с пересекающимися габаритными коробками.
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

    print_if_debug("\nПересечения всех твёрдых тел:")
    print_if_debug(f"  Твёрдых тел: {len(bodies)}")
    print_if_debug(f"  Всего пар: {total_pairs}")
    print_if_debug(f"  Пар, прошедших фильтр габаритов: {bbox_candidates}")
    print_if_debug(f"  Пар с пересечением: {len(overlaps)}")
    print_if_debug(f"  Ошибок вычисления: {len(errors)}")

    for name_a, name_b, volume_m3 in overlaps:
        print_if_debug(f"  WARNING: {name_a} / {name_b}: {volume_m3 * 1e9:.6f} мм³")

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
    """Слегка штрафует за малые пересечения и сильнее — за значительные."""
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
            warnings.append(f"Малое пересечение {name_a} / {name_b}: {volume_mm3:.3f} мм³")
        else:
            significant_volume_mm3 += volume_mm3 - warning_volume_mm3
            warnings.append(f"Пересечение {name_a} / {name_b}: {volume_mm3:.3f} мм³")

    score = (0.98 ** minor_count) * math.exp(-significant_volume_mm3 / decay_volume_mm3)

    for warning in warnings:
        print_if_debug(f"  WARNING: {warning}")

    print_if_debug(f"Оценка пересечений всех тел: {score:.3f} / 1.000")

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
    """Проверяет пересечения дополнительных тел размером с корпус."""
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

    print_if_debug("\nПересечения дополнительных крупногабаритных тел:")

    if not candidates:
        print_if_debug("  Нераспределённых тел размером с корпус не найдено.")
        return

    for index, body, volume_mm3, dx, dz in candidates:
        print_if_debug(f"  #{index:02d}: V={volume_mm3:.1f} мм³, X={dx:.1f} мм, Z={dz:.1f} мм")

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
                print_if_debug(f"    ОШИБКА с {ref_id}: {exc}")
                continue

            if volume > 1e-15:
                overlaps += 1
                print_if_debug(f"    Пересечение с {role}/{ref_id}: {volume * 1e9:.6f} мм³")

        if overlaps == 0:
            print_if_debug(f"    Пересечений нет среди проверенных пар ({tested}).")


def calculate_total_score(tree, widening, button_layout, centre, light, glyph, interference):
    """Собирает отдельные оценки в итоговый результат из 4 баллов."""

    # Наша согласованная весовая шкала, сумма весов равна 4 баллам.
    weights = {
        "Дерево построения (⍋)": 0.45,
        "Расширение (↔)": 1.20,
        "Кластеры кнопок (|)": 1.10,
        "Центральные элементы (|)": 0.35,
        "Маркировка корпуса (R)": 0.55,
        "Символы кнопок (✗○☐△)": 0.25,
        "Пересечения (∩)": 0.10,
    }

    components = [
        ("Дерево построения (⍋)", tree.get("score")),
        ("Расширение (↔)", widening.get("score")),
        ("Кластеры кнопок (|)", button_layout.get("score")),
        ("Центральные элементы (|)", centre.get("score")),
        ("Маркировка корпуса (R)", light.get("score")),
        ("Символы кнопок (✗○☐△)", glyph),
        ("Пересечения (∩)", interference.get("score")),
    ]

    maximum_points = sum(weights.values())
    earned_points = 0.0
    unavailable = []
    details = {}

    print_if_debug("\n" + "=" * 64)
    print_if_debug("ИТОГОВАЯ ОЦЕНКА")
    print_if_debug("=" * 64)

    for name, raw_score in components:
        weight = weights[name]

        if raw_score is None or not math.isfinite(float(raw_score)):
            unavailable.append({"name": name, "weight": weight})

            # Если не удалось проверить пересечения повреждённой модели,
            # начисляем за этот критерий 0 баллов, но сохраняем предупреждение.
            points = 0.0 if name == "Пересечения (∩)" else None
            details[name] = {"score": None, "weight": weight, "points": points}

            if name == "Пересечения (∩)":
                print_if_debug(f"  {name:<26} N/A -> 0.000 (консервативно, вес {weight:.2f})")
            else:
                print_if_debug(f"  {name:<26} N/A   (вес {weight:.2f})")

            continue

        normalized_score = max(0.0, min(1.0, float(raw_score)))
        points = normalized_score * weight
        earned_points += points

        details[name] = {"score": normalized_score, "weight": weight, "points": points}
        print_if_debug(f"  {name:<26} {normalized_score:.3f} * {weight:.2f} = {points:.3f}")

    if unavailable and all(item["name"] == "Пересечения (∩)" for item in unavailable):
        final_score = earned_points
        status = "conservative_incomplete"
        print_if_debug("\nWARNING: пересечения не удалось полностью проверить.")
        print_if_debug(f"ИТОГО (консервативно): {final_score:.3f} / {maximum_points:.2f}")
    elif unavailable:
        unavailable_weight = sum(item["weight"] for item in unavailable)
        print_if_debug(f"\nЧастичная сумма: {earned_points:.3f} / {maximum_points:.2f}")
        print_if_debug(f"Итог не финализирован: недоступны критерии на {unavailable_weight:.2f} балла.")
        final_score = None
        status = "incomplete"
    else:
        final_score = earned_points
        status = "ok"
        print_if_debug(f"\nИТОГО: {final_score:.3f} / {maximum_points:.2f}")

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
    Определяет зеркальные пары для оценки расширения.

    Берём пары стиков, триггеров и бамперов непосредственно из геометрической разметки baseline.
    """

    pairs = {}

    for role in ("sticks", "triggers", "bumpers"):
        body_ids = baseline["roles"].get(role, [])

        if len(body_ids) == 2:
            pairs[role] = tuple(body_ids)

    return pairs


def _relative_log_error(actual, expected):
    """
    Считает относительное различие двух положительных величин.
    Логарифм отношения позволяет сравнивать тела разных размеров.
    """
    actual = max(float(actual), 1e-30)
    expected = max(float(expected), 1e-30)
    return abs(math.log(actual / expected))


def _body_dimensions(body):
    """
    Получает три размера габаритного параллелепипеда.
    Сортировка исключает зависимость от перестановки осей.
    Координаты центра тела здесь намеренно не используются.
    """
    box = body["bbox_m"]
    return sorted([
        box[3] - box[0],
        box[4] - box[1],
        box[5] - box[2],
    ])


def body_shape_distance(reference, candidate):
    """
    Оценивает различие формы эталонного и проверяемого тела.
    Используем объём, площадь и размеры тела.
    """

    volume_error = _relative_log_error(candidate["volume_m3"], reference["volume_m3"])

    area_error = _relative_log_error(candidate["area_m2"], reference["area_m2"])

    ref_dims = _body_dimensions(reference)
    cand_dims = _body_dimensions(candidate)

    dimensions_error = sum(_relative_log_error(current, expected) for current, expected in zip(cand_dims, ref_dims)) / 3

    return volume_error + 0.7 * area_error + 0.7 * dimensions_error


def get_transform_hypotheses(reference, role, baseline):
    """
    Формирует гипотезы положения эталонных тел для глобального сопоставления.

    Гипотезы влияют на выбор соответствий и косвенно на итоговые
    геометрические оценки, но не являются самостоятельным критерием.
    """

    x, y, z = reference["centroid_m"]
    plane_x = baseline["plane_x_m"]

    def make_hypothesis(state, target_x, target_y=y, target_z=z):
        return {"state": state, "centroid_m": [target_x, target_y, target_z]}

    # Неизвестные тела рассматриваем отдельно, как остающиеся
    # на исходных местах. Корпус обрабатывается другим методом.
    if role == "housing":
        return []

    if role == "unclassified":
        return [make_hypothesis("original", x)]

    # D-pad и лицевые кнопки перемещаются как целые кластеры.
    if role in ("dpad", "face_buttons"):
        reference_by_id = {body["id"]: body for body in baseline["bodies"]}

        group_ids = baseline["roles"][role]
        group_center_x = sum(
            reference_by_id[body_id]["centroid_m"][0]
            for body_id in group_ids
        ) / len(group_ids)

        # Отступ этого тела от центра своего кластера.
        # Он сохраняется при переносе всей группы.
        local_offset_x = x - group_center_x

        source_direction = (-1.0 if group_center_x < plane_x else 1.0)

        # Перенос кластера на противоположную сторону.
        swapped_center_x = 2.0 * plane_x - group_center_x

        swapped_direction = (1.0 if swapped_center_x > plane_x else -1.0)

        return [
            make_hypothesis("original", x),
            make_hypothesis("widened", x + source_direction * HALF_TARGET_WIDENING_M),
            make_hypothesis("swapped", swapped_center_x + local_offset_x),
            make_hypothesis("swapped+widened", swapped_center_x + local_offset_x +
                            swapped_direction * HALF_TARGET_WIDENING_M),
        ]

    # START/SELECT могут поменяться сторонами.
    # Центральная кнопка PS должна остаться на оси симметрии.
    if role == "centre":
        if abs(x - plane_x) <= 0.010:
            return [make_hypothesis("original", x)]

        return [make_hypothesis("original", x), make_hypothesis("mirrored", 2.0 * plane_x - x)]

    # Стики раздвигаются вместе с соответствующими сторонами.
    if role == "sticks":
        direction = 1.0 if x > plane_x else -1.0

        return [
            make_hypothesis("original", x),
            make_hypothesis("widened", x + direction * HALF_TARGET_WIDENING_M),
        ]

    # Для триггеров и бамперов проверяем не только расширение,
    # но и перестановку их назначений на каждой стороне.
    if role in ("triggers", "bumpers"):
        direction = 1.0 if x > plane_x else -1.0

        other_role = ("bumpers" if role == "triggers" else "triggers")

        reference_by_id = {body["id"]: body for body in baseline["bodies"]}

        # Находим соответствующий элемент другой группы
        # с той же стороны корпуса.
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
    Сравнивает относительное расположение четырёх тел в кластере.

    Учитываем только координаты X и Z центров масс. Перед сравнением
    вычитаем центр кластера, поэтому его перенос по модели не мешает.

    Проверяем все 4 перестановок тел и выбираем ту, при которой
    среднеквадратичное расстояние между относительными координатами
    минимально. Значения возвращаем в миллиметрах.
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
    Выделяет два четырёхэлементных кластера кнопок по геометрии.
    Результат разделён на левый и правый кластеры относительно центральной плоскости модели.

    Фильтры подобрны вручную исходя из габаритов и замеров.
    """

    # Ищем тела в области кнопок: небольшие твёрдые тела,
    # расположенные около плоскости Y=0 и не слишком далеко по Z.
    plane_x_mm = plane_x_m * 1000.0
    candidates = []

    for body in records:
        x, y, z = (value * 1000.0 for value in body["centroid_m"])
        volume_mm3 = body["volume_m3"] * 1e9

        # Отбираем небольшие тела в области кнопок.
        if not (-y_range <= y <= y_range):
            continue
        if not (-z_range <= z <= z_range):
            continue
        if not (v_min <= volume_mm3 <= v_max):
            continue

        # Исключаем центральную область корпуса и центральные кнопки.
        if abs(x - plane_x_mm) < 30.0:
            continue

        candidates.append(body)

    clusters = {"left": [body for body in candidates if body["centroid_m"][0] * 1000.0 < plane_x_mm],
                "right": [body for body in candidates if body["centroid_m"][0] * 1000.0 > plane_x_mm]}

    return clusters


def classify_button_clusters(baseline, clusters):
    """
    Определяет, какой из двух текущих кластеров соответствует эталонному D-pad, а какой - кластеру лицевых кнопок.

    Сравниваем внутреннее пространственное расположение тел.
    Выбор делается глобально для двух кластеров: нельзя назначить обе роли одному и тому же кластеру.
    """

    if len(clusters["left"]) != 4 or len(clusters["right"]) != 4:
        return {
            "confident": False,
            "reason": (
                "Ожидалось по четыре тела в каждом кластере, получено:"
                f" left={len(clusters['left'])},"
                f" right={len(clusters['right'])}"
            ),
        }

    reference_by_id = {body["id"]: body for body in baseline["bodies"]}

    reference_groups = {role: [reference_by_id[body_id] for body_id in baseline["roles"][role]]
                        for role in ("dpad", "face_buttons")}

    # Найдем подходящие роли и сохраним их индексы на будущее
    layout_details = {
        role: {
            # Для каждой роли измеряем отличие формы кластера от левого и правого текущих кластеров.
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

    # Рассматриваем оба возможных назначения.
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

    # Сохраняем второго, чтобы рассчитать разницу и точно убедиться
    best = assignments[0]
    second = assignments[1]

    # Порог пока диагностический. Подтвердим его на негативных моделях.
    margin_mm = second["total_rms_mm"] - best["total_rms_mm"]
    confident = (best["total_rms_mm"] <= 2.0 and margin_mm >= 1.0)

    # Если классификация уверенная, сохраняем соответствие каждой эталонной кнопки конкретному текущему телу.
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
    """Находит назначение с минимальной суммарной стоимостью. Строк не больше, чем столбцов."""

    n = len(cost_matrix)
    if not n:
        return {}

    m = len(cost_matrix[0])
    if n > m or any(len(row) != m for row in cost_matrix):
        raise ValueError("Некорректный размер матрицы стоимости.")

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
    """Евклидово расстояние между двумя центрами масс, заданными в метрах."""
    return math.sqrt(sum((a - b) ** 2 for a, b in zip(point_a, point_b))) * 1000.0


def _centroid_distance_xz_mm(point_a, point_b):
    """Расстояние между центрами в плоскости X-Z, в миллиметрах."""
    return math.hypot(point_a[0] - point_b[0], point_a[2] - point_b[2]) * 1000.0


def _body_boxes_may_intersect(body_a, body_b, padding_m=1e-5):
    """Быстро отбрасывает пары тел с разнесёнными габаритными коробками."""
    box_a = [float(v) for v in body_a.GetBodyBox()]
    box_b = [float(v) for v in body_b.GetBodyBox()]

    for axis in range(3):
        overlap = min(box_a[axis + 3], box_b[axis + 3]) - max(box_a[axis], box_b[axis])
        if overlap < -padding_m:
            return False

    return True


def body_intersection_volume_m3(body_a, body_b):
    """Возвращает объём пересечения двух тел в м³, не изменяя исходную геометрию."""
    if not _body_boxes_may_intersect(body_a, body_b):
        return 0.0

    # Булева операция работает с временными копиями тел.
    temp_a = body_a.Copy()
    temp_b = body_b.Copy()

    # PyWin32 автоматически возвращает out-параметр error_code вторым элементом.
    intersection_bodies, error_code = temp_a.Operations2(SW_BODY_INTERSECT, temp_b)
    error_code = int(error_code)
    # Известные коды отсутствия пересечения не считаем ошибкой расчёта.
    if error_code in (5, 6, 1067):  # - пересечений НЕТ
        # Коды: 5 - тела не пересекаются, 6 - пустое тело, 1067 - пересечения нет.
        return 0.0

    if error_code != 0:
        # Код 0 - операция без ошибки. Иначе - ошибка.
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
    """Возвращает ID эталонных тел конкретной роли, включая unclassified."""
    if role == "unclassified":
        assigned = {body_id for ids in baseline["roles"].values() for body_id in ids}
        return [body["id"] for body in baseline["bodies"] if body["id"] not in assigned]
    if role not in baseline["roles"]:
        raise KeyError(f"Неизвестная роль: {role}")
    return list(baseline["roles"][role])


def get_reference_bodies(baseline, role=None):
    """Возвращает все эталонные тела либо тела указанной роли."""
    bodies_by_id = {body["id"]: body for body in baseline["bodies"]}
    if role is None:
        return list(baseline["bodies"])
    return [bodies_by_id[body_id] for body_id in get_role_ids(baseline, role)]


def get_matched_bodies_by_role(baseline, matches, role, include_missing=False):
    """Возвращает текущие тела, сопоставленные эталонной роли, с ключами по ID эталона."""
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
    """Фильтрует тела по геометрии. Интервалы включают обе границы; центр проверяется по евклидову расстоянию."""
    criteria = (centroid_mm, x_range_mm, y_range_mm, z_range_mm, volume_range_mm3, area_range_mm2, dimensions_range_mm,
                face_count_range, small_face_count_range, predicate)
    if not any(value is not None for value in criteria):
        raise ValueError("Нужно задать хотя бы один геометрический признак.")

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
    """Имя тела для диагностики: подходит и для input.json, и для текущей модели."""
    return str(body.get("name_debug", body.get("name", "")))


def body_centroid_mm(body):
    """Центр масс тела в миллиметрах."""
    return tuple(value * 1000.0 for value in body["centroid_m"])


def body_volume_mm3(body):
    """Объём тела в мм³."""
    return body["volume_m3"] * 1e9


def body_area_mm2(body):
    """Площадь поверхности тела в мм²."""
    return body["area_m2"] * 1e6


def body_bbox_mm(body):
    """Габаритная коробка в формате xmin, ymin, zmin, xmax, ymax, zmax."""
    return tuple(value * 1000.0 for value in body["bbox_m"])


def body_dimensions_mm(body):
    """Три отсортированных размера габаритной коробки в мм."""
    bbox = body_bbox_mm(body)
    return tuple(sorted((bbox[3] - bbox[0], bbox[4] - bbox[1], bbox[5] - bbox[2])))
...  # ========================= /Dict helpers =========================


# ========================= Attempts to analyze glyphs =========================
def get_feature_sketch(feature):
    """Пытается получить объект ISketch из операции или её данных."""

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
    """Выводит блоки эскиза, их определения, заметки и геометрические элементы."""

    segments = sw_as_list(sw_value(sketch, "GetSketchSegments"))
    blocks = sw_as_list(sw_value(sketch, "GetSketchBlockInstances"))

    print_if_debug(f"    Сегментов эскиза: {len(segments)}; блоков: {len(blocks)}")

    for block in blocks:
        definition = sw_value(block, "Definition")
        block_name = sw_value(block, "Name", "<без имени>")
        definition_name = sw_value(definition, "Name", "<без определения>")

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

        print_if_debug(f"      Блок {block_name!r}, определение {definition_name!r}")
        print_if_debug(f"        Геометрия определения: {counts}")
        print_if_debug(f"        Текстовые заметки: {note_texts or 'не найдены'}")


def inspect_feature_faces(feature):
    """Показывает площадь граней, принадлежащих операции построения."""

    faces = sw_as_list(sw_value(feature, "GetFaces"))
    areas = []

    for face in faces:
        area = sw_value(face, "GetArea")
        if area is not None:
            areas.append(float(area) * 1e6)

    areas.sort()

    print_if_debug(f"    Граней операции: {len(faces)}")
    print_if_debug(f"    Площади граней, мм²: {[round(a, 4) for a in areas[:30]]}")

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
            f"    Габариты граней операции, мм: "
            f"X=[{xmin * 1000:.2f}, {xmax * 1000:.2f}], "
            f"Y=[{ymin * 1000:.2f}, {ymax * 1000:.2f}], "
            f"Z=[{zmin * 1000:.2f}, {zmax * 1000:.2f}]"
        )


def diagnose_glyph_features(document):
    """
    Ищет операции и эскизы, потенциально связанные с буквами L/R,
    надписями START/SELECT и блоками эскиза.
    """

    print_if_debug("\n" + "=" * 90)
    print_if_debug("ДИАГНОСТИКА НАДПИСЕЙ И ЭСКИЗОВ")
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

        # Показываем все операции вытягивания, линии разделения,
        # нужные эскизы и любые эскизы, содержащие блоки.
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
    """Запускает или находит SolidWorks и открывает нужный документ."""
    print_if_debug("Запуск/Поиск SolidWorks...")

    # =============== SolidWorks ===============
    try:
        # Подключаемся к уже запущенному SolidWorks.
        sw_app = win32com.client.GetActiveObject("SldWorks.Application")
        if sw_app.Visible:
            print_if_debug("Найден запущенный видимый SolidWorks. Если Вы закроете его - будет работать быстрее :)")
    except Exception as exc:
        print_if_debug("Запущенный SolidWorks не найден. Попытка нового headless запуска...")
        try:
            # Если не запущен, создаем новый процесс
            sw_app = win32com.client.Dispatch("SldWorks.Application")
            print_if_debug("SolidWorks успешно запущен (или найден в headless).")
        except Exception as e:
            print_if_debug(f"Ошибка: Не удалось запустить SolidWorks. {e}")
            raise RuntimeError("Не удалось подключиться к SolidWorks. Проверьте его доступность на ПК.") from exc

    # =============== Document ===============
    if document_path:
        model_path = Path(document_path).expanduser().resolve()

        if not model_path.is_file():
            raise FileNotFoundError(f"Файл модели не найден: {model_path}")

        if model_path.suffix.lower() != ".sldprt":
            raise ValueError("Ожидался файл детали SolidWorks (*.SLDPRT).")

        # Если нужный файл уже активен, повторно открывать его не нужно.
        active_doc = sw_app.ActiveDoc
        active_path = (active_doc.GetPathName if active_doc is not None else "")

        # Проверка совпадения путей
        same_path = (active_path and str(Path(active_path).resolve()).casefold() == str(model_path).casefold())

        if same_path:
            document = active_doc
        else:
            # NOTE: 'OpenDoc7' - принимает всего один объект спецификации и более удобный для работы внутри Python

            # Создаем объект спецификации документа
            open_spec = sw_app.GetOpenDocSpec(str(model_path))

            # Настраиваем свойства внутри объекта
            open_spec.DocumentType = SW_DOC_PART
            # open_spec.Silent = False

            # Передаем объект спецификации в OpenDoc7 и открываем переданную модель без диалогов.
            document = sw_app.OpenDoc7(open_spec)

            if document is not None:
                print_if_debug(f"SolidWorks успешно открыл файл: '{document.GetTitle}'")
            else:
                # Ошибки хранятся прямо в объекте спецификации
                raise RuntimeError(f"SolidWorks не открыл {model_path}. Код ошибки={open_spec.Error}")

            # NOTE: Если был Headless запуск, то можно вернуть видимость:
            # document.Visible = True  # OR sw_app.Visible
    else:
        document = sw_app.ActiveDoc

        if document is None:
            raise RuntimeError("В SolidWorks нет активного документа.")

    print_if_debug(f"Работаем с документом: {document.GetTitle}")

    return sw_app, document


def close_sw(sw_app):
    # Закрываем текущий документ (можно также передать имя .SLDPRT файла)
    sw_app.CloseDoc(sw_app.ActiveDoc.GetTitle)

    # Закрываем само приложение (если его открыли мы как headless)
    if CLOSE_OUR_SW_IN_THE_END and not sw_app.Visible:
        print_if_debug("\nЗакрываем headless окно SolidWorks...\n")
        sw_app.ExitApp()


def sw_value(obj, prop, default=None):
    """Безопасно читает COM-свойство."""
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
    """Преобразует результат COM-массива в обычный список."""
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
    Обходит дерево операций SolidWorks, включая подоперации.
    Имена используются только для диагностики, не для оценки.
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
    """Преобразует SAFEARRAY или VARIANT с массивом в обычный список."""
    if hasattr(value, "value"):
        value = value.value
    if value is None:
        return []
    if isinstance(value, (tuple, list)):
        return list(value)
    return [value]


def get_feature_error(feature):
    """Возвращает код ошибки операции и признак предупреждения."""
    is_warning = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_BOOL, False)
    error_code = int(feature.GetErrorCode2(is_warning) or 0)
    return error_code, bool(is_warning.value) if error_code else False


def get_whats_wrong(document):
    """Читает список проблем из окна What's Wrong в SolidWorks."""
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
    Извлекает геометрические характеристики одного твёрдого тела.

    Все величины сохраняются в единицах API SolidWorks:
    координаты - метры, объём - м³, площадь - м².
    Это позволит напрямую сравнивать данные с input.json.
    """
    print_if_debug(f"{index + 1} ", end="")

    # Все размеры сохраняем из API в исходных единицах измерения.
    # GetMassProperties возвращает массив геометрических характеристик.
    #
    # Порядок элементов массива документирован в API SolidWorks:
    # 0..2 - центр масс XYZ, 3 - объём, 4 - площадь,
    # 5 - масса, 6..11 - компоненты момента инерции.
    props = body.GetMassProperties(1.0)

    if props is None or len(props) < 12:
        raise RuntimeError(f"\nНе удалось получить свойства тела с индексом {index}")

    # Габаритный параллелепипед в координатах документа.
    # Это приближённые габариты, а не точная геометрия поверхности.
    # xmin, ymin, zmin, xmax, ymax, zmax.
    bbox = body.GetBodyBox()

    if bbox is None or len(bbox) != 6:
        raise RuntimeError(f"\nНе удалось получить габариты тела с индексом {index}")

    # Получаем все грани тела, чтобы кроме общего числа граней
    # посчитать мелкие грани - один из признаков рельефных символов.
    faces = sw_as_list(body.GetFaces())
    small_face_count = sum(1 for face in faces if float(face.GetArea) < SMALL_FACE_AREA_M2)

    # Имя оставляем исключительно для отладки.
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
    """Извлекает геометрию всех Solid Bodies открытой детали."""

    if int(document.GetType) != SW_DOC_PART:
        raise RuntimeError("Ожидался документ SolidWorks Part (*.SLDPRT).")

    # NOTE: второй аргумнт означает - нужны ли нам только видимые тела
    bodies = sw_as_list(document.GetBodies2(SW_SOLID_BODY, False))
    # И также получаем все плоские тела
    sheet_bodies = com_array_to_list(document.GetBodies2(1, False))

    if not bodies:
        raise RuntimeError("В документе не найдено ни одного Solid Body.")

    print_if_debug(f"\nАнализ тел в документе (всего тел={len(bodies)}):")
    result = [extract_body_geometry(body, index) for index, body in enumerate(bodies)]
    print_if_debug("")
    return result, sheet_bodies


def print_document_geometry(records):
    print_if_debug(f"\nКоличество Solid Bodies: {len(records)}")
    print_if_debug("-" * 100)

    # Сортируем только отображение, чтобы было проще изучать геометрию.
    # Исходные индексы в данных при этом не изменяются.
    for rec in sorted(records, key=lambda item: item["centroid_m"][0]):
        # COM и габариты переводим из метров в миллиметры только для удобного чтения человеком.
        x, y, z = (v * 1000 for v in rec["centroid_m"])

        # Для объёма: м³ -> мм³; для площади: м² -> мм².
        volume = rec["volume_m3"] * 1e9
        area = rec["area_m2"] * 1e6

        print_if_debug(
            f"#{rec['index']:02d} "
            f"{rec['name_debug'][:22]:22} | "
            f"COM=({x:8.2f}, {y:8.2f}, {z:8.2f}) мм | "
            f"V={volume:12.1f} мм³ | "
            f"A={area:10.1f} мм² | "
            f"faces={rec['face_count']} | "
            f"s_faces={rec['small_face_count']}"
        )

    # Максимальный объём используем только для диагностической проверки.
    largest = max(records, key=lambda item: item["volume_m3"])

    print_if_debug("-" * 100)
    print_if_debug(f"Самое объёмное тело: {largest['name_debug']}; V={largest['volume_m3'] * 1e9:.1f} мм³")
...  # ========================= /SolidWorks API =========================


def score_model(document):
    # diagnose_glyph_features(document); return  # Тут, в теории, можно подробнее изучить надписи и их отзеркаливание

    # ============= Получаем все твёрдые тела =============
    records, sheet_bodies = extract_document_geometry(document)
    print_document_geometry(records)

    # ============= Загружаем эталон =============
    baseline, unassigned_ids = load_baseline()
    print_baseline(baseline, unassigned_ids)

    # ============= Сопоставляем детали =============
    global_result = match_bodies_global(baseline, records)
    print_global_match_report(baseline, global_result)

    # ========================== ПРОВОДИМ ПРОВЕРКИ ==========================
    # --------- Дерево построения ---------
    tree_report = diagnose_feature_tree_health(document)
    tree_score = score_feature_tree_health(tree_report)
    print_feature_tree_score(tree_score)

    # --------- Расширение корпуса (по парам) ---------
    widening_report = diagnose_widening_pairs(baseline, global_result["matches"])
    widening_score = score_widening_pairs(widening_report)
    print_widening_pairs_score(widening_score)

    # --------- Положения кластеров кнопок ---------
    diagnose_button_clusters(baseline, records)
    button_layout_score = score_button_clusters(baseline, records, global_result["button_classification"])
    print_button_clusters_score(button_layout_score)

    # --------- Центральные кнопки (|) ---------
    centre_score = score_centre_positions(baseline, global_result["matches"])

    # --------- Отзеркаливание символов (L      - .... R) ---------
    light_cluster_result = diagnose_light_cluster(baseline, records)
    light_score = score_light_cluster(light_cluster_result, baseline)
    print_light_cluster_score(light_score)

    # --------- Наличие фигурных символов на кнопках ---------
    glyph_score = score_glyph_preservation(baseline, global_result["matches"])

    # --------- Пересечения деталей ---------
    interference_result = diagnose_control_interference(baseline, records, sheet_bodies, global_result["matches"])
    interference_score = score_control_interference(interference_result, baseline)

    # Если брать всё - то там кучу с болтами и прочим внутрянками и по итогу получаем: Пересечения  0.874 x 0.10 = 0.087
    # all_intersections_result = diagnose_all_solid_intersections(records)
    # interference_score = score_all_solid_intersections(all_intersections_result)

    # Можем так же проверить на пересечения с другими крупными телами, на случай если это будет важно
    if DEBUG_FULL_OUTPUT:
        diagnose_unmatched_housing_sized_bodies(baseline, records, global_result["matches"])

    # ============= Подсчёт итоговой оценки =============
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
    # Запуск SolidWorks и открытие документа
    sw_app, document = get_sw_and_doc(document_path)

    try:
        # Пробуем оценить модель
        return score_model(document)
    except Exception as e:
        # Выводим ошибку в том же месте, где она появилась
        raise e
    finally:
        # Закрываем текущий документ
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
        print_if_debug("Режим отладки ВКЛЮЧЕН.")
        argv.remove("--debug") if "--debug" in argv else argv.remove("-d")

    sys.exit(main(argv))
