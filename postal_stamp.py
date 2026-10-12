"""Local incoming-stamp localization, with text evidence before confirmation.

Numbers alone never confirm a stamp. This helper supplements page OCR when
small printed stamp labels disappear during whole-page detection.
"""
import difflib
import math
import re


def _key(value, norm):
    text = norm(str(value or "")) if norm else str(value or "")
    text = text.translate(str.maketrans({"أ": "ا", "إ": "ا", "آ": "ا",
                                       "ة": "ه", "ى": "ي"}))
    return re.sub(r"[^\u0621-\u064a0-9]", "", text)


def _similar(text, target):
    if target in text:
        return 1.0
    if not text:
        return 0.0
    if len(text) <= len(target) + 5:
        return difflib.SequenceMatcher(None, text, target).ratio()
    return max(difflib.SequenceMatcher(None, text[i:i + len(target)], target).ratio()
               for i in range(len(text) - len(target) + 1))


def _incoming(text, norm):
    key = _key(text, norm)
    if "صادر" in key or "عامل" in key:
        return False
    return len(key) >= 6 and _similar(key, "رقمالوارد") >= .72


def _sector(text, norm):
    key = _key(text, norm)
    return len(key) >= 10 and _similar(key, "قطاعالتنميهالاداريه") >= .66


def _overlap_text(box, lines):
    x0, y0, x1, y1 = box
    return [line for line in lines
            if x0 <= (line.get("x0", 0) + line.get("x1", 0)) / 2 <= x1
            and y0 <= (line.get("y0", 0) + line.get("y1", 0)) / 2 <= y1]


def _cluster(values, tolerance):
    groups = []
    for value in sorted(values):
        if groups and abs(value - sum(groups[-1]) / len(groups[-1])) <= tolerance:
            groups[-1].append(value)
        else:
            groups.append([value])
    return [sum(group) / len(group) for group in groups]


def _geometry(gray):
    """Propose small ruled tables anywhere; OCR still decides their purpose."""
    import cv2
    height, width = gray.shape
    scale = min(1., 1300. / width)
    image = cv2.resize(gray, (round(width * scale), round(height * scale)))
    h, w = image.shape
    candidates = []
    for offset in (2, 9):
        binary = cv2.adaptiveThreshold(image, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                       cv2.THRESH_BINARY_INV, 41, offset)
        raw = cv2.HoughLinesP(binary, 1, math.pi / 180, threshold=35,
                             minLineLength=round(w * .05), maxLineGap=round(w * .025))
        horizontals, verticals = [], []
        for segment in ([] if raw is None else raw.reshape(-1, 4)):
            x0, y0, x1, y1 = map(float, segment)
            dx, dy = x1 - x0, y1 - y0
            if abs(dx) >= w * .13 and abs(dy) <= abs(dx) * .10:
                if x1 < x0:
                    x0, y0, x1, y1 = x1, y1, x0, y0
                horizontals.append((x0, y0, x1, y1))
            elif abs(dy) >= h * .035 and abs(dx) <= abs(dy) * .12:
                if y1 < y0:
                    x0, y0, x1, y1 = x1, y1, x0, y0
                verticals.append((x0, y0, x1, y1))
        # Quantize repeated detections before pairing them.
        unique = {}
        for line in horizontals:
            key = tuple(round(v / 6) for v in line)
            unique[key] = line
        horizontals = list(unique.values())[:450]
        for a in horizontals:
            ay = (a[1] + a[3]) / 2
            for b in horizontals:
                by = (b[1] + b[3]) / 2
                if not h * .055 <= by - ay <= h * .15:
                    continue
                left, right = max(a[0], b[0]), min(a[2], b[2])
                if not w * .15 <= right - left <= w * .35:
                    continue
                if abs(a[0] - b[0]) > w * .03 or abs(a[2] - b[2]) > w * .03:
                    continue
                slope_a = (a[3] - a[1]) / (a[2] - a[0])
                slope_b = (b[3] - b[1]) / (b[2] - b[0])
                if abs(slope_a - slope_b) > .04:
                    continue
                top, bottom = min(a[1], a[3]), max(b[1], b[3])
                edge_positions = []
                for vx0, vy0, vx1, vy1 in verticals:
                    vx = (vx0 + vx1) / 2
                    coverage = max(0, min(vy1, bottom) - max(vy0, top)) / (bottom - top)
                    if coverage >= .45 and (abs(vx - left) < w * .018
                                             or abs(vx - right) < w * .018):
                        edge_positions.append(vx)
                edges = _cluster(edge_positions, w * .012)
                if len(edges) < 2:
                    continue
                rows = _cluster([(r[1] + r[3]) / 2 for r in horizontals
                                 if r[0] < left + w * .03 and r[2] > right - w * .03
                                 and ay - 12 <= (r[1] + r[3]) / 2 <= by + 12], 9)
                if len(rows) < 3:
                    continue
                box = [max(0, round((min(a[0], b[0]) - 8) / scale)),
                       max(0, round((top - 15) / scale)),
                       min(width, round((max(a[2], b[2]) + 8) / scale)),
                       min(height, round((bottom + 15) / scale))]
                candidates.append({"box": box, "slope": (slope_a + slope_b) / 2,
                                   "row_boundaries": [round(y / scale) for y in rows],
                                   "geometry": True, "strength": min(len(rows), 7)})
    return candidates


def _proposals(gray, lines, norm, ar_digits):
    height, width = gray.shape
    proposals = []
    for line in lines:
        if not all(k in line for k in ("x0", "x1", "y0", "y1")):
            continue
        text = line.get("text", "")
        if _incoming(text, norm):
            proposals.append({"box": [max(0, round(line["x0"] - width * .15)),
                                      max(0, round(line["y0"] - height * .03)),
                                      min(width, round(line["x1"] + width * .04)),
                                      min(height, round(line["y1"] + height * .075))],
                              "slope": 0., "geometry": False, "strength": 100})
        converted = text.translate(ar_digits)
        numbers = re.findall(r"(?<!\d)\d{4}(?!\d)", converted)
        if not numbers or all(1900 <= int(n) <= 2099 for n in numbers):
            continue
        if line["x1"] - line["x0"] > width * .35:
            continue
        if re.search(r"عامل|مرتب|راتب|المرقى|حرفيين", text):
            continue
        # This is only a search window. Its number cannot confirm anything.
        proposals.append({"box": [max(0, round(line["x0"] - width * .035)),
                                  max(0, round(line["y0"] - height * .075)),
                                  min(width, round(line["x1"] + width * .04)),
                                  min(height, round(line["y1"] + height * .045))],
                          "slope": 0., "geometry": False, "strength": 4})
    anchored = any(p["strength"] == 100 for p in proposals)
    # Whole-page Hough is unnecessary for an already legible incoming label.
    if not anchored:
        proposals += _geometry(gray)
    for candidate in proposals:
        local = _overlap_text(candidate["box"], lines)
        joined = " ".join(row.get("text", "") for row in local)
        if any(_incoming(row.get("text", ""), norm) for row in local):
            candidate["strength"] += 100
        if any(_sector(row.get("text", ""), norm) for row in local):
            candidate["strength"] += 25
        if re.search(r"رقم\s*العامل|المرتب|المرقى|حرفيين|علاوة", joined):
            candidate["strength"] -= 50
        if "صادر" in joined:
            candidate["strength"] -= 50
    selected = []
    for candidate in sorted(proposals,
                            key=lambda c: (c["strength"], c["box"][3] - c["box"][1]),
                            reverse=True):
        if candidate["strength"] < 0:
            continue
        x0, y0, x1, y1 = candidate["box"]
        area = max(1, (x1 - x0) * (y1 - y0))
        duplicate = False
        for previous in selected:
            bx0, by0, bx1, by1 = previous["box"]
            intersection = max(0, min(x1, bx1) - max(x0, bx0)) * max(0, min(y1, by1) - max(y0, by0))
            if intersection / min(area, max(1, (bx1 - bx0) * (by1 - by0))) > .83:
                duplicate = True
                break
        if not duplicate:
            selected.append(candidate)
        if len(selected) >= 3:
            break
    return selected


def _field_windows(candidate, lines, page_width, page_height, ar_digits):
    """Keep handwriting that extends past a ruled border inside the date crop."""
    left, top, right, bottom = candidate["box"]
    width, height = right - left, bottom - top
    date_box = [max(0, round(left - page_width * .04)),
                max(0, round(top + height * .59 - 12)),
                min(page_width, round(right + page_width * .015)),
                min(page_height, round(bottom + 14))]
    mid = left + width * .46
    number_rows = []
    for row in _overlap_text(candidate["box"], lines):
        values = re.findall(r"(?<!\d)\d{4}(?!\d)", row.get("text", "").translate(ar_digits))
        if any(not 1900 <= int(value) <= 2099 for value in values):
            if (row["x0"] + row["x1"]) / 2 >= mid:
                number_rows.append(row)
    if len(number_rows) == 1:
        row = number_rows[0]
        number_box = [max(0, round(max(mid, row["x0"] - 18))),
                      max(0, round(row["y0"] - 14)),
                      min(page_width, round(max(right, row["x1"] + 16))),
                      min(page_height, round(row["y1"] + 14))]
    else:
        number_box = [round(mid), max(0, round(top + height * .31 - 10)),
                      min(page_width, round(right + 12)),
                      min(page_height, round(top + height * .62 + 10))]
    return {"number_box": number_box, "date_box": date_box}


def _local_grid(gray, box):
    """Read the small grid inside a proposed window, including faint borders."""
    import cv2
    x0, y0, x1, y1 = box
    crop = gray[y0:y1, x0:x1]
    h, w = crop.shape
    horizontal, vertical = [], []
    for offset in (2, 9):
        binary = cv2.adaptiveThreshold(crop, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                       cv2.THRESH_BINARY_INV, 41, offset)
        raw = cv2.HoughLinesP(binary, 1, math.pi / 180, threshold=25,
                             minLineLength=round(min(w, h) * .2),
                             maxLineGap=round(w * .09))
        for line in ([] if raw is None else raw.reshape(-1, 4)):
            ax, ay, bx, by = map(float, line)
            dx, dy = bx - ax, by - ay
            if abs(dx) >= w * .62 and abs(dy) <= abs(dx) * .10:
                horizontal.append((ax, ay, bx, by))
            elif abs(dy) >= h * .45 and abs(dx) <= abs(dy) * .12:
                vertical.append((ax, ay, bx, by))
    if len(horizontal) < 4 or len(vertical) < 2:
        return {}
    edges = _cluster([(a + c) / 2 for a, b, c, d in vertical], w * .025)
    if len(edges) < 2 or edges[-1] - edges[0] < w * .65:
        return {}
    rows = _cluster([(b + d) / 2 for a, b, c, d in horizontal], max(10, h * .035))
    if len(rows) < 5 or rows[-1] - rows[0] < h * .45:
        return {}
    left, right = edges[0] + x0, edges[-1] + x0
    interior = [x + x0 for x in edges[1:-1] if left + (right - left) * .35 <= x + x0
                <= left + (right - left) * .70]
    mid = interior[0] if interior else (left + right) / 2
    slopes = [(d - b) / (c - a) for a, b, c, d in horizontal if abs(c - a) > 1]
    slope = sorted(slopes)[len(slopes) // 2]
    rows = [y + y0 for y in rows]
    margin = abs(slope) * (right - left) / 2 + 8
    # This form has a header, labels, numeric value row, source row, date row.
    # Cell hints are only returned for complete five-row grids.
    if len(rows) != 6:
        return {"row_boundaries": rows, "slope": slope,
                "box": [round(left), round(rows[0] - margin),
                        round(right), round(rows[-1] + margin)]}
    return {"row_boundaries": rows, "slope": slope,
            "box": [round(left), round(rows[0] - margin),
                    round(right), round(rows[-1] + margin)],
            "number_box": [round(mid), round(rows[2] - margin),
                           round(right), round(rows[3] + margin)],
            "date_box": [round(left), round(rows[-2] - margin),
                         round(right), round(rows[-1] + margin)]}


def _read_crop(gray, candidate, engine, run_engine, res_lines, page, page_size):
    import cv2
    import numpy as np
    x0, y0, x1, y1 = candidate["box"]
    crop = gray[y0:y1, x0:x1]
    h, w = crop.shape
    angle = math.degrees(math.atan(candidate.get("slope", 0)))
    matrix = cv2.getRotationMatrix2D((w / 2, h / 2), angle, 1.)
    cos, sin = abs(matrix[0, 0]), abs(matrix[0, 1])
    new_w, new_h = round(w * cos + h * sin), round(h * cos + w * sin)
    matrix[0, 2] += (new_w - w) / 2
    matrix[1, 2] += (new_h - h) / 2
    rotated = cv2.warpAffine(crop, matrix, (new_w, new_h), borderValue=255)
    inverse = cv2.invertAffineTransform(matrix)
    factor = min(3., 2200. / max(new_w, new_h))
    variants = [rotated, cv2.createCLAHE(clipLimit=3., tileGridSize=(8, 8)).apply(rotated),
                cv2.adaptiveThreshold(rotated, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                      cv2.THRESH_BINARY, 41, 9)]
    mapped = []
    for variant in variants:
        enlarged = cv2.resize(variant, (round(new_w * factor), round(new_h * factor)),
                              interpolation=cv2.INTER_CUBIC)
        result = run_engine(engine, enlarged, use_cls=False)
        for line in res_lines(result):
            if not all(k in line for k in ("x0", "x1", "y0", "y1")):
                continue
            points = np.array([[line["x0"], line["y0"]], [line["x1"], line["y0"]],
                               [line["x1"], line["y1"]], [line["x0"], line["y1"]]], dtype=float) / factor
            points = np.column_stack([points, np.ones(4)]) @ inverse.T
            row = dict(line, x0=float(points[:, 0].min() + x0),
                       x1=float(points[:, 0].max() + x0),
                       y0=float(points[:, 1].min() + y0),
                       y1=float(points[:, 1].max() + y0),
                       _page=page, _page_size=page_size, _stamp_localized=True)
            if not any(row["text"] == old["text"]
                       and abs(row["x0"] - old["x0"]) < 20
                       and abs(row["y0"] - old["y0"]) < 20 for old in mapped):
                mapped.append(row)
    return mapped


def recover_incoming_stamp(prepped, lines, engine, run_engine, res_lines, norm, ar_digits,
                           anchor=None):
    """Return a locally re-read, text-confirmed stamp with original coordinates."""
    import numpy as np
    from PIL import Image
    with Image.open(prepped) as source:
        gray = np.asarray(source.convert("L"))
        page_size = [source.width, source.height]
    page = next((line.get("_page", 1) for line in lines), 1)
    if anchor and anchor not in lines:
        lines = list(lines) + [anchor]
    layout_result = {}
    for candidate in _proposals(gray, lines, norm, ar_digits):
        grid = _local_grid(gray, candidate["box"])
        if grid and "number_box" in grid:
            # Keep the full proposal. A faint outer border can disappear while
            # an internal rule survives, so a smaller grid must not crop dates.
            candidate = dict(candidate, geometry=True,
                             row_boundaries=grid["row_boundaries"])
        candidate = dict(candidate, **_field_windows(candidate, lines,
                                                     page_size[0], page_size[1], ar_digits))
        mapped = _read_crop(gray, candidate, engine, run_engine, res_lines, page, page_size)
        local = mapped + _overlap_text(candidate["box"], lines)
        incoming_label = any(_incoming(row.get("text", ""), norm) for row in local)
        sector_label = any(_sector(row.get("text", ""), norm) for row in local)
        outgoing_label = any("صادر" in _key(row.get("text", ""), norm) for row in local)
        if outgoing_label and not incoming_label:
            continue
        if incoming_label or (sector_label and candidate.get("geometry")
                              and len(candidate.get("row_boundaries", [])) >= 4):
            result = {"lines": mapped, "box": candidate["box"],
                      "slope": candidate.get("slope", 0.), "confirmed": True,
                      "row_boundaries": candidate.get("row_boundaries", [])}
            for cell in ("number_box", "date_box"):
                if cell in candidate:
                    result[cell] = candidate[cell]
            return result
        if candidate.get("geometry") and len(candidate.get("row_boundaries", [])) >= 5 \
                and not layout_result:
            number_rows = _overlap_text(candidate["number_box"], local)
            values = {n for row in number_rows
                      for n in re.findall(r"(?<!\d)\d{4}(?!\d)",
                                          row.get("text", "").translate(ar_digits))
                      if not 1900 <= int(n) <= 2099}
            unrelated = any(re.search(r"عامل|مرتب|المرقى|حرفيين|علاوة", row.get("text", ""))
                            for row in local)
            if len(values) == 1 and not unrelated:
                layout_result = {"lines": mapped, "box": candidate["box"],
                                 "slope": candidate.get("slope", 0.),
                                 "confirmed": False, "layout_candidate": True,
                                 "number_box": candidate["number_box"],
                                 "date_box": candidate["date_box"],
                                 "row_boundaries": candidate["row_boundaries"]}
    return layout_result
