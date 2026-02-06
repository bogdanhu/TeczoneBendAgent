import json


def extract_parts(obj):
    parts = []
    if isinstance(obj, dict):
        if "partId" in obj and isinstance(obj.get("partId"), int):
            parts.append(obj)
        for v in obj.values():
            parts.extend(extract_parts(v))
    elif isinstance(obj, list):
        for item in obj:
            parts.extend(extract_parts(item))
    return parts


def load_xometry_map(json_path, logger=None):
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    parts = extract_parts(data)
    result = {}
    for part in parts:
        part_id = part.get("partId")
        if part_id is None:
            continue
        result[part_id] = {
            "material": part.get("material"),
            "processes": part.get("processes"),
            "quantityPieces": part.get("quantityPieces"),
            "fileNameOnPage": part.get("fileNameOnPage"),
            "tolerance": part.get("tolerance"),
            "ra": part.get("ra"),
            "productionRemarks": part.get("productionRemarks"),
            "thicknessMm": part.get("thicknessMm"),
        }

    if logger:
        logger.info("Parsed %d parts from xometry json", len(result))
    return result
