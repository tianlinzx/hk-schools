"""Fetch school data from Tencent Docs API and regenerate index.html."""
import json
import os
import sys
import urllib.request

API_BASE = "https://docs.qq.com/openapi/spreadsheet/v3/files"

DISTRICT_MAP = {
    "九龍城": "九龍城區",
    "油尖旺": "油尖旺區",
    "深水埗": "深水埗區",
    "黃大仙": "黃大仙區",
}

ADDRESS_DISTRICT_HINTS = [
    ("旺角", "油尖旺區"), ("油麻地", "油尖旺區"), ("佐敦", "油尖旺區"),
    ("九龍塘", "九龍城區"), ("九龍城", "九龍城區"), ("土瓜灣", "九龍城區"), ("紅磡", "九龍城區"),
    ("銅鑼灣", "灣仔區"), ("灣仔", "灣仔區"), ("大坑", "灣仔區"),
    ("將軍澳", "西貢區"), ("寶林", "西貢區"),
    ("元朗", "元朗區"), ("天水圍", "元朗區"),
    ("屯門", "屯門區"),
    ("沙田", "沙田區"), ("大圍", "沙田區"),
    ("深水埗", "深水埗區"), ("石硤尾", "深水埗區"), ("美孚", "深水埗區"), ("界限街", "深水埗區"), ("白田街", "深水埗區"),
    ("荃灣", "荃灣區"),
    ("青衣", "葵青區"),
    ("中環", "中西區"), ("西營盤", "中西區"), ("般咸道", "中西區"), ("堅道", "中西區"), ("皇后大道", "中西區"), ("域多利道", "中西區"),
    ("觀塘", "觀塘區"),
    ("黃大仙", "黃大仙區"), ("慈雲山", "黃大仙區"),
    ("北角", "東區"), ("柴灣", "東區"), ("小西灣", "東區"),
    ("香港仔", "南區"), ("薄扶林", "南區"), ("赤柱", "南區"), ("黃竹坑", "南區"), ("田灣", "南區"),
    ("大埔", "大埔區"),
]


def fetch_data():
    client_id = os.environ["TD_CLIENT_ID"]
    access_token = os.environ["TD_ACCESS_TOKEN"]
    open_id = os.environ["TD_OPEN_ID"]
    file_id = os.environ["TD_FILE_ID"]
    sheet_id = os.environ["TD_SHEET_ID"]

    url = f"{API_BASE}/{file_id}/{sheet_id}/A1:J600"
    req = urllib.request.Request(url, headers={
        "Access-Token": access_token,
        "Open-Id": open_id,
        "Client-Id": client_id,
    })

    with urllib.request.urlopen(req, timeout=30) as resp:
        data = json.loads(resp.read().decode("utf-8"))

    grid = data.get("gridData", {})
    rows = grid.get("rows", [])
    if not rows:
        print("ERROR: No data returned from API", file=sys.stderr)
        sys.exit(1)

    # Parse header
    headers = []
    for cell in rows[0].get("values", []):
        cv = cell.get("cellValue") or {}
        headers.append(cv.get("text", "").strip())

    # Parse data rows
    schools = []
    for row in rows[1:]:
        vals = []
        for cell in row.get("values", []):
            cv = cell.get("cellValue")
            if cv is None:
                vals.append("")
            elif "text" in cv:
                vals.append(cv["text"].strip())
            elif "number" in cv:
                vals.append(str(cv["number"]))
            elif "link" in cv:
                link = cv["link"]
                vals.append(link.get("url", str(link)) if isinstance(link, dict) else str(link))
            else:
                vals.append(str(cv))

        if not any(v.strip() for v in vals):
            continue

        school = {
            "name": vals[0] if len(vals) > 0 else "",
            "district": vals[1] if len(vals) > 1 else "",
            "level": vals[2] if len(vals) > 2 else "",
            "gender": vals[3] if len(vals) > 3 else "",
            "type": vals[4] if len(vals) > 4 else "",
            "status": vals[5] if len(vals) > 5 else "",
            "deadline": vals[6] if len(vals) > 6 else "",
            "address": vals[7] if len(vals) > 7 else "",
            "website": vals[8] if len(vals) > 8 else "",
            "phone": vals[9] if len(vals) > 9 else "",
        }
        schools.append(school)

    # Clean data
    for s in schools:
        s["district"] = DISTRICT_MAP.get(s["district"], s["district"])
        if not s["level"]:
            s["level"] = "小学"
        if not s["gender"]:
            s["gender"] = "男女"
        if not s["district"] and s["address"]:
            for keyword, district in ADDRESS_DISTRICT_HINTS:
                if keyword in s["address"]:
                    s["district"] = district
                    break

    return schools


def generate_html(schools):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    template_before = os.path.join(script_dir, "template_before.html")
    template_after = os.path.join(script_dir, "template_after.html")

    with open(template_before, "r", encoding="utf-8") as f:
        before = f.read()
    with open(template_after, "r", encoding="utf-8") as f:
        after = f.read()

    data_js = json.dumps(schools, ensure_ascii=False, separators=(",", ":"))
    data_line = f"var SCHOOLS = {data_js};\n"

    return before + data_line + after


def main():
    schools = fetch_data()
    print(f"Fetched {len(schools)} schools")

    html = generate_html(schools)

    script_dir = os.path.dirname(os.path.abspath(__file__))
    output = os.path.join(script_dir, "index.html")
    with open(output, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"Generated {output} ({len(html)} bytes)")


if __name__ == "__main__":
    main()
