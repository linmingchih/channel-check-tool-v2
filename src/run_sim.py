import re, os
import sys
import json
from pathlib import Path

#json_path = '../data2/simulation.json'
json_path = sys.argv[1]

with open(json_path) as f:
    info = json.load(f)
    
edb_path = info['aedb_path'].replace('.aedb', '_applied.aedb')
from pyaedt import Hfss3dLayout

hfss = Hfss3dLayout(edb_path, version='2025.1', non_graphical=True)

hfss.export_touchstone_on_completion()
hfss.analyze()

root = Path(info['aedb_path'].replace('.aedb', '_applied.aedtexport'))

pattern = re.compile(r"\.s(\d{1,3})p$", re.IGNORECASE)

# 遞迴搜尋並篩選出符合條件的檔案
matched_files = [
    p.resolve() for p in root.rglob("*")
    if p.is_file() and pattern.search(p.name)
]

hfss.release_desktop()

if matched_files:
    touchstone_path = str(matched_files[0])
    print(touchstone_path)
    
    # Save the result to a json file
    output_dir = Path(json_path).parent
    result_json_path = output_dir / "result.json"
    with open(result_json_path, "w") as f:
        json.dump({"touchstone_path": touchstone_path}, f, indent=2)
else:
    print("Error: No Touchstone file found.")
    sys.exit(1)


