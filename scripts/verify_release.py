#!/usr/bin/env python3
import csv
import hashlib
import json
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

def sha256(path):
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()

def read_csv(name):
    with (ROOT / "data" / name).open("r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))

checks = {}
checksum_file = ROOT / "CHECKSUMS.sha256"
checksum_rows = [line.split("  ", 1) for line in checksum_file.read_text("utf-8").splitlines() if line.strip()]
checks["checksums_all_match"] = all((ROOT / rel).is_file() and sha256(ROOT / rel) == digest for digest, rel in checksum_rows)

desc = read_csv("week9_descriptive_statistics.csv")
variables = read_csv("week9_variable_dictionary.csv")
availability = read_csv("week9_variable_availability.csv")
long_rows = read_csv("week9_company_variable_long.csv")
integrated = read_csv("week9_integrated_variable_map.csv")
profiles = read_csv("week9_company_profiles.csv")
portfolios = read_csv("week9_candidate_portfolios.csv")

checks.update({
    "descriptive_rows_6": len(desc) == 6,
    "variable_rows_32": len(variables) == 32,
    "availability_rows_32": len(availability) == 32,
    "long_rows_256": len(long_rows) == 256,
    "integrated_rows_32": len(integrated) == 32,
    "profile_rows_8": len(profiles) == 8,
    "portfolio_rows_8": len(portfolios) == 8,
})

status = Counter(r["final_status"] for r in long_rows)
checks["status_counts_exact"] = status == Counter({"VALID": 215, "STRUCTURAL_NA": 38, "NOT_COMPUTABLE": 3})
checks["repaired_cells_10"] = sum(bool(r["repair_stage"].strip()) for r in long_rows) == 10
checks["composite_keys_unique"] = len({r["composite_key"] for r in long_rows}) == 256
checks["company_ids_8"] = len({r["company_id"] for r in long_rows}) == 8
checks["variable_ids_32"] = len({r["variable_id"] for r in long_rows}) == 32
checks["full_company_variable_grid"] = all(sum(1 for r in long_rows if r["company_id"] == cid) == 32 for cid in {r["company_id"] for r in long_rows})
not_comp = [r for r in long_rows if r["final_status"] == "NOT_COMPUTABLE"]
checks["not_computable_all_null"] = len(not_comp) == 3 and all(not r["final_value"].strip() for r in not_comp)
checks["not_computable_insta360_only"] = {r["company_id"] for r in not_comp} == {"688775"}
checks["not_computable_variables_exact"] = {r["variable_id"] for r in not_comp} == {"CV003", "CV004", "CV026"}
checks["no_final_selected"] = all(r.get("Stage6C状态", "") != "FINAL_SELECTED" for r in portfolios)

manifest = json.loads((ROOT / "metadata" / "release_manifest.json").read_text("utf-8"))
checks["manifest_counts_exact"] = manifest["counts"] == {"companies":8,"variables":32,"company_variable_cells":256,"valid":215,"structural_na":38,"not_computable":3,"repaired_cells":10,"candidate_portfolios":8,"final_selected":0}

failed = [k for k, v in checks.items() if not v]
result = {"status": "PASS" if not failed else "FAIL", "checks": checks, "failed": failed, "status_counts": dict(status)}
print(json.dumps(result, ensure_ascii=False, indent=2))
sys.exit(1 if failed else 0)
