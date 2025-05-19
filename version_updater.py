import os
import json
import shutil
from datetime import datetime
import sys
from CPA_WCL import resource_path


def migrate_rules():
    """Ensure rules are in the correct persistent location"""
    old_locations = [
        os.path.join(os.path.dirname(sys.executable), "djezzy_rules.json"),
        os.path.join(sys._MEIPASS, "djezzy_rules.json") if hasattr(
            sys, '_MEIPASS') else None
    ]

    new_location = resource_path("djezzy_rules.json")

    # If no rules file exists, create default
    if not os.path.exists(new_location):
        default_rules = {
            "2G": [
                {
                    "kpi": "CSSR_OPTIMUM",
                    "operator": ">",
                    "threshold": 98.0,
                    "count_threshold": 100.0,
                    "count_column": "CS Traffic OPTIMUM Daily"
                },
                {
                    "kpi": "CDR_OPTIMUM",
                    "operator": "<",
                    "threshold": 1.0,
                    "count_threshold": 0.0
                },
                {
                    "kpi": "HSR_OPTIMUM",
                    "operator": ">=",
                    "threshold": 98.0,
                    "count_threshold": 0.0
                }
            ],
            "3G": [
                {
                    "kpi": "Call Drop Rate PS_OPTIMUM",
                    "operator": "<",
                    "threshold": 1.0,
                    "count_threshold": 100.0,
                    "count_column": "CDR_PS_Number_optimum"
                },
                {
                    "kpi": "RTWP_optimum(dBm)",
                    "operator": "<",
                    "threshold": -95.0,
                    "count_threshold": 0
                },
                {
                    "kpi": "Call Setup Success Rate CS_OPTIMUM",
                    "operator": ">=",
                    "threshold": 98.0,
                    "count_threshold": 0.0
                },
                {
                    "kpi": "Call Drop Rate CS_OPTIMUM",
                    "operator": "<",
                    "threshold": 1.0,
                    "count_threshold": 0.0
                }
            ],
            "4G": [
                {
                    "kpi": "LTE Setup Success Rate_OPTIMUM(%)",
                    "operator": ">",
                    "threshold": 99,
                    "count_column": "LTE_Attempts",
                    "count_threshold": 0
                },
                {
                    "kpi": "LTE Call Drop Rate_OPTIMUM",
                    "operator": "<",
                    "threshold": 0.8,
                    "count_threshold": 0
                },
                {
                    "kpi": "CSFB Success Rate_OPTIMUM(%)",
                    "operator": ">",
                    "threshold": 95.5,
                    "count_threshold": 0.0
                }
            ]
        }
        with open(new_location, 'w') as f:
            json.dump(default_rules, f, indent=4)
        return

    # Migrate from old locations if found
    for old_loc in old_locations:
        if old_loc and os.path.exists(old_loc):
            try:
                shutil.copy2(old_loc, new_location)
                print(f"Migrated rules from {old_loc} to {new_location}")
            except Exception as e:
                print(f"Migration failed: {e}")


if __name__ == "__main__":
    migrate_rules()
