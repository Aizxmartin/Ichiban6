def calculate_adjustments(row, subject_info, schema):
    try:
        ag_sf = float(row.get("Above Grade Finished Area", 0))
        subj_sf = float(subject_info.get("Above Grade Finished Area", 0))
        close_price = float(row.get("Close Price", 0))
        concessions = float(row.get("Concessions", 0))

        ag_diff = ag_sf - subj_sf
        ag_rate = schema["above_grade"]["rate"]
        ag_adj = ag_diff * ag_rate

        total_adj = ag_adj
        adjusted_price = close_price + concessions + total_adj

        return total_adj, adjusted_price, ag_diff
    except Exception:
        return None, None, None