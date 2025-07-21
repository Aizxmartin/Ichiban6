from docx import Document
from docx.shared import Inches
from adjustments import calculate_adjustments
import io

def generate_report(df, subject_info, online_estimate_avg):
    from market_adjustment_schema import schema

    adjusted_rows = []

    for _, row in df.iterrows():
        total_adj, adjusted_price, ag_diff = calculate_adjustments(row, subject_info, schema)
        if total_adj is None:
            continue
        adjusted_rows.append({
            "Address": row.get("Street Address", "N/A"),
            "Close Price": row.get("Close Price", 0),
            "Concessions": row.get("Concessions", 0),
            "AG SF": row.get("Above Grade Finished Area", 0),
            "AG Diff": ag_diff,
            "Total Adj": total_adj,
            "Adjusted Price": adjusted_price,
            "Adj PPSF": adjusted_price / row.get("Above Grade Finished Area", 1)
        })

    doc = Document()
    doc.add_heading("Market Valuation Report – Adjusted Comparison with Breakdown", 0)
    doc.add_paragraph(f"Subject Property\nAddress: {subject_info.get('Address')}\n"
                      f"Above Grade SF: {subject_info.get('Above Grade Finished Area')}\n"
                      f"Bedrooms: {subject_info.get('Bedrooms')}\n"
                      f"Bathrooms: {subject_info.get('Bathrooms')}")

    if adjusted_rows:
        doc.add_paragraph("Comparable Properties")
        table = doc.add_table(rows=1, cols=8)
        hdr = table.rows[0].cells
        hdr[:] = [cell.text for cell in [
            "Address", "Close Price", "Concessions", "AG SF",
            "AG Diff", "Total Adj", "Adjusted Price", "Adj PPSF"
        ]]

        for comp in adjusted_rows:
            row_cells = table.add_row().cells
            row_cells[0].text = str(comp["Address"])
            row_cells[1].text = f"${comp['Close Price']:,.0f}"
            row_cells[2].text = f"${comp['Concessions']:,.0f}"
            row_cells[3].text = str(comp["AG SF"])
            row_cells[4].text = str(int(comp["AG Diff"]))
            row_cells[5].text = f"${comp['Total Adj']:,.0f}"
            row_cells[6].text = f"${comp['Adjusted Price']:,.0f}"
            row_cells[7].text = f"${comp['Adj PPSF']:,.2f}"

        avg_price = sum(c["Adjusted Price"] for c in adjusted_rows) / len(adjusted_rows)
        avg_ppsf = sum(c["Adj PPSF"] for c in adjusted_rows) / len(adjusted_rows)

        doc.add_paragraph(f"\nValuation Summary\n"
                          f"Average Adjusted Price: ${avg_price:,.0f}\n"
                          f"Average Price Per SF: ${avg_ppsf:,.2f}")
        if online_estimate_avg:
            doc.add_paragraph(f"Online Estimate Average: ${online_estimate_avg:,.0f}\n"
                              f"Recommended Market Range: "
                              f"${min(online_estimate_avg, avg_price):,.0f} – "
                              f"${max(online_estimate_avg, avg_price):,.0f}")
    else:
        doc.add_paragraph("No valid comps available for valuation.")

    doc.add_paragraph("\nMethodology\n"
                      "Close Price is the original sale price. Adjustments are based on AG SF differences.")

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer