from flask import Flask, request, send_file, jsonify, render_template, after_this_request
from fpdf import FPDF
import pandas as pd
import tempfile
import os
import time

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

class PDF(FPDF):
    def header(self):
        pass

def generate_summary_block(pdf, title_lines, summary_lines, table_headers, table_data):
    pdf.set_font("Arial", "B", 14)
    for line in title_lines:
        pdf.cell(0, 10, line, ln=True, align="C")
    pdf.ln(2)

    pdf.set_font("Arial", "B", 12)
    for line in summary_lines:
        pdf.cell(0, 10, line, ln=True, align="C")
    pdf.ln(2)

    pdf.set_fill_color(255, 153, 51)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", "B", 12)
    col_widths = [90, 90]
    for i in range(len(table_headers)):
        pdf.cell(col_widths[i], 10, table_headers[i], 1, 0, 'C', 1)
    pdf.ln()

    pdf.set_fill_color(255, 255, 255)
    pdf.set_text_color(0, 0, 153)
    pdf.set_font("Arial", "", 12)
    for row in table_data:
        for i in range(len(row)):
            pdf.cell(col_widths[i], 10, str(row[i]), 1, 0, 'C', 1)
        pdf.ln()
    pdf.ln(10)

def get_fuzzy_column(df, keyword_parts):
    for col in df.columns:
        col_lower = col.lower()
        if all(k.lower() in col_lower for k in keyword_parts):
            return col
    return None

def draw_custom_ageing_chart(pdf, df):
    ageing_col = get_fuzzy_column(df, ["call", "ageing"])
    if not ageing_col:
        return

    df[ageing_col] = pd.to_numeric(df[ageing_col], errors='coerce').fillna(0)

    bins = [-1, 0, 1, 2, float('inf')]
    labels = ["d0", "d1", "d2", "D2+"]
    df["ageing_bin"] = pd.cut(df[ageing_col], bins=bins, labels=labels)
    grouped = df["ageing_bin"].value_counts().sort_index()

    max_count = grouped.max()
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "Call Close Ageing Distribution", ln=True, align="C")
    pdf.ln(5)

    chart_width = 120
    chart_height = 60
    bar_width = 20
    bar_gap = 10
    total_width = len(grouped) * (bar_width + bar_gap) - bar_gap
    chart_x = (210 - total_width) / 2
    chart_y = pdf.get_y()

    pdf.set_draw_color(0, 0, 0)
    pdf.rect(chart_x - 5, chart_y, total_width + 10, chart_height)

    bar_height_factor = chart_height / max_count if max_count else 1

    for idx, (label, count) in enumerate(grouped.items()):
        x = chart_x + idx * (bar_width + bar_gap)
        height = count * bar_height_factor
        y = chart_y + chart_height - height

        pdf.set_fill_color(100, 149, 237)
        pdf.rect(x, y, bar_width, height, 'FD')

        pdf.set_xy(x, y - 8)
        pdf.set_font("Arial", "B", 10)
        pdf.cell(bar_width, 8, str(count), 0, 0, 'C')

        pdf.set_xy(x, chart_y + chart_height + 2)
        pdf.set_font("Arial", "", 10)
        pdf.cell(bar_width, 6, label, 0, 0, 'C')

    pdf.ln(chart_height + 15)

@app.route('/convert', methods=['POST'])
def convert_excel_to_pdf():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400

    excel_temp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    file.save(excel_temp.name)

    with pd.ExcelFile(excel_temp.name) as xls:
        pdf = PDF()

        if "Account_Receivable" in xls.sheet_names:
            df = pd.read_excel(xls, "Account_Receivable")
            df.columns = df.columns.str.strip()
            bucket_col = get_fuzzy_column(df, ['bucket'])
            amount_col = get_fuzzy_column(df, ['total', 'amount'])
            sf_name = df.get("SF_Name", ["GOLDEN SERVICE ELECTRONICS SALES AND SERVICE"])[0]
            sf_code = df.get("SF_Code", ["BASCGOLDEN"])[0]
            if bucket_col and amount_col:
                df = df[df[bucket_col].notna() & df[amount_col].notna()]
                df_grouped = df.groupby(bucket_col)[amount_col].sum().round(0).reset_index()
                df_grouped = df_grouped[~df_grouped[bucket_col].str.lower().str.contains("sale of defectives")]
                total = int(df_grouped[amount_col].sum())
                table = [[b, f"{int(a):,}"] for b, a in zip(df_grouped[bucket_col], df_grouped[amount_col])]
                pdf.add_page()
                generate_summary_block(
                    pdf,
                    [sf_name, f"SF_Code:({sf_code})", "Account_Receivable_Summary", "Category - All"],
                    [f"Total Receivable: Rs. {total:,}"],
                    ["Source", "Receivable_Value(INR)"],
                    table
                )
                draw_custom_ageing_chart(pdf, df)

        def process_payable_sheet(sheet, category):
            if sheet not in xls.sheet_names:
                return

            df = pd.read_excel(xls, sheet)
            df.columns = df.columns.str.strip()
            if df.empty:
                return

            sf_name = df.get("SF_Name", ["GOLDEN SERVICE-BANGALORE"])[0]
            sf_code = df.get("SF_Code", ["BASCGOLDEN"])[0]
            total_calls = len(df)

            pdf.add_page()

            install_col = get_fuzzy_column(df, ["installation"])
            df[install_col] = pd.to_numeric(df[install_col], errors='coerce').fillna(0)
            installation_sum = round(df[install_col].sum())
            gst = round(installation_sum * 0.18)
            table = [
                ["Earnings_Based_on_Rate_Card", f"{installation_sum:,}"],
                ["GST", f"{gst:,}"]
            ]

            if sheet == "Large_Account_Payable" and "OCRM - Cost" in xls.sheet_names:
                df_ocrm = pd.read_excel(xls, "OCRM - Cost")
                df_ocrm.columns = df_ocrm.columns.str.strip()
                cost_fields = [
                    ("EWC Ince", ["ewc"]),
                    ("OCRM Cost", ["ocrm", "cost"]),
                    ("SRMS Cancel - Instal cost", ["cancel"]),
                    ("OCRM Transp Conv", ["transp"])
                ]
                for label, keys in cost_fields:
                    col = get_fuzzy_column(df_ocrm, keys)
                    if col:
                        val = round(pd.to_numeric(df_ocrm[col], errors='coerce').fillna(0).sum())
                        table.append([label, f"{val:,}"])

            total_earnings = sum([int(t[1].replace(',', '')) for t in table])

            generate_summary_block(
                pdf,
                [sf_name, f"SF_Code:({sf_code})", "Account_Payable_Summary", f"Category - {category}"],
                [f"Total Closed Calls: {total_calls}", f"Total Earnings Inc.GST:Rs. {total_earnings:,}"],
                ["Source", "Earning_Value(INR)"],
                table
            )

            draw_custom_ageing_chart(pdf, df)

        process_payable_sheet("Furniture_Account_Payable", "Furniture")
        process_payable_sheet("Large_Account_Payable", "Large")
        process_payable_sheet("Mobile_Account_Payable", "Mobile")

    output_temp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    output_path = output_temp.name
    pdf.output(output_path)

    @after_this_request
    def cleanup_temp_file(response):
        try:
            os.remove(excel_temp.name)
            time.sleep(1)
        except Exception as e:
            print("Warning: temp file not deleted:", e)
        return response

    return send_file(output_path, as_attachment=True, download_name="Service_Center_BASCGOLDEN_Summary_Merged.pdf")

if __name__ == '__main__':
    app.run(debug=True)
