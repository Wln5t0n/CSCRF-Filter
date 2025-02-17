from flask import Flask, render_template, request, send_file
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, Font

pd.options.mode.chained_assignment = None

app = Flask(__name__)

COLUMNS = ['MIIs', 'Qualified Res', 'Mid-size Res', 'Small-size Res', 'Self-certification Res']
APPLICABILITY = ["Alternative Investment Fund (AIF)", "Banker to an Issue and Self-Certified Syndicate Banks (SCSBs)", "Client-based and Proprietary stock brokers", "Collective Investment Scheme (CIS)", "Credit Rating Agency (CRA)", "Custodians", "Debenture Trustee (DT)", "Depository Participants (DPs)", "Designated Depository Participants (DDPs)", "Investment Advisors (IAs)/ Research Analysts (RAs)", "Investment Advisors (IAs) - Non-individual Ias", "Institutional RAs who are registered in other category of REs", "KYC Registration Agencies (KRAs)", "Merchant Bankers (MBs)", "Mutual Funds (MFs)/ Asset Management Companies (AMCs)", "Portfolio Managers", "Registrar to an Issue and Share Transfer Agents (RTA)", "Venture Capital Funds (VCFs)"]
FILE_PATH = "CSCRF_DUMP.xlsx"
SHEET_NAME = "IDR"

df = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME)
df = df.ffill(axis=0)

def main_shit(applicability_filter, company_size):
    filtered_df = df[df['Applicability'] == applicability_filter]
    final_filtered_df = filtered_df[filtered_df[company_size] == 'Yes']
    final_filtered_df.insert(0, 'SR.NO', range(1, len(final_filtered_df) + 1))

    print("\t\t\t\tDEBUG START\n\n")
    print(final_filtered_df['Applicability'])
    print(final_filtered_df[company_size])
    print("\t\t\t\tDEBUG END\n\n")

    final_output_df = final_filtered_df[['SR.NO', 'Standards', 'CSCRF guidelines']]
    final_output_df['compliant status'] = ''
    final_output_df['Auditor remarks'] = ''
    final_output_df['client remarks'] = ''

    output_path = "filtered_data.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"

    for row in dataframe_to_rows(final_output_df, index=False, header=True):
        ws.append(row)

    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if cell.row == 1:
                cell.font = Font(bold=True)
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max_length + 2, 50)

    ws.auto_filter.ref = ws.dimensions
    wb.save(output_path)
    return output_path

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        applicability_filter = request.form['applicability']
        company_size = request.form['company_size']
        output_path = main_shit(applicability_filter, company_size)
        return send_file(output_path, as_attachment=True)

    return render_template('index.html', columns=COLUMNS, applicability=APPLICABILITY)

if __name__ == '__main__':
    app.run(host="0.0.0.0", debug=True)
