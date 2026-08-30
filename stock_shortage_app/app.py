import io
import pandas as pd
from flask import Flask, render_template, request, send_file

app = Flask(__name__)

def normalize_cols(df):
    df.columns = [str(col).strip() for col in df.columns]
    # Standardize common column name variations
    col_map = {}
    for col in df.columns:
        c_lower = col.lower().replace(' ', '').replace('_', '')
        if c_lower in ['article', 'material', 'materialnumber', 'item']:
            col_map[col] = 'Article'
        elif c_lower in ['unrestricted', 'unrestrictedstock', 'unrestricteduse', 'labst']:
            col_map[col] = 'Unrestricted'
        elif c_lower in ['salesdocument', 'salesdoc', 'so', 'salesorder', 'vbeln']:
            col_map[col] = 'Sales Document'
        elif c_lower in ['requirementquantity(einheit)', 'requirementquantity', 'reqqty', 'requirementqty', 'quantity', 'bdmng']:
            col_map[col] = 'Requirement quantity (EINHEIT)'
        elif c_lower in ['materialdescription', 'description', 'articledescription', 'maktx']:
            col_map[col] = 'Material Description'
        elif c_lower in ['phantomitem', 'phantom', 'phantomflag', 'phantomit']:
            col_map[col] = 'Phantom item'
    df.rename(columns=col_map, inplace=True)
    return df

def generate_shortage_report(file_mb52, cohv_files):
    # 1. Read the MB52 Stock file
    df_mb52 = pd.read_excel(file_mb52)
    df_mb52 = normalize_cols(df_mb52)
    
    if 'Article' not in df_mb52.columns or 'Unrestricted' not in df_mb52.columns:
        raise ValueError(f"MB52 file missing required columns. Found: {list(df_mb52.columns)}")

    if 'Phantom item' in df_mb52.columns:
        df_mb52 = df_mb52[df_mb52['Phantom item'].astype(str).str.strip().str.upper() != 'X'].copy()

    if 'Material Description' in df_mb52.columns:
        df_mb52 = df_mb52[~df_mb52['Material Description'].astype(str).str.contains('gls', case=False, na=False)].copy()

    df_mb52['Article'] = df_mb52['Article'].astype(str).str.strip()

    # 2. Dynamically read and combine all uploaded COHV files
    df_cohv_list = []
    for f in cohv_files:
        if f.filename: # Ensure the file is not empty
            df_temp = pd.read_excel(f)
            df_temp = normalize_cols(df_temp)
            df_cohv_list.append(df_temp)
            
    if not df_cohv_list:
        raise ValueError("No valid COHV data found.")

    df_cohv = pd.concat(df_cohv_list, ignore_index=True)
    
    req_cohv_cols = ['Sales Document', 'Article', 'Requirement quantity (EINHEIT)']
    missing = [c for c in req_cohv_cols if c not in df_cohv.columns]
    if missing:
        raise ValueError(f"COHV file(s) missing required columns: {missing}. Found columns: {list(df_cohv.columns)}")

    if 'Phantom item' in df_cohv.columns:
        df_cohv = df_cohv[df_cohv['Phantom item'].astype(str).str.strip().str.upper() != 'X'].copy()

    if 'Material Description' not in df_cohv.columns:
        df_cohv['Material Description'] = df_cohv['Article']

    if 'Material Description' in df_cohv.columns:
        df_cohv = df_cohv[~df_cohv['Material Description'].astype(str).str.contains('gls', case=False, na=False)].copy()

    df_cohv = df_cohv.dropna(subset=['Sales Document', 'Article'])
    df_cohv['Article'] = df_cohv['Article'].astype(str).str.strip()
    
    # 3. MB52 Aggregation (SUMIF equivalent for Unrestricted Stock)
    stock_summary = df_mb52.groupby('Article')['Unrestricted'].sum().reset_index()
    stock_summary.rename(columns={'Unrestricted': 'Total_Unrestricted_Stock'}, inplace=True)

    # 4. Article-Level Aggregation (Total Demand vs Stock)
    article_summary = df_cohv.groupby(['Article', 'Material Description'], as_index=False)[
        'Requirement quantity (EINHEIT)'
    ].sum()
    article_summary.rename(columns={'Requirement quantity (EINHEIT)': 'Total_Required_Qty'}, inplace=True)

    # Merge stock into article summary
    article_summary = pd.merge(article_summary, stock_summary, on='Article', how='left')
    article_summary['Total_Unrestricted_Stock'] = article_summary['Total_Unrestricted_Stock'].fillna(0)
    article_summary['Shortage_Qty'] = (
        article_summary['Total_Required_Qty'] - article_summary['Total_Unrestricted_Stock']
    ).apply(lambda x: max(0, x))
    article_summary['Status'] = article_summary['Shortage_Qty'].apply(
        lambda x: 'SHORTAGE' if x > 0 else 'AVAILABLE'
    )

    # 5. Line-Item Level Details
    df_lines = pd.merge(df_cohv, stock_summary, on='Article', how='left')
    df_lines['Total_Unrestricted_Stock'] = df_lines['Total_Unrestricted_Stock'].fillna(0)
    df_lines['Line_Shortage_Qty'] = (
        df_lines['Requirement quantity (EINHEIT)'] - df_lines['Total_Unrestricted_Stock']
    ).apply(lambda x: max(0, x))
    df_lines['Line_Status'] = df_lines['Line_Shortage_Qty'].apply(
        lambda x: 'INSUFFICIENT PLANT STOCK' if x > 0 else 'STOCK COVERS LINE'
    )

    # 6. Sales Order Summary
    so_summary = df_lines.groupby('Sales Document').agg(
        Total_Line_Items=('Article', 'count'),
        Distinct_Articles=('Article', 'nunique'),
        Total_Requirement=('Requirement quantity (EINHEIT)', 'sum'),
        Lines_With_Shortage=('Line_Shortage_Qty', lambda x: (x > 0).sum())
    ).reset_index()

    so_summary['Order_Status'] = so_summary['Lines_With_Shortage'].apply(
        lambda x: 'HAS SHORTAGE' if x > 0 else 'AVAILABLE'
    )

    # 7. Write all summaries to a single Excel workbook in memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        so_summary.sort_values(by='Lines_With_Shortage', ascending=False).to_excel(
            writer, sheet_name='Sales_Order_Summary', index=False
        )
        article_summary.sort_values(by='Shortage_Qty', ascending=False).to_excel(
            writer, sheet_name='Article_Shortage_Summary', index=False
        )
        df_lines.to_excel(
            writer, sheet_name='Line_Item_Details', index=False
        )

    output.seek(0)
    return output


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file_mb52 = request.files.get('mb52_file')
        cohv_files = request.files.getlist('cohv_files')

        if not file_mb52 or not cohv_files or cohv_files[0].filename == '':
            return "Please upload one MB52 file and at least one COHV file.", 400

        try:
            # Process and generate the download file
            processed_file = generate_shortage_report(file_mb52, cohv_files)
            
            return send_file(
                processed_file,
                as_attachment=True,
                download_name="Consolidated_Shortage_Report.xlsx",
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            import traceback
            print("ERROR processing files:", str(e))
            traceback.print_exc()
            return f"An error occurred while processing the files: {str(e)}", 500

    return render_template('index.html')


if __name__ == '__main__':
    # Using threaded=True to handle multiple connections/uploads smoothly
    app.run(debug=True, port=5000, threaded=True)
