
{
"cells":[
0:{
"cell_type":"code"
"execution_count":26
"id":"70a0102b-8d7d-426e-85c2-8de50b6677fe"
"metadata":{}
"outputs":[
0:{
"name":"stdout"
"output_type":"stream"
"text":[
0:"✅ Styled summary written to: C:\Users\najwa_talentbank\Documents\Campus CF Dahsboard\UPM\test - Clean Tables Styled.xlsx
"
]
}
]
"source":[
0:"import pandas as pd
"
1:"from openpyxl import load_workbook
"
2:"from openpyxl.styles import Font, Border, Side
"
3:"import re
"
4:"
"
5:"# Load Excel file
"
6:"file_path = r"C:\Users\najwa_talentbank\Documents\Campus CF Dahsboard\UPM\test.xlsx"
"
7:"xl = pd.ExcelFile(file_path)
"
8:"
"
9:"# Sheet and column setup
"
10:"main_sheet = 'Event-Checkin'
"
11:"main_columns = ['D', 'F', 'G', 'I', 'J', 'K', 'L', 'M']
"
12:"
"
13:"school_sheet = 'School & Course'
"
14:"school_columns = ['B','C']
"
15:"
"
16:"source_sheet = 'How Did You Find Out About This'
"
17:"source_columns = ['B']
"
18:"
"
19:"# Helper: convert Excel letter to index
"
20:"def col_letter_to_index(letter):
"
21:"    return ord(letter.upper()) - ord('A')
"
22:"
"
23:"# Helper: build clean summary table
"
24:"def clean_summary(df, col_name, table_title):
"
25:"    df_filtered = df[~df[col_name].isin(["", "(Empty)", "(Skip)"])]
"
26:"    counts = df_filtered[col_name].value_counts(dropna=False).reset_index()
"
27:"    counts.columns = ['Category', 'Count']
"
28:"    counts['%'] = (counts['Count'] / counts['Count'].sum() * 100).round(2)
"
29:"
"
30:"    total_row = pd.DataFrame([['Total', counts['Count'].sum(), 100.0]], columns=['Category', 'Count', '%'])
"
31:"    summary = pd.concat([counts, total_row], ignore_index=True)
"
32:"
"
33:"    title_row = pd.DataFrame([[f"Table: {table_title}", "", ""]], columns=['Category', 'Count', '%'])
"
34:"    header_row = pd.DataFrame([['Category', 'Count', '%']], columns=['Category', 'Count', '%'])
"
35:"    blank_row = pd.DataFrame([["", "", ""]], columns=['Category', 'Count', '%'])
"
36:"
"
37:"    return pd.concat([title_row, header_row, summary, blank_row], ignore_index=True)
"
38:"
"
39:"# Collect all tables
"
40:"all_tables = []
"
41:"
"
42:"def process_sheet(sheet, cols, label=""):
"
43:"    df = xl.parse(sheet)
"
44:"    for letter in cols:
"
45:"        idx = col_letter_to_index(letter)
"
46:"        col_name = df.columns[idx]
"
47:"        clean_col_name = re.sub(r"^\d+[:\- ]*\s*", "", str(col_name))  # remove leading numbers, colons, dashes
"
48:"        title = f"{clean_col_name} ({label})" if label else clean_col_name
"
49:"        table_df = clean_summary(df, col_name, title)
"
50:"        all_tables.append(table_df)
"
51:"
"
52:"# Run for all defined sheets
"
53:"process_sheet(main_sheet, main_columns)
"
54:"process_sheet(school_sheet, school_columns, "School & Courses")
"
55:"process_sheet(source_sheet, source_columns, "Event Awareness")
"
56:"
"
57:"# Combine and export to Excel
"
58:"final_df = pd.concat(all_tables, ignore_index=True)
"
59:"output_path = file_path.replace(".xlsx", " - Clean Tables Styled.xlsx")
"
60:"final_df.to_excel(output_path, sheet_name="Summaries", index=False)
"
61:"
"
62:"# === Style with openpyxl ===
"
63:"wb = load_workbook(output_path)
"
64:"ws = wb["Summaries"]
"
65:"
"
66:"bold_font = Font(bold=True)
"
67:"thin_border = Border(
"
68:"    left=Side(style='thin'), right=Side(style='thin'),
"
69:"    top=Side(style='thin'), bottom=Side(style='thin')
"
70:")
"
71:"
"
72:"for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
"
73:"    category_cell = row[0]
"
74:"    if category_cell.value and str(category_cell.value).startswith("Table:"):
"
75:"        # Bold title row
"
76:"        for cell in row:
"
77:"            cell.font = bold_font
"
78:"    elif category_cell.value == "Category":
"
79:"        # Bold header row with border
"
80:"        for cell in row:
"
81:"            cell.font = bold_font
"
82:"            cell.border = thin_border
"
83:"    elif category_cell.value:  # Data or Total rows
"
84:"        for cell in row:
"
85:"            cell.border = thin_border
"
86:"
"
87:"wb.save(output_path)
"
88:"print("✅ Styled summary written to:", output_path)
"
]
}
1:{
"cell_type":"code"
"execution_count":NULL
"id":"7e09a2e5-2d7d-4bbb-9acc-e60d0bbbc79a"
"metadata":{}
"outputs":[]
"source":[]
}
2:{
"cell_type":"code"
"execution_count":NULL
"id":"dfdf3d85-d2a2-4df3-a4bf-dc622ce85fd9"
"metadata":{}
"outputs":[]
"source":[]
}
]
"metadata":{
"kernelspec":{
"display_name":"Python 3 (ipykernel)"
"language":"python"
"name":"python3"
}
"language_info":{
"codemirror_mode":{
"name":"ipython"
"version":3
}
"file_extension":".py"
"mimetype":"text/x-python"
"name":"python"
"nbconvert_exporter":"python"
"pygments_lexer":"ipython3"
"version":"3.13.3"
}
}
"nbformat":4
"nbformat_minor":5
}
