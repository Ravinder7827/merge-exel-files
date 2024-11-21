from flask import Flask, render_template, request, send_file
import pandas as pd
from io import BytesIO

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/merge', methods=['POST'])
def merge_files():
    files = request.files.getlist('files')
    dataframes = []

    # Iterate through the uploaded files and read them into DataFrames
    for file in files:
        try:
            df = pd.read_excel(file)
            df.columns = df.columns.str.strip()  # Strip spaces in column names
            dataframes.append(df)
        except Exception as e:
            print(f"Error reading {file.filename}: {e}")
            return f"Error reading {file.filename}. Please ensure it is a valid Excel file.", 400

    if not dataframes:
        return "No valid files uploaded.", 400

    # Check if all DataFrames have the same columns
    same_columns = all(df.columns.equals(dataframes[0].columns) for df in dataframes)

    if same_columns:
        # If columns are the same, concatenate vertically
        merged_df = pd.concat(dataframes, ignore_index=True)
    else:
        # If columns differ, concatenate horizontally (aligning on index)
        merged_df = pd.concat(dataframes, axis=1, ignore_index=False)

    if merged_df.empty:
        return "The merged file is empty. Please check the input files.", 400

    output = BytesIO()
    try:
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            merged_df.to_excel(writer, index=False)
        output.seek(0)
    except Exception as e:
        print(f"Error writing Excel file: {e}")
        return "Error generating Excel file.", 500

    return send_file(output, download_name='merged.xlsx', as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
