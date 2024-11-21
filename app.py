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

    # Read each file into a DataFrame
    for file in files:
        try:
            df = pd.read_excel(file)
            df.columns = df.columns.str.strip()  # Clean up column names
            dataframes.append(df)
        except Exception as e:
            print(f"Error reading {file.filename}: {e}")
            return f"Error reading {file.filename}. Please ensure it is a valid Excel file.", 400

    if len(dataframes) < 2:
        return "Please upload at least two files.", 400

    # Merge the files based on the common 'Roll number' column
    merged_df = dataframes[0]
    
    for df in dataframes[1:]:
        try:
            # Merge on 'Roll number', and combine horizontally where roll numbers match
            merged_df = pd.merge(merged_df, df, on=['Roll number', 'name'], how='outer')
        except Exception as e:
            print(f"Error merging files: {e}")
            return "Error merging the files.", 500

    # Check if the resulting merged DataFrame is empty
    if merged_df.empty:
        return "The merged file is empty. Please check the input files.", 400

    # Save merged DataFrame to an Excel file and send as response
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
