from flask import Flask, jsonify, request
import os
from utils.script import load_and_prepare_data, apply_ai_analysis, create_excel_report

app = Flask(__name__)

@app.route('/', methods=['GET'])
def welcome():
    return jsonify({"message": "Welcome to the AI Sales Report API!"})

@app.route('/generate-report', methods=['POST'])
def generate_report():
    try:
        print("Loading data...")
        data = load_and_prepare_data()

        print("Applying AI analysis...")
        enhanced = apply_ai_analysis(data)

        print("Generating Excel report...")
        report_file = create_excel_report(enhanced)

        return jsonify({
            "status": "success",
            "message": "Report generated successfully.",
            "file_path": os.path.abspath(report_file)
        })
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)
