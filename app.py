from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from docx import Document
from datetime import datetime
import os
from openai import OpenAI

app = Flask(__name__)
CORS(app, origins=["https://financial-optimization-dashboard.netlify.app"])

REPORT_FOLDER = os.path.join(app.root_path, 'static', 'reports')
os.makedirs(REPORT_FOLDER, exist_ok=True)

client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

def generate_section(prompt):
    try:
        response = client.chat.completions.create(
            model="gpt-4",
            messages=[
                {"role": "system", "content": "You are a business financial advisor generating financial analysis reports."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"Error generating this section: {e}"

@app.route('/')
def home():
    return "Financial Optimization Backend is Running!"

@app.route('/generate', methods=['POST'])
def generate_report():
    files = [request.files.get('file1'), request.files.get('file2'), request.files.get('file3')]
    context = ""

    for file in files:
        if file:
            try:
                content = file.read().decode("utf-8", errors="ignore")
            except:
                content = "Unable to read file content."
            context += f"\n--- {file.filename} ---\n{content}\n"

    if not context.strip():
        return jsonify({"error": "No valid file content found."}), 400

    doc = Document()
    doc.add_heading("Financial Metric Optimization Report", 0)

    sections = [
        ("Executive Summary", "Summarize key findings from financial documents and high-level trends."),
        ("1. Revenue & Income Patterns", "Analyze trends in sales, income sources, and pricing structures."),
        ("2. Expense Breakdown & Cost Management", "Assess fixed vs. variable costs and highlight opportunities to reduce expenses."),
        ("3. Profitability Analysis", "Evaluate gross margin, net margin, and profitability over time."),
        ("4. Cash Flow Insights", "Analyze inflows and outflows of cash to assess liquidity and runway."),
        ("5. Forecasting & Financial Risk", "Identify risk areas and provide projection insights."),
        ("6. Recommendations & Financial Optimization", "Offer specific strategies to improve financial health. Recommendations are based on data and not intended as legal or tax advice."),
        ("Conclusion", "Wrap up the financial overview and suggest next steps for improvement.")
    ]

    for title, instruction in sections:
        doc.add_heading(title, level=1)
        prompt = f"{instruction}\n\nBusiness Context:\n{context}"
        content = generate_section(prompt)
        doc.add_paragraph(content)

    filename = f"financial_report_{datetime.now().strftime('%Y%m%d%H%M%S')}.docx"
    file_path = os.path.join(REPORT_FOLDER, filename)
    doc.save(file_path)

    return jsonify({'download_url': f'/static/reports/{filename}'})

@app.route('/static/reports/<path:filename>')
def download_file(filename):
    return send_from_directory(REPORT_FOLDER, filename, as_attachment=True)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
