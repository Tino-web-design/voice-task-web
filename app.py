from flask import Flask, render_template, request, send_file, jsonify
import os
from report_generator import generate_reports

app = Flask(__name__)
REPORTS = 'reports'
os.makedirs(REPORTS, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    try:
        data = request.get_json()
        notes = data.get('notes', '').strip()
        name = data.get('customer_name', 'Customer').strip()
        if not notes:
            return jsonify({'error': 'Please enter notes.'}), 400
        result = generate_reports(notes, name, REPORTS)
        return jsonify({'success': True, 'excel': result['excel'], 'pdf': result['pdf'], 'summary': result['summary']})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download/<filename>')
def download(filename):
    filepath = os.path.join(REPORTS, filename)
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True)
    return "File not found", 404

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8080)
