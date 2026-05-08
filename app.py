from flask import Flask, request, send_file, jsonify
import subprocess
import json
import os
import tempfile
import uuid

app = Flask(__name__)

@app.route('/health', methods=['GET'])
def health():
    return jsonify({"status": "ok"})

@app.route('/generate', methods=['POST'])
def generate():
    try:
        data = request.get_json()
        if not data:
            return jsonify({"error": "No JSON data provided"}), 400

        # Write data to temp file
        tmp_dir = tempfile.mkdtemp()
        data_file = os.path.join(tmp_dir, 'data.json')
        output_file = os.path.join(tmp_dir, f'report_{uuid.uuid4().hex}.pptx')

        with open(data_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False)

        # Run Node.js script
        script_path = os.path.join(os.path.dirname(__file__), 'generate.js')
        result = subprocess.run(
            ['node', script_path, data_file, output_file],
            capture_output=True, text=True, timeout=60
        )

        if result.returncode != 0:
            return jsonify({
                "error": "PowerPoint generation failed",
                "details": result.stderr
            }), 500

        if not os.path.exists(output_file):
            return jsonify({"error": "Output file not created"}), 500

        return send_file(
            output_file,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            as_attachment=True,
            download_name='Kampagnenanalyse.pptx'
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
