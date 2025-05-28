from flask import Flask, request, send_file, jsonify
from flask_cors import CORS  # ✅ 추가
import os
import uuid
from dutyple_core import run_dutyple
from datetime import datetime

UPLOAD_FOLDER = 'uploads'
RESULT_FOLDER = 'results'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

app = Flask(__name__)
CORS(app)  # ✅ CORS 허용

@app.route('/')
def index():
    return 'DUTYPLE API 서버 작동 중'

@app.route('/download-template', methods=['GET'])
def download_template():
    template_path = os.path.join(UPLOAD_FOLDER, 'dutyple.xlsx')
    if not os.path.exists(template_path):
        return '템플릿 파일이 없습니다', 404
    return send_file(template_path, as_attachment=True)

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part', 400

    file = request.files['file']
    if file.filename == '':
        return 'No selected file', 400

    uid = uuid.uuid4().hex[:6]
    input_path = os.path.join(UPLOAD_FOLDER, f'{uid}.xlsx')
    output_path = os.path.join(RESULT_FOLDER, f'result_{uid}.xlsx')

    file.save(input_path)

    try:
        run_dutyple(input_path, output_path,
                    int(request.form['nurse_count']),
                    int(request.form['year']),
                    int(request.form['month']),
                    int(request.form['weekday_D']),
                    int(request.form['weekday_E']),
                    int(request.form['weekday_N']),
                    int(request.form['holiday_D']),
                    int(request.form['holiday_E']),
                    int(request.form['holiday_N']),
                    int(request.form['N_count_nurse']))
    except Exception as e:
        return f"배정 실패: {str(e)}", 500

    return jsonify({"uuid": uid})

@app.route('/generate', methods=['GET'])
def generate_duty():
    try:
        log_path = "log.txt"
        with open(log_path, "w", encoding="utf-8") as log_file:
            log_file.write("배정 시작...\n")

        uid = uuid.uuid4().hex[:6]
        input_path = os.path.join(UPLOAD_FOLDER, 'dutyple.xlsx')
        output_path = os.path.join(RESULT_FOLDER, f'result_{uid}.xlsx')

        if not os.path.exists(input_path):
            return "입력 템플릿 파일이 없습니다", 400

        run_dutyple(input_path, output_path,
                    nurse_count=10,
                    year=2025,
                    month=5,
                    weekday_D=2, weekday_E=2, weekday_N=2,
                    holiday_D=1, holiday_E=1, holiday_N=2,
                    N_count_nurse=6)

        with open(log_path, "a", encoding="utf-8") as log_file:
            log_file.write("배정 성공\n")

        with open("success_uid.txt", "w") as f:
            f.write(uid)

        return jsonify({"status": "success", "uid": uid})

    except Exception as e:
        with open("log.txt", "a", encoding="utf-8") as log_file:
            log_file.write(f"배정 실패: {e}\n")
        return f"배정 실패: {e}", 500

@app.route('/result/<uid>', methods=['GET'])
def get_result(uid):
    path = os.path.join(RESULT_FOLDER, f'result_{uid}.xlsx')
    if not os.path.exists(path):
        return '파일 없음', 404
    return send_file(path, as_attachment=True)

@app.route('/result_success', methods=['GET'])
def get_success_result():
    uid_path = "success_uid.txt"
    if not os.path.exists(uid_path):
        return "성공한 결과가 없습니다", 404
    with open(uid_path, "r") as f:
        uid = f.read().strip()
    path = os.path.join("results", f"result_{uid}.xlsx")
    if not os.path.exists(path):
        return "결과 파일이 없습니다", 404
    return send_file(path, as_attachment=True)

@app.route('/log', methods=['GET'])
def get_log():
    if not os.path.exists("log.txt"):
        return "", 200
    with open("log.txt", encoding="utf-8") as f:
        return f.read()

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=10000)
