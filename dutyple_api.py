from flask import Flask, request, send_file, jsonify
import os
import uuid
from dutyple_core import run_dutyple

UPLOAD_FOLDER = 'uploads'
RESULT_FOLDER = 'results'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

app = Flask(__name__)

@app.route('/')
def index():
    return 'DUTYPLE API 서버 작동 중'

@app.route('/download-template', methods=['GET'])
def download_template():
    template_path = os.path.join(UPLOAD_FOLDER, 'dutyple.xlsx')
    if not os.path.exists(template_path):
        print("[TEMPLATE] 템플릿 파일 없음")
        return '템플릿 파일이 없습니다', 404
    print("[TEMPLATE] 템플릿 다운로드 요청됨")
    return send_file(template_path, as_attachment=True)

@app.route('/upload', methods=['POST'])
def upload_file():
    print("[UPLOAD] 요청 수신됨")

    if 'file' not in request.files:
        print("[UPLOAD] 파일 없음 ('file' not in request.files)")
        return 'No file part', 400

    file = request.files['file']
    if file.filename == '':
        print("[UPLOAD] 파일 이름 없음")
        return 'No selected file', 400

    uid = uuid.uuid4().hex[:6]
    input_path = os.path.join(UPLOAD_FOLDER, f'{uid}.xlsx')
    output_path = os.path.join(RESULT_FOLDER, f'result_{uid}.xlsx')

    print(f"[UPLOAD] 저장 경로: {input_path}")
    file.save(input_path)
    print(f"[UPLOAD] 파일 저장 완료")

    try:
        print("[UPLOAD] 폼 데이터 파싱 시작")
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
        print("[UPLOAD] run_dutyple 성공")
    except Exception as e:
        print(f"[UPLOAD] run_dutyple 실패: {e}")
        return f"배정 실패: {str(e)}", 500

    return jsonify({"uuid": uid})

@app.route('/generate', methods=['GET'])
def generate_schedule():
    print("[GENERATE] 요청 수신됨")
    try:
        # 가장 최근 업로드된 파일 사용
        files = sorted([f for f in os.listdir(UPLOAD_FOLDER) if f.endswith('.xlsx')])
        if not files:
            print("[GENERATE] 업로드된 파일이 없음")
            return "입력 템플릿 파일이 없습니다", 400

        latest_file = files[-1]
        uid = latest_file.replace('.xlsx', '')
        input_path = os.path.join(UPLOAD_FOLDER, latest_file)
        output_path = os.path.join(RESULT_FOLDER, f"result_{uid}.xlsx")

        print(f"[GENERATE] 입력 파일: {input_path}")
        print(f"[GENERATE] 출력 파일: {output_path}")

        run_dutyple(input_path, output_path,
                    nurse_count=10,
                    year=2025,
                    month=5,
                    weekday_D=2, weekday_E=2, weekday_N=2,
                    holiday_D=1, holiday_E=1, holiday_N=2,
                    N_count_nurse=6)

        with open("success_uid.txt", "w") as f:
            f.write(uid)

        with open("log.txt", "w", encoding="utf-8") as log_file:
            log_file.write("배정 성공\n")

        print("[GENERATE] 배정 성공")
        return jsonify({"status": "success", "uid": uid})

    except Exception as e:
        with open("log.txt", "a", encoding="utf-8") as log_file:
            log_file.write(f"배정 실패: {e}\n")
        print(f"[GENERATE] 실패: {e}")
        return f"배정 실패: {str(e)}", 500

@app.route('/result/<uid>', methods=['GET'])
def get_result(uid):
    path = os.path.join(RESULT_FOLDER, f'result_{uid}.xlsx')
    if not os.path.exists(path):
        print(f"[RESULT] {uid} 파일 없음")
        return '파일 없음', 404
    print(f"[RESULT] {uid} 파일 다운로드됨")
    return send_file(path, as_attachment=True)

@app.route('/result_success', methods=['GET'])
def get_success_result():
    uid_path = "success_uid.txt"
    if not os.path.exists(uid_path):
        print("[SUCCESS] UID 없음")
        return "성공한 결과가 없습니다", 404
    with open(uid_path, "r") as f:
        uid = f.read().strip()
    path = os.path.join("results", f"result_{uid}.xlsx")
    if not os.path.exists(path):
        print(f"[SUCCESS] {uid} 결과 파일 없음")
        return "결과 파일이 없습니다", 404
    print(f"[SUCCESS] {uid} 결과 파일 다운로드됨")
    return send_file(path, as_attachment=True)

@app.route('/log', methods=['GET'])
def get_log():
    if not os.path.exists("log.txt"):
        print("[LOG] log.txt 없음")
        return "", 200
    with open("log.txt", encoding="utf-8") as f:
        return f.read()

if __name__ == '__main__':
    print("[SERVER] DUTYPLE API 시작")
    app.run(debug=True, host='0.0.0.0', port=10000)
