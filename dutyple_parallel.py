from dutyple_core import run_dutyple
import os
import json
import uuid

# 로그 파일 경로
log_path = os.path.join("log.txt")
def write_log(text):
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(text + "\n")

# 입력값 불러오기
with open("input.json", "r", encoding="utf-8") as f:
    config = json.load(f)

input_path = config["input_path"]
nurse_count = config["nurse_count"]
year = config["year"]
month = config["month"]
weekday_D = config["weekday_D"]
weekday_E = config["weekday_E"]
weekday_N = config["weekday_N"]
holiday_D = config["holiday_D"]
holiday_E = config["holiday_E"]
holiday_N = config["holiday_N"]
N_count_nurse = config["N_count_nurse"]

# 결과 저장 폴더
RESULT_FOLDER = "results"
os.makedirs(RESULT_FOLDER, exist_ok=True)

# 반복 시도
max_attempts = 10
success = False
for i in range(1, max_attempts + 1):
    uid = uuid.uuid4().hex[:6]
    output_path = os.path.join(RESULT_FOLDER, f"result_{uid}.xlsx")
    write_log(f"{i}회차 시도 중...")

    try:
        run_dutyple(
            input_path=input_path,
            output_path=output_path,
            nurse_count=nurse_count,
            year=year,
            month=month,
            weekday_D=weekday_D,
            weekday_E=weekday_E,
            weekday_N=weekday_N,
            holiday_D=holiday_D,
            holiday_E=holiday_E,
            holiday_N=holiday_N,
            N_count_nurse=N_count_nurse
        )
        write_log(f"배정 성공! 엑셀 저장 완료 → {output_path}")
        with open("success_uid.txt", "w") as f:
            f.write(uid)
        success = True
        break
    except Exception as e:
        write_log(f"오류 발생: {str(e)}")

if not success:
    write_log("모든 시도 실패. 배정에 실패했습니다.")
