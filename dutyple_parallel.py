from dutyple_core import run_dutyple
import os

# 고정된 입력 파일 (바탕화면 기준)
desktop = os.path.join(os.path.expanduser("~"), "Desktop")
input_path = os.path.join(desktop, "dutyple.xlsx")

# results 폴더 내에 result 번호 붙이기
def get_next_output_path():
    result_dir = "results"
    os.makedirs(result_dir, exist_ok=True)
    existing = [f for f in os.listdir(result_dir) if f.startswith("★dutyple_example_") and f.endswith(".xlsx")]
    numbers = [int(f.split("_")[2].split(".")[0]) for f in existing if f.split("_")[2].split(".")[0].isdigit()]
    next_num = max(numbers) + 1 if numbers else 1
    return os.path.join(result_dir, f"★dutyple_example_{next_num}★.xlsx")

output_path = get_next_output_path()

# 기본 테스트 파라미터 (수정 가능)
run_dutyple(
    input_path=input_path,
    output_path=output_path,
    nurse_count=10,
    year=2025,
    month=5,
    weekday_D=2, weekday_E=3, weekday_N=2,
    holiday_D=1, holiday_E=1, holiday_N=2,
    N_count_nurse=7
)

print(f"✅ 저장 완료: {output_path}")