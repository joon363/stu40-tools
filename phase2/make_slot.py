import pandas as pd
from collections import defaultdict
import re

# =========================
# 설정
# =========================
INPUT_FILE = "input_final.xlsx"

DAYS = ["월", "화", "수", "목", "금"]
SLOTS = [
    "11:00-12:20",
    "12:25-13:50",
    "13:55-15:20",
    "15:25-16:50",
    "16:55~18:20",
    "18:25~19:50",
]
CAPACITY_PHASE1 = 1
CAPACITY_TOTAL = 2

# =========================
# 유틸
# =========================
def extract_dept(name):
    match = re.search(r"\((.*?)\)", str(name))
    return match.group(1) if match else None

# =========================
# 데이터 로드
# =========================
df = pd.read_excel(INPUT_FILE)

person_available = defaultdict(set)
slot_available = defaultdict(set)

for _, row in df.iterrows():
    name = row["이름"]

    for slot in SLOTS:
        cell = row.get(slot)
        if pd.isna(cell):
            continue

        days = [d.strip() for d in str(cell).split(",")]
        for day in days:
            if day in DAYS:
                key = (slot, day)
                person_available[name].add(key)
                slot_available[key].add(name)

# =========================
# 배정 저장
# =========================
slot_assigned = defaultdict(list)
assigned_global = set()  # ✅ 1인 1타임 강제용
all_available_df = pd.DataFrame("", index=SLOTS, columns=DAYS)

for slot in SLOTS:
    for day in DAYS:
        key = (slot, day)

        people_list = sorted(slot_available.get(key, []))

        if people_list:
            all_available_df.at[slot, day] = "\n".join(people_list)

# slot 이름을 실제 시간으로 변경
all_available_df.index = SLOTS

all_available_df.to_excel("all_available.xlsx")

print("가능 인원 전체 저장 완료 → all_available.xlsx")