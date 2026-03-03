import pandas as pd
from collections import defaultdict
import random

# =========================
# 설정
# =========================
INPUT_FILE = "input.xlsx"
MIN_PER_PERSON = 2
MAX_PER_SLOT = 2
MIN_PER_SLOT = 1

DAYS = ["월", "화", "수", "목", "금"]
SLOT_COLUMNS = [f"slot{i}" for i in range(1, 6)]

# =========================
# 1. 데이터 로드
# =========================
df = pd.read_excel(INPUT_FILE)

if "이름" not in df.columns:
    raise Exception("A열에 '이름' 컬럼이 필요합니다.")

for col in SLOT_COLUMNS:
    if col not in df.columns:
        raise Exception(f"{col} 컬럼이 없습니다.")

people = df["이름"].tolist()

person_available = defaultdict(set)
slot_available = defaultdict(set)

# =========================
# 2. 가능 슬롯 구성
# =========================
for _, row in df.iterrows():
    name = row["이름"]

    for slot_col in SLOT_COLUMNS:
        cell = row[slot_col]

        if pd.isna(cell):
            continue

        days = [d.strip() for d in str(cell).split(",")]

        for day in days:
            if day in DAYS:
                slot_key = f"{day}_{slot_col}"
                person_available[name].add(slot_key)
                slot_available[slot_key].add(name)

# =========================
# 3. 초기화
# =========================
person_assigned = defaultdict(set)
slot_assigned = defaultdict(set)

# =========================
# 4. 단계1: 모든 슬롯 최소 1명 채우기
# =========================
for slot in slot_available:
    candidates = [
        p for p in slot_available[slot]
        if len(person_assigned[p]) < MIN_PER_PERSON
    ]

    if not candidates:
        raise Exception(f"{slot} 슬롯 최소 인원 충족 불가")

    # 배정 적은 사람 우선
    candidates.sort(key=lambda x: len(person_assigned[x]))

    chosen = candidates[0]
    person_assigned[chosen].add(slot)
    slot_assigned[slot].add(chosen)

# =========================
# 5. 단계2: 사람 최소 2슬롯 맞추기
# =========================
sorted_people = sorted(people, key=lambda x: len(person_available[x]))

for person in sorted_people:
    print(f"I am{person}")
    while len(person_assigned[person]) < MIN_PER_PERSON:

        possible_slots = [
            s for s in person_available[person]
            if len(slot_assigned[s]) < MAX_PER_SLOT
        ]
        print(f"my slots: {possible_slots}")

        if not possible_slots:
            print(f"{person} 추가 배정 불가")
            person_assigned[person].add("slot0")
            break

        # 인원 적은 슬롯 우선
        possible_slots.sort(key=lambda s: len(slot_assigned[s]))

        chosen_slot = possible_slots[0]

        person_assigned[person].add(chosen_slot)
        slot_assigned[chosen_slot].add(person)

# =========================
# 6. 검증
# =========================
for slot, assigned in slot_assigned.items():
    if len(assigned) < MIN_PER_SLOT:
        raise Exception(f"{slot} 최소 인원 미달")

    if len(assigned) > MAX_PER_SLOT:
        raise Exception(f"{slot} 최대 인원 초과")

# =========================
# 7. 결과 저장
# =========================
result = []

for person in people:
    for slot in person_assigned[person]:
        result.append([person, slot])

result_df = pd.DataFrame(result, columns=["이름", "배정슬롯"])
result_df.to_excel("result.xlsx", index=False)

print("배정 완료 → result.xlsx 생성")