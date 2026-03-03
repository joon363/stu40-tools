import pandas as pd
from collections import defaultdict
import re

# =========================
# 설정
# =========================
INPUT_FILE = "input.xlsx"
OUTPUT_FILE = "final_assignment.xlsx"

DAYS = ["월", "화", "수", "목", "금"]
SLOTS = [f"slot{i}" for i in range(1, 7)]
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

# =========================
# Phase 1 (각 슬롯 1명)
# =========================
def phase1():
    while True:
        unfilled = [
            s for s in slot_available
            if len(slot_assigned[s]) < CAPACITY_PHASE1
        ]

        if not unfilled:
            break

        def available_count(s):
            return len([
                p for p in slot_available[s]
                if p not in assigned_global
            ])

        unfilled.sort(key=available_count)
        slot = unfilled[0]

        candidates = [
            p for p in slot_available[slot]
            if p not in assigned_global
        ]

        if not candidates:
            slot_assigned[slot].append("(없음1)")
            continue

        candidates.sort(key=lambda p: len(person_available[p]))
        chosen = candidates[0]

        slot_assigned[slot].append(chosen)
        assigned_global.add(chosen)

phase1()

# =========================
# Phase 2.1
# 국서 중복 우선 배제
# =========================
def phase2_1():
    while True:
        unfilled = [
            s for s in slot_available
            if len(slot_assigned[s]) < CAPACITY_TOTAL
        ]

        if not unfilled:
            break

        def available_count(s):
            existing_depts = {extract_dept(p) for p in slot_assigned[s]}
            return len([
                p for p in slot_available[s]
                if p not in assigned_global
                and extract_dept(p) not in existing_depts
            ])

        unfilled.sort(key=available_count)
        slot = unfilled[0]

        existing_depts = {extract_dept(p) for p in slot_assigned[slot]}

        candidates = [
            p for p in slot_available[slot]
            if p not in assigned_global
            and extract_dept(p) not in existing_depts
        ]

        if not candidates:
            break

        candidates.sort(key=lambda p: len(person_available[p]))
        chosen = candidates[0]

        slot_assigned[slot].append(chosen)
        assigned_global.add(chosen)

# =========================
# Phase 2.2
# 국서 중복 허용
# =========================
def phase2_2():
    while True:
        unfilled = [
            s for s in slot_available
            if len(slot_assigned[s]) < CAPACITY_TOTAL
        ]

        if not unfilled:
            break

        def available_count(s):
            return len([
                p for p in slot_available[s]
                if p not in assigned_global
            ])

        unfilled.sort(key=available_count)
        slot = unfilled[0]

        candidates = [
            p for p in slot_available[slot]
            if p not in assigned_global
        ]

        if not candidates:
            break

        candidates.sort(key=lambda p: len(person_available[p]))
        chosen = candidates[0]

        slot_assigned[slot].append(chosen)
        assigned_global.add(chosen)

phase2_1()
phase2_2()

# =========================
# 부족 슬롯 채우기
# =========================
for slot in slot_available:
    while len(slot_assigned[slot]) < CAPACITY_TOTAL:
        slot_assigned[slot].append("(없음2)")

# =========================
# 결과 저장
# =========================
result = pd.DataFrame("", index=SLOTS, columns=DAYS)

for (slot, day), assigned in slot_assigned.items():
    result.at[slot, day] = ", ".join(assigned)

result.to_excel(OUTPUT_FILE)

print("\n슬롯별 최종 인원 수")
for k, v in slot_assigned.items():
    print(k, v)

# =========================
# 사람별 역매핑 (검증용)
# =========================
person_slots = defaultdict(list)

for (slot, day), people in slot_assigned.items():
    for person in people:
        if "없음" not in person:
            person_slots[person].append((slot, day))

print("\n사람별 배정 슬롯")
for person, slots in person_slots.items():
    print(person, "->", slots)

print("\n완료 →", OUTPUT_FILE)