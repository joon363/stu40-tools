import pandas as pd
from collections import defaultdict

# =========================
# 설정
# =========================
INPUT_FILE = "input.xlsx"
OUTPUT_FILE = "final_assignment.xlsx"

DAYS = ["월", "화", "수", "목", "금"]
SLOTS = [f"slot{i}" for i in range(1, 7)]

# =========================
# 데이터 로드
# =========================
df = pd.read_excel(INPUT_FILE)
people = df["이름"].tolist()

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
phase1_assigned_people = set()

# =========================
# Phase 1
# =========================
def phase1():
    used_this_phase = set()

    while True:
        progress = False

        unfilled = [
            s for s in slot_available
            if len(slot_assigned[s]) < 1
        ]

        if not unfilled:
            break

        # 가장 후보 적은 슬롯 선택
        unfilled.sort(key=lambda s: len([
            p for p in slot_available[s]
            if p not in used_this_phase
        ]))

        for slot in unfilled:
            candidates = [
                p for p in slot_available[slot]
                if p not in used_this_phase
            ]

            if not candidates:
                if len(slot_assigned[slot]) == 0:
                    slot_assigned[slot].append("(없음1)")
                    progress = True
                continue

            candidates.sort(key=lambda p: len(person_available[p]))
            chosen = candidates[0]

            slot_assigned[slot].append(chosen)
            used_this_phase.add(chosen)
            phase1_assigned_people.add(chosen)
            progress = True
            break  # 한 명 배정 후 다시 탐색

        if not progress:
            break

phase1()

# =========================
# Phase1 저장
# =========================
phase1_result = pd.DataFrame("(없음1)", index=SLOTS, columns=DAYS)

for (slot, day), assigned in slot_assigned.items():
    if assigned:
        phase1_result.at[slot, day] = assigned[0]

phase1_result.to_excel("phase1_assignment.xlsx")
print("Phase1 저장 완료")

# =========================
# Phase 2
# =========================
def phase2(exclude_people):
    used_this_phase = set()

    while True:
        progress = False

        unfilled = [
            s for s in slot_available
            if len(slot_assigned[s]) < 2
        ]

        if not unfilled:
            break

        # 후보 적은 슬롯 우선
        unfilled.sort(key=lambda s: len([
            p for p in slot_available[s]
            if p not in exclude_people
            and p not in used_this_phase
            and p not in slot_assigned[s]
        ]))

        for slot in unfilled:
            candidates = [
                p for p in slot_available[slot]
                if p not in exclude_people
                and p not in used_this_phase
                and p not in slot_assigned[slot]
            ]

            if not candidates:
                continue

            candidates.sort(key=lambda p: len(person_available[p]))
            chosen = candidates[0]

            slot_assigned[slot].append(chosen)
            used_this_phase.add(chosen)
            progress = True
            break  # 한 명 배정 후 다시 탐색

        if not progress:
            break

# Phase2-1 (Phase1 인원 제외)
phase2(exclude_people=phase1_assigned_people)

# Phase2-2 (전원 포함)
phase2(exclude_people=set())

# =========================
# 2명 안된 슬롯 채우기
# =========================
for slot in slot_available:
    while len(slot_assigned[slot]) < 2:
        slot_assigned[slot].append("(없음2)")

# =========================
# 최종 결과 저장
# =========================
result = pd.DataFrame("", index=SLOTS, columns=DAYS)

for (slot, day), assigned in slot_assigned.items():
    result.at[slot, day] = ", ".join(assigned)

result.to_excel(OUTPUT_FILE)

# =========================
# 사람별 최종 통합 출력
# =========================
person_result = defaultdict(list)

for (slot, day), assigned_list in slot_assigned.items():
    for person in assigned_list:
        if person.startswith("(없음"):
            continue
        person_result[person].append((day, slot))

print("\n===== 사람별 최종 배정 결과 =====")

for person in sorted(people):
    assigned = person_result.get(person, [])

    if not assigned:
        print(f"{person}: 배정없음")
        continue

    assigned.sort(key=lambda x: (DAYS.index(x[0]), x[1]))
    formatted = [f"{day} {slot}" for day, slot in assigned]
    print(f"{person}: {', '.join(formatted)}")

print("\n완료 →", OUTPUT_FILE)