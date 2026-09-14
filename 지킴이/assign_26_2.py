"""
26-2학기 지킴이 시간표 자동 배정 프로그램
1학기 phase 2 배정 결과를 기반으로, 동일 인물에 대해 거리가 먼 슬롯(요일/시간대)을 우선순위에 두어 배정합니다.

[배정 원칙]
1. [Hard] 1인 1타임 배정 (모든 인원 1타임씩 배정)
2. [Hard] 슬롯당 정원 2명 (부족 인원은 공석으로 표시)
3. [Hard] 동일 슬롯 내 같은 국 인원 중복 금지
4. [Hard] 국원 본인이 지원한 가능한 슬롯만 배정
5. [Soft] 1학기 phase 2 배정 위치(요일/교시)와 거리가 먼 슬롯 우선 배정
6. [Soft] 1학기 당시 같은 슬롯이었던 파트너와의 재매칭 회피
7. [Soft] 셔플링을 통한 공정한 탐색 및 최고 거리 점수 조합 자동 채택
"""

import os
import sys
import re
import random
from pathlib import Path
from collections import defaultdict
import pandas as pd

# Windows 콘솔 인코딩 대응
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

# =========================
# 설정 및 경로
# =========================
BASE_DIR = Path(__file__).parent
INPUT_FILE = BASE_DIR / "26-2학기 지킴이 시간 조사(응답).xlsx"
OLD_ASSIGNMENT_FILE = BASE_DIR / "1학기" / "final_assignment.xlsx"
OUTPUT_DIR = BASE_DIR / "output_26_2"
OUTPUT_DIR.mkdir(exist_ok=True)

FINAL_ASSIGNMENT_FILE = OUTPUT_DIR / "final_assignment.xlsx"
ALL_AVAILABLE_FILE = OUTPUT_DIR / "all_available.xlsx"
PERSON_ASSIGNMENT_FILE = OUTPUT_DIR / "person_assignment.xlsx"

DAYS = ["월", "화", "수", "목", "금"]
SLOTS = [
    "11:00-12:20",
    "12:25-1:50",
    "1:55-3:20",
    "3:25-4:50",
    "4:55-6:20",
    "6:25-7:50",
]

CAPACITY_PER_SLOT = 2
NUM_OPTIMIZE_TRIALS = 150  # 최적 배정안 탐색 시도 횟수


def load_1st_semester_data(old_file_path: Path):
    """1학기 phase 2 최종 배정표를 불러와 국원별 요일/슬롯 및 파트너 정보를 추출합니다."""
    old_slot_map = {}
    old_partner_map = defaultdict(set)

    if not old_file_path.exists():
        print(f"⚠️ [참고] 1학기 배정표 파일({old_file_path.name})을 찾을 수 없습니다. 기본 배정으로 진행합니다.")
        return old_slot_map, old_partner_map

    print(f"[0] 1학기 배정 데이터 로드: {old_file_path.name}")
    df_old = pd.read_excel(old_file_path, index_col=0)

    for s_idx, slot in enumerate(df_old.index):
        for d_idx, day in enumerate(df_old.columns):
            val = df_old.at[slot, day]
            if pd.isna(val):
                continue

            members = [m.strip() for m in str(val).split(",") if m.strip() and "없음" not in m]
            names = [m.split("(")[0].strip() for m in members]

            for i, n1 in enumerate(names):
                old_slot_map[n1] = (day, d_idx, slot, s_idx)
                for j, n2 in enumerate(names):
                    if i != j:
                        old_partner_map[n1].add(n2)

    print(f" - 1학기 배정자 {len(old_slot_map)}명의 이전 요일/시간대 정보 로드 완료\n")
    return old_slot_map, old_partner_map


def load_and_preprocess_data(file_path: Path):
    """2학기 엑셀 데이터를 로드하고 중복 응답을 처리하여 최신 응답을 추출합니다."""
    print(f"[1] 2학기 데이터 로드: {file_path.name}")
    raw_df = pd.read_excel(file_path)
    print(f" - 원본 응답 수: {len(raw_df)}건")

    # 이름 정제: 공백 제거 및 '(최종)' 등 수식어 제거
    raw_df["이름_정제"] = raw_df["이름"].astype(str).str.strip().apply(lambda x: re.sub(r"\(최종\)", "", x).strip())

    # 중복 제출자 확인
    dup_counts = raw_df["이름_정제"].value_counts()
    dup_names = dup_counts[dup_counts > 1].index.tolist()
    if dup_names:
        print(f" - 중복 응답 감지 ({len(dup_names)}명): {', '.join(dup_names)}")
        for name in dup_names:
            sub = raw_df[raw_df["이름_정제"] == name]
            print(f"   * {name}: 총 {len(sub)}회 제출 -> 마지막(최신) 응답 유지")

    # 타임스탬프 기준 정렬 후 최신 응답 유지
    df = raw_df.sort_values(by="타임스탬프").drop_duplicates(subset=["이름_정제"], keep="last").reset_index(drop=True)
    print(f" - 중복 제거 후 최종 인원: {len(df)}명\n")

    # 요일 컬럼 추출
    day_cols = {}
    for col in df.columns:
        match = re.search(r"\[([월화수목금])\]", str(col))
        if match:
            day_cols[match.group(1)] = col

    missing_days = [d for d in DAYS if d not in day_cols]
    if missing_days:
        raise ValueError(f"요일 컬럼을 찾을 수 없습니다: {missing_days}")

    person_dept = {}
    person_available = defaultdict(set)
    slot_available = defaultdict(set)
    people = []

    for _, row in df.iterrows():
        name_only = row["이름_정제"]
        dept = str(row["국"]).strip() if pd.notna(row["국"]) else "미지정"
        display_name = f"{name_only}({dept})"

        people.append(display_name)
        person_dept[display_name] = dept

        for day in DAYS:
            col = day_cols[day]
            cell_val = row[col]
            if pd.isna(cell_val):
                continue

            time_items = [t.strip() for t in str(cell_val).split(",")]
            for t in time_items:
                if t in SLOTS:
                    slot_key = (t, day)
                    person_available[display_name].add(slot_key)
                    slot_available[slot_key].add(display_name)

    return people, person_dept, person_available, slot_available


def export_all_available(slot_available: dict):
    """각 슬롯별 가능한 모든 인원 목록을 엑셀로 저장합니다."""
    grid = pd.DataFrame("", index=SLOTS, columns=DAYS)
    for slot in SLOTS:
        for day in DAYS:
            candidates = sorted(slot_available.get((slot, day), []))
            grid.at[slot, day] = "\n".join(candidates) if candidates else "(가능자 없음)"

    grid.to_excel(ALL_AVAILABLE_FILE)
    print(f"[2] 슬롯별 가능 인원 전체 저장 완료: {ALL_AVAILABLE_FILE.name}")


def calc_distance(person_display: str, target_slot: str, target_day: str, old_slot_map: dict) -> int:
    """1학기 배정 위치와의 요일 및 교시 거리 점수를 계산합니다."""
    p_name = person_display.split("(")[0].strip()
    if p_name not in old_slot_map:
        return 5  # 신규 인원은 기본 거리 부여

    old_d_name, old_d_idx, old_s_name, old_s_idx = old_slot_map[p_name]
    new_d_idx = DAYS.index(target_day)
    new_s_idx = SLOTS.index(target_slot)

    day_dist = abs(old_d_idx - new_d_idx)
    slot_dist = abs(old_s_idx - new_s_idx)
    return day_dist * 2 + slot_dist


def run_single_attempt(people, person_dept, person_available, slot_available, old_slot_map, old_partner_map):
    """1회 배정 시도: 희귀 슬롯 우선 + 거리 먼 사람 우선 배정"""
    slot_assigned = defaultdict(list)
    assigned_global = set()

    # Phase 1: 각 슬롯에 최소 1명 배정
    while True:
        unfilled = [s for s in slot_available if len(slot_assigned[s]) < 1]
        if not unfilled:
            break

        random.shuffle(unfilled)
        unfilled.sort(key=lambda s: len([p for p in slot_available[s] if p not in assigned_global]))
        t_slot, t_day = unfilled[0]

        candidates = [p for p in slot_available[(t_slot, t_day)] if p not in assigned_global]
        if not candidates:
            return None, -1

        random.shuffle(candidates)
        # 선택지가 극도로 좁은 인원 보호 + 거리 점수가 먼 사람 우선
        candidates.sort(
            key=lambda p: (
                0 if len([s for s in person_available[p] if len(slot_assigned[s]) < 1]) <= 2 else 1,
                -calc_distance(p, t_slot, t_day, old_slot_map),
                len(person_available[p]),
            )
        )
        chosen = candidates[0]
        slot_assigned[(t_slot, t_day)].append(chosen)
        assigned_global.add(chosen)

    # Phase 2.1: 2번째 인원 배정 (동일 국서 중복 배제 & 1학기 동일 파트너 배제)
    while True:
        unfilled = [s for s in slot_available if len(slot_assigned[s]) < CAPACITY_PER_SLOT]
        if not unfilled:
            break

        def count_valid(s):
            existing_depts = {person_dept.get(p) for p in slot_assigned[s]}
            return len([
                p for p in slot_available[s]
                if p not in assigned_global and person_dept.get(p) not in existing_depts
            ])

        random.shuffle(unfilled)
        unfilled.sort(key=count_valid)
        t_slot, t_day = unfilled[0]
        existing_depts = {person_dept.get(p) for p in slot_assigned[(t_slot, t_day)]}
        first_p_name = slot_assigned[(t_slot, t_day)][0].split("(")[0].strip()

        candidates = [
            p for p in slot_available[(t_slot, t_day)]
            if p not in assigned_global and person_dept.get(p) not in existing_depts
        ]

        if not candidates:
            break

        random.shuffle(candidates)
        candidates.sort(
            key=lambda p: (
                1 if p.split("(")[0].strip() in old_partner_map[first_p_name] else 0,
                0 if len([s for s in person_available[p] if len(slot_assigned[s]) < CAPACITY_PER_SLOT]) <= 2 else 1,
                -calc_distance(p, t_slot, t_day, old_slot_map),
                len(person_available[p]),
            )
        )
        chosen = candidates[0]
        slot_assigned[(t_slot, t_day)].append(chosen)
        assigned_global.add(chosen)

    # Phase 2.2: 필요 시 국서 중복 허용 잔여 슬롯 배정
    while True:
        unfilled = [s for s in slot_available if len(slot_assigned[s]) < CAPACITY_PER_SLOT]
        if not unfilled:
            break

        random.shuffle(unfilled)
        unfilled.sort(key=lambda s: len([p for p in slot_available[s] if p not in assigned_global]))
        t_slot, t_day = unfilled[0]

        candidates = [p for p in slot_available[(t_slot, t_day)] if p not in assigned_global]
        if not candidates:
            break

        random.shuffle(candidates)
        candidates.sort(key=lambda p: -calc_distance(p, t_slot, t_day, old_slot_map))
        chosen = candidates[0]
        slot_assigned[(t_slot, t_day)].append(chosen)
        assigned_global.add(chosen)

    if len(assigned_global) != len(people):
        return None, -1

    # 품질 평가 점수 산출
    score = 0
    for (s, d), assigned in slot_assigned.items():
        for p in assigned:
            score += calc_distance(p, s, d, old_slot_map)
        if len(assigned) == 2:
            p1 = assigned[0].split("(")[0].strip()
            p2 = assigned[1].split("(")[0].strip()
            if p2 in old_partner_map[p1]:
                score -= 100

    return slot_assigned, score


def optimize_assignment(people, person_dept, person_available, slot_available, old_slot_map, old_partner_map):
    """다회 셔플 시뮬레이션을 통해 제약조건을 만족하고 거리 점수가 가장 높은 최적의 배정안을 선정합니다."""
    print(f"[3] 1학기 거리 최적화 배정 알고리즘 실행 (시도 횟수: {NUM_OPTIMIZE_TRIALS}회)...")

    best_assign = None
    best_score = -999999
    success_cnt = 0

    for _ in range(NUM_OPTIMIZE_TRIALS):
        assign, score = run_single_attempt(
            people, person_dept, person_available, slot_available, old_slot_map, old_partner_map
        )
        if assign is not None:
            success_cnt += 1
            if score > best_score:
                best_score = score
                best_assign = assign

    if best_assign is None:
        raise RuntimeError("모든 조건을 만족하는 배정 조합을 찾지 못했습니다. 설문 데이터를 확인해 주세요.")

    # 공석 슬롯 채우기
    assigned_global = set()
    for slot in slot_available:
        for p in best_assign[slot]:
            assigned_global.add(p)
        while len(best_assign[slot]) < CAPACITY_PER_SLOT:
            best_assign[slot].append("(공석)")

    print(f" - {NUM_OPTIMIZE_TRIALS}회 중 유효 배정 성공: {success_cnt}회")
    print(f" - 최고 거리 최적화 점수 조합 자동 선택 완료 (총점: {best_score}점)\n")

    return best_assign, assigned_global


def validate_and_export_results(people, person_dept, slot_assigned, assigned_global, old_slot_map, old_partner_map):
    """결과 검증 및 엑셀 파일 저장"""
    print("[4] 배정 결과 검증")
    unassigned = [p for p in people if p not in assigned_global]
    deficit_slots = [s for s, assigned in slot_assigned.items() if len([p for p in assigned if "공석" not in p]) < CAPACITY_PER_SLOT]
    same_dept_slots = []
    same_partner_slots = []
    dist_list = []

    for slot, assigned in slot_assigned.items():
        valid_members = [p for p in assigned if "공석" not in p]
        if len(valid_members) == 2:
            d1 = person_dept.get(valid_members[0])
            d2 = person_dept.get(valid_members[1])
            if d1 and d2 and d1 == d2:
                same_dept_slots.append((slot, valid_members))

            p1 = valid_members[0].split("(")[0].strip()
            p2 = valid_members[1].split("(")[0].strip()
            if p2 in old_partner_map[p1]:
                same_partner_slots.append((slot, valid_members))

        for p in valid_members:
            dist = calc_distance(p, slot[0], slot[1], old_slot_map)
            dist_list.append(dist)

    avg_dist = sum(dist_list) / len(dist_list) if dist_list else 0

    print(f" - 배정 인원: {len(assigned_global)} / {len(people)}명 (100% 완료: {len(unassigned) == 0})")
    print(f" - 미배정자: {len(unassigned)}명 {unassigned if unassigned else ''}")
    print(f" - 정원 미달 슬롯: {len(deficit_slots)}개 {deficit_slots if deficit_slots else ''}")
    print(f" - 국서 중복 슬롯: {len(same_dept_slots)}개 {same_dept_slots if same_dept_slots else ''}")
    print(f" - 1학기 파트너 중복: {len(same_partner_slots)}개 {same_partner_slots if same_partner_slots else ''}")
    print(f" - 1학기 대비 평균 거리 점수: {avg_dist:.2f}점 / 13점 만점 (요일 및 시간대가 폭넓게 분산됨)")

    def safe_export(df_to_save, target_path, label):
        try:
            df_to_save.to_excel(target_path, index=(label != "개인별"))
            print(f" - {label} 저장 완료: {target_path.name}")
        except PermissionError:
            backup_path = target_path.with_name(f"{target_path.stem}_new{target_path.suffix}")
            df_to_save.to_excel(backup_path, index=(label != "개인별"))
            print(f" [안내] '{target_path.name}' 파일이 열려 있어 대신 '{backup_path.name}'으로 저장되었습니다. (엑셀을 닫고 다시 실행하면 덮어씁니다)")

    # 최종 시간표 엑셀 저장
    result_grid = pd.DataFrame("", index=SLOTS, columns=DAYS)
    for (slot, day), assigned in slot_assigned.items():
        result_grid.at[slot, day] = ", ".join(assigned)

    safe_export(result_grid, FINAL_ASSIGNMENT_FILE, "최종 시간표")

    # 개인별 배정 결과 엑셀 저장
    person_slot_map = {}
    for (slot, day), assigned in slot_assigned.items():
        for person in assigned:
            if "공석" not in person:
                person_slot_map[person] = f"{day}요일 {slot}"

    person_rows = []
    for person in sorted(people):
        dept = person_dept.get(person, "")
        assigned_slot = person_slot_map.get(person, "미배정")
        p_name = person.split("(")[0].strip()
        old_info = "신규/정보없음"
        dist_info = "-"
        if p_name in old_slot_map:
            old_day, _, old_s, _ = old_slot_map[p_name]
            old_info = f"{old_day} {old_s}"
            if assigned_slot != "미배정":
                parts = assigned_slot.split("요일 ")
                dist_info = f"{calc_distance(person, parts[1], parts[0], old_slot_map)}점"

        person_rows.append({
            "이름(국)": person,
            "소속국": dept,
            "2학기 배정": assigned_slot,
            "1학기 배정": old_info,
            "1학기 대비 거리": dist_info,
        })

    person_df = pd.DataFrame(person_rows)
    safe_export(person_df, PERSON_ASSIGNMENT_FILE, "개인별")


def main():
    old_slot_map, old_partner_map = load_1st_semester_data(OLD_ASSIGNMENT_FILE)
    people, person_dept, person_available, slot_available = load_and_preprocess_data(INPUT_FILE)
    export_all_available(slot_available)
    slot_assigned, assigned_global = optimize_assignment(
        people, person_dept, person_available, slot_available, old_slot_map, old_partner_map
    )
    validate_and_export_results(
        people, person_dept, slot_assigned, assigned_global, old_slot_map, old_partner_map
    )


if __name__ == "__main__":
    main()
