"""
26-2학기 지킴이 시간표 중간이후 셔플(재배정) 자동화 프로그램

[배정 및 셔플 원칙]
1. [Hard] 1인 1타임 배정 (총 59명 전원 1타임씩 배정, 60석 중 1석은 공석)
2. [Hard] 슬롯당 정원 2명 (정원 부족 1개 슬롯은 '(공석)' 처리)
3. [Hard] 동일 슬롯 내 같은 국(부서) 인원 중복 엄격 금지 (Phase 2.1 원칙)
4. [Hard] 국원 본인이 응답한 가용 시간대 내에서만 배정
5. [Hard / 지정 조건] 이수연 학우 2학기 첫 번째 시간(월요일 1:55-3:20) 고정 배정
6. [Soft / 셔플 핵심] 2학기 1차(전반기) 배정 요일 및 시간대 대비 변경 극대화 (요일 변경 최우선 보장)
7. [Soft / 셔플 핵심] 2학기 1차 파트너와의 재매칭 100% 배제 (파트너 로테이션)
8. [Soft] 1학기 당시 배정 요일/시간 및 1학기 파트너와의 재매칭 추가 회피
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
INPUT_FILE = BASE_DIR / "26-2학기 지킴이 시간 조사(응답) - 중간이후.xlsx"
OLD_26_2_FILE = BASE_DIR / "output_26_2" / "final_assignment_new.xlsx"
OLD_1ST_FILE = BASE_DIR / "1학기" / "final_assignment.xlsx"

OUTPUT_DIR = BASE_DIR / "output_26_2_중간이후"
OUTPUT_DIR.mkdir(exist_ok=True)

FINAL_ASSIGNMENT_FILE = OUTPUT_DIR / "final_assignment.xlsx"
ALL_AVAILABLE_FILE = OUTPUT_DIR / "all_available.xlsx"
PERSON_ASSIGNMENT_FILE = OUTPUT_DIR / "person_assignment.xlsx"
COMPARISON_FILE = OUTPUT_DIR / "shuffle_comparison.xlsx"

DAYS = ["월", "화", "수", "목", "금"]
SLOTS = [
    "11:00-12:20",
    "12:25-1:50",
    "1:55-3:20",
    "3:25-4:50",
    "4:55-6:20",
    "6:25-7:50",
]
ALL_SLOT_KEYS = [(d, s) for d in DAYS for s in SLOTS]

CAPACITY_PER_SLOT = 2
NUM_OPTIMIZE_TRIALS = 3000

# [고정 조건 대상자]
FIXED_PERSON_NAME = "이수연"
FIXED_TARGET_SLOT = ("월", "1:55-3:20")  # 이수연 학우의 2학기 1차 배정 슬롯


def normalize_slot_idx(s_str: str) -> int:
    """1학기 및 2학기의 다양한 슬롯 문자열 표기를 0~5 교시 인덱스로 통일합니다."""
    s_clean = str(s_str).strip()
    if s_clean in SLOTS:
        return SLOTS.index(s_clean)
    mapping = {
        "11:00-12:20": 0, "12:25-13:50": 1, "13:55-15:20": 2,
        "15:25-16:50": 3, "16:55~18:20": 4, "18:25~19:50": 5,
        "slot1": 0, "slot2": 1, "slot3": 2, "slot4": 3, "slot5": 4, "slot6": 5
    }
    return mapping.get(s_clean, 0)


def load_previous_assignments():
    """2학기 1차 배정 및 1학기 배정 내역을 로드합니다."""
    print("[0] 이전 배정 데이터 로드")
    
    # 2학기 1차 배정 데이터 로드
    old_26_2_slot = {}
    old_26_2_partners = defaultdict(set)
    if OLD_26_2_FILE.exists():
        grid_26_2 = pd.read_excel(OLD_26_2_FILE, index_col=0)
        for s in grid_26_2.index:
            for d in grid_26_2.columns:
                members = [m.strip() for m in str(grid_26_2.at[s, d]).split(',') if m.strip() and '공석' not in m]
                for m1 in members:
                    old_26_2_slot[m1] = (d, s)
                    for m2 in members:
                        if m1 != m2:
                            old_26_2_partners[m1].add(m2)
        print(f" - 2학기 1차 배정 인원: {len(old_26_2_slot)}명 로드 완료")
    else:
        print(f" ⚠️ 2학기 1차 파일({OLD_26_2_FILE.name})이 없습니다.")

    # 1학기 배정 데이터 로드
    old_1st_slot = {}
    old_1st_partners = defaultdict(set)
    if OLD_1ST_FILE.exists():
        grid_1st = pd.read_excel(OLD_1ST_FILE, index_col=0)
        for s in grid_1st.index:
            for d in grid_1st.columns:
                members = [m.strip() for m in str(grid_1st.at[s, d]).split(',') if m.strip() and '공석' not in m]
                names = [m.split('(')[0].strip() for m in members]
                for i, n1 in enumerate(names):
                    old_1st_slot[n1] = (d, s)
                    for j, n2 in enumerate(names):
                        if i != j:
                            old_1st_partners[n1].add(n2)
        print(f" - 1학기 최종 배정 인원: {len(old_1st_slot)}명 로드 완료\n")
        
    return old_26_2_slot, old_26_2_partners, old_1st_slot, old_1st_partners


def load_and_preprocess_survey(file_path: Path):
    """중간이후 설문 엑셀 데이터를 정제하고 검증합니다."""
    print(f"[1] 설문 데이터 로드 및 정제: {file_path.name}")
    raw_df = pd.read_excel(file_path)
    print(f" - 원본 응답 건수: {len(raw_df)}건")

    # 이름 정제: 공백 및 (최종) 등 수식어 제거
    raw_df["이름_정제"] = raw_df["이름"].astype(str).str.strip().apply(lambda x: re.sub(r"\(.*?\)", "", x).strip())

    # 중복 제출자 감지
    dup_counts = raw_df["이름_정제"].value_counts()
    dup_names = dup_counts[dup_counts > 1].index.tolist()
    if dup_names:
        print(f" - 중복 응답자 감지 ({len(dup_names)}명): {', '.join(dup_names)}")
        for name in dup_names:
            sub = raw_df[raw_df["이름_정제"] == name]
            print(f"   * {name}: 총 {len(sub)}회 제출 -> 최신 타임스탬프 응답 유지")

    # 최신 응답 유지
    df = raw_df.sort_values(by="타임스탬프").drop_duplicates(subset=["이름_정제"], keep="last").reset_index(drop=True)
    print(f" - 중복 제거 후 최종 고유 인원: {len(df)}명\n")

    # 요일 컬럼 추출
    day_cols = {}
    for col in df.columns:
        match = re.search(r"\[([월화수목금])\]", str(col))
        if match:
            day_cols[match.group(1)] = col

    missing_days = [d for d in DAYS if d not in day_cols]
    if missing_days:
        raise ValueError(f"요일 컬럼 누락: {missing_days}")

    people = []
    person_dept = {}
    person_available = defaultdict(set)
    slot_available = defaultdict(set)

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
                    slot_key = (day, t)
                    person_available[display_name].add(slot_key)
                    slot_available[slot_key].add(display_name)

    return people, person_dept, person_available, slot_available


def export_all_available(slot_available: dict):
    """각 슬롯별 지원 가능한 전원 명단을 엑셀로 저장합니다."""
    grid = pd.DataFrame("", index=SLOTS, columns=DAYS)
    for d in DAYS:
        for s in SLOTS:
            candidates = sorted(slot_available.get((d, s), []))
            grid.at[s, d] = "\n".join(candidates) if candidates else "(가능자 없음)"

    grid.to_excel(ALL_AVAILABLE_FILE)
    print(f"[2] 슬롯별 가용 인원 전체 저장 완료: {ALL_AVAILABLE_FILE.name}")


def calc_shuffle_score(p_disp: str, target_day: str, target_slot: str,
                       old_26_2_slot: dict, old_1st_slot: dict, fixed_person_disp: str) -> float:
    """
    셔플 목적함수 점수를 계산합니다.
    - 2학기 1차 대비 요일 변경을 가장 강력하게 보상합니다.
    - 2학기 1차 대비 시간대 변경을 추가 보상합니다.
    - 1학기 배정 위치와의 거리도 보조적으로 반영합니다.
    """
    if p_disp == fixed_person_disp:
        return 100.0  # 고정 인원은 기본 점수

    score = 0.0
    if p_disp in old_26_2_slot:
        prev_d, prev_s = old_26_2_slot[p_disp]
        d1 = DAYS.index(prev_d)
        d2 = DAYS.index(target_day)
        s1 = normalize_slot_idx(prev_s)
        s2 = normalize_slot_idx(target_slot)

        day_diff = abs(d1 - d2)
        circ_day_diff = min(day_diff, 5 - day_diff)
        slot_diff = abs(s1 - s2)

        if day_diff == 0:
            score -= 100.0  # 2학기 1차와 동일 요일 배정 시 페널티
        else:
            score += 30.0 + circ_day_diff * 10.0 + slot_diff * 2.0
    else:
        score += 30.0

    p_name = p_disp.split("(")[0].strip()
    if p_name in old_1st_slot:
        prev_d1, prev_s1 = old_1st_slot[p_name]
        d1 = DAYS.index(prev_d1)
        d2 = DAYS.index(target_day)
        s1 = normalize_slot_idx(prev_s1)
        s2 = normalize_slot_idx(target_slot)
        day_diff_1st = abs(d1 - d2)
        circ_1st = min(day_diff_1st, 5 - day_diff_1st)
        slot_diff_1st = abs(s1 - s2)
        score += circ_1st * 2.0 + slot_diff_1st * 0.5

    return score


def run_single_shuffle_attempt(people, person_dept, person_available, slot_available,
                               old_26_2_slot, old_26_2_partners, old_1st_slot, old_1st_partners,
                               fixed_person_disp):
    """
    1회 무작위 셔플 배정 시도:
    - 이수연 고정 슬롯 우선 배정
    - 제약조건이 큰 희귀 슬롯 우선 충원
    - 동일 국 배제 & 2학기 1차 파트너 재매칭 배제
    """
    slot_assigned = defaultdict(list)
    assigned_global = set()

    # 1. 이수연 학우 고정 배정
    if fixed_person_disp in people:
        if FIXED_TARGET_SLOT not in person_available[fixed_person_disp]:
            raise ValueError(f"{fixed_person_disp} 학우의 가용 시간에 고정 슬롯 {FIXED_TARGET_SLOT}이 없습니다.")
        slot_assigned[FIXED_TARGET_SLOT].append(fixed_person_disp)
        assigned_global.add(fixed_person_disp)

    # Phase 1: 모든 30개 슬롯에 최소 1명씩 우선 배정
    while True:
        unfilled = [s for s in ALL_SLOT_KEYS if len(slot_assigned[s]) < 1]
        if not unfilled:
            break

        random.shuffle(unfilled)
        unfilled.sort(key=lambda s: len([p for p in slot_available[s] if p not in assigned_global]))
        target_s = unfilled[0]

        candidates = [p for p in slot_available[target_s] if p not in assigned_global]
        if not candidates:
            return None, -999999

        random.shuffle(candidates)
        candidates.sort(key=lambda p: (
            0 if len([s for s in person_available[p] if len(slot_assigned[s]) < 1]) <= 2 else 1,
            -calc_shuffle_score(p, target_s[0], target_s[1], old_26_2_slot, old_1st_slot, fixed_person_disp),
            len(person_available[p])
        ))

        chosen = candidates[0]
        slot_assigned[target_s].append(chosen)
        assigned_global.add(chosen)

    # Phase 2: 각 슬롯의 2번째 인원 배정 (동일 국 중복 배제 & 이전 파트너 배제)
    while True:
        unfilled = [s for s in ALL_SLOT_KEYS if len(slot_assigned[s]) < CAPACITY_PER_SLOT]
        if not unfilled:
            break

        def count_valid(s):
            exist_dept = {person_dept[p] for p in slot_assigned[s]}
            p1 = slot_assigned[s][0] if slot_assigned[s] else None
            return len([
                p for p in slot_available[s]
                if p not in assigned_global
                and person_dept[p] not in exist_dept
                and (p1 is None or p not in old_26_2_partners[p1])
            ])

        random.shuffle(unfilled)
        unfilled.sort(key=count_valid)
        target_s = unfilled[0]

        exist_dept = {person_dept[p] for p in slot_assigned[target_s]}
        p1 = slot_assigned[target_s][0] if slot_assigned[target_s] else None

        candidates = [
            p for p in slot_available[target_s]
            if p not in assigned_global
            and person_dept[p] not in exist_dept
            and (p1 is None or p not in old_26_2_partners[p1])
        ]

        if not candidates:
            # 59명이므로 정확히 1개 슬롯만 1인으로 남으면 정상 종료
            if len(assigned_global) == len(people):
                break
            return None, -999999

        random.shuffle(candidates)
        p1_name = p1.split("(")[0].strip() if p1 else ""
        candidates.sort(key=lambda p: (
            1 if p.split("(")[0].strip() in old_1st_partners[p1_name] else 0,
            0 if len([s for s in person_available[p] if len(slot_assigned[s]) < CAPACITY_PER_SLOT]) <= 2 else 1,
            -calc_shuffle_score(p, target_s[0], target_s[1], old_26_2_slot, old_1st_slot, fixed_person_disp),
            len(person_available[p])
        ))

        chosen = candidates[0]
        slot_assigned[target_s].append(chosen)
        assigned_global.add(chosen)

    if len(assigned_global) != len(people):
        return None, -999999

    # 품질 평가 점수 합산
    total_score = 0.0
    for (d, s), assigned in slot_assigned.items():
        for p in assigned:
            total_score += calc_shuffle_score(p, d, s, old_26_2_slot, old_1st_slot, fixed_person_disp)
        if len(assigned) == 2:
            p1, p2 = assigned
            if p2 in old_26_2_partners[p1]:
                total_score -= 10000.0  # 파트너 중복 절대 방지
            p1_n = p1.split("(")[0].strip()
            p2_n = p2.split("(")[0].strip()
            if p2_n in old_1st_partners[p1_n]:
                total_score -= 200.0   # 1학기 파트너 중복 방지

    return slot_assigned, total_score


def optimize_shuffle(people, person_dept, person_available, slot_available,
                     old_26_2_slot, old_26_2_partners, old_1st_slot, old_1st_partners,
                     fixed_person_disp):
    """다회 셔플 시뮬레이션을 통해 최적의 배정안을 선정합니다."""
    print(f"[3] 셔플 최적화 알고리즘 실행 (시뮬레이션 횟수: {NUM_OPTIMIZE_TRIALS}회)...")

    best_assign = None
    best_score = -999999
    success_cnt = 0

    for _ in range(NUM_OPTIMIZE_TRIALS):
        assign, score = run_single_shuffle_attempt(
            people, person_dept, person_available, slot_available,
            old_26_2_slot, old_26_2_partners, old_1st_slot, old_1st_partners,
            fixed_person_disp
        )
        if assign is not None:
            success_cnt += 1
            if score > best_score:
                best_score = score
                best_assign = assign

    if best_assign is None:
        raise RuntimeError("모든 조건을 만족하는 배정 조합을 찾지 못했습니다. 설문 및 가용 시간을 확인해 주세요.")

    # 공석 슬롯 명시 처리
    assigned_global = set()
    for s_key in ALL_SLOT_KEYS:
        for p in best_assign[s_key]:
            assigned_global.add(p)
        while len(best_assign[s_key]) < CAPACITY_PER_SLOT:
            best_assign[s_key].append("(공석)")

    print(f" - {NUM_OPTIMIZE_TRIALS}회 중 유효 배정 성공: {success_cnt}회")
    print(f" - 최적 셔플 배정안 선정 완료 (최고 품질 점수: {best_score:.1f}점)\n")

    return best_assign, assigned_global


def validate_and_export_results(people, person_dept, slot_assigned, assigned_global,
                                old_26_2_slot, old_26_2_partners, old_1st_slot, old_1st_partners,
                                fixed_person_disp):
    """결과 검증 및 엑셀 산출물 저장"""
    print("[4] 셔플 배정 결과 검증")
    unassigned = [p for p in people if p not in assigned_global]
    deficit_slots = [s for s, assigned in slot_assigned.items() if len([p for p in assigned if "공석" not in p]) < CAPACITY_PER_SLOT]
    same_dept_slots = []
    same_partner_26_2_slots = []
    same_partner_1st_slots = []
    same_day_members = []
    different_day_members = []

    for (d, s), assigned in slot_assigned.items():
        valid_members = [p for p in assigned if "공석" not in p]
        if len(valid_members) == 2:
            d1 = person_dept.get(valid_members[0])
            d2 = person_dept.get(valid_members[1])
            if d1 and d2 and d1 == d2:
                same_dept_slots.append(((d, s), valid_members))

            p1, p2 = valid_members
            if p2 in old_26_2_partners[p1]:
                same_partner_26_2_slots.append(((d, s), valid_members))

            p1_n = p1.split("(")[0].strip()
            p2_n = p2.split("(")[0].strip()
            if p2_n in old_1st_partners[p1_n]:
                same_partner_1st_slots.append(((d, s), valid_members))

        for p in valid_members:
            if p != fixed_person_disp and p in old_26_2_slot:
                if old_26_2_slot[p][0] == d:
                    same_day_members.append((p, d, s, old_26_2_slot[p]))
                else:
                    different_day_members.append(p)

    print(f" - 전체 인원 배정: {len(assigned_global)} / {len(people)}명 (100% 완료: {len(unassigned) == 0})")
    print(f" - 미배정자: {len(unassigned)}명")
    print(f" - 공석 슬롯 (정원 1인 슬롯): {deficit_slots}")
    print(f" - 국서 중복 슬롯 (0개 권장): {len(same_dept_slots)}개 {same_dept_slots if same_dept_slots else '✅ 완벽 배제'}")
    print(f" - 2학기 1차 파트너 재매칭 (0개 권장): {len(same_partner_26_2_slots)}개 {same_partner_26_2_slots if same_partner_26_2_slots else '✅ 완벽 배제'}")
    print(f" - 1학기 파트너 재매칭: {len(same_partner_1st_slots)}개")
    print(f" - [고정 확인] {fixed_person_disp} -> {slot_assigned[FIXED_TARGET_SLOT]}")
    print(f" - 요일 변경 달성률: {len(different_day_members)} / {len(people)-1}명 ({len(different_day_members)/(len(people)-1)*100:.1f}%)")
    if same_day_members:
        print(f"   * 동일 요일 유지자 ({len(same_day_members)}명):")
        for p, d, s, old_info in same_day_members:
            print(f"     - {p}: 1차({old_info[0]} {old_info[1]}) -> 2차({d} {s})")

    # 1. 최종 시간표 그리드 저장
    result_grid = pd.DataFrame("", index=SLOTS, columns=DAYS)
    for (d, s), assigned in slot_assigned.items():
        result_grid.at[s, d] = ", ".join(assigned)

    result_grid.to_excel(FINAL_ASSIGNMENT_FILE)
    print(f"\n - [최종 시간표] 저장 완료: {FINAL_ASSIGNMENT_FILE.name}")

    # 2. 개인별 배정 현황 및 비교 저장
    person_slot_map = {}
    person_partner_map = {}
    for (d, s), assigned in slot_assigned.items():
        valid_members = [p for p in assigned if "공석" not in p]
        for p in valid_members:
            person_slot_map[p] = f"{d}요일 {s}"
            partners = [other for other in assigned if other != p]
            person_partner_map[p] = ", ".join(partners)

    comparison_rows = []
    for person in sorted(people):
        dept = person_dept.get(person, "")
        new_slot = person_slot_map.get(person, "미배정")
        new_partner = person_partner_map.get(person, "(없음)")

        old_26_2_info = "신규/정보없음"
        day_change_info = "신규"
        partner_change_info = "신규"
        old_partner_str = "(없음)"

        if person in old_26_2_slot:
            o_d, o_s = old_26_2_slot[person]
            old_26_2_info = f"{o_d}요일 {o_s}"
            old_parts = sorted(list(old_26_2_partners.get(person, [])))
            old_partner_str = ", ".join(old_parts) if old_parts else "(공석)"

            if person == fixed_person_disp:
                day_change_info = "고정 (지정 조건)"
            else:
                n_d = new_slot.split("요일")[0]
                day_change_info = f"요일 변경 ({o_d} -> {n_d})" if o_d != n_d else f"동일 요일 유지 ({o_d})"

            partner_change_info = f"파트너 로테이션 완료" if new_partner not in old_parts else "파트너 동일 (주의)"

        p_name = person.split("(")[0].strip()
        old_1st_info = "신규/정보없음"
        if p_name in old_1st_slot:
            d_1st, s_1st = old_1st_slot[p_name]
            old_1st_info = f"{d_1st}요일 {s_1st}"

        comparison_rows.append({
            "이름(국)": person,
            "소속국": dept,
            "중간이후 배정": new_slot,
            "중간이후 파트너": new_partner,
            "2학기 1차 배정": old_26_2_info,
            "2학기 1차 파트너": old_partner_str,
            "요일 변경 여부": day_change_info,
            "파트너 변경 여부": partner_change_info,
            "1학기 배정": old_1st_info
        })

    comp_df = pd.DataFrame(comparison_rows)
    comp_df.to_excel(PERSON_ASSIGNMENT_FILE, index=False)
    comp_df.to_excel(COMPARISON_FILE, index=False)
    print(f" - [개인별 배정표] 저장 완료: {PERSON_ASSIGNMENT_FILE.name}")
    print(f" - [셔플 전후 비교표] 저장 완료: {COMPARISON_FILE.name}")


def main():
    old_26_2_slot, old_26_2_partners, old_1st_slot, old_1st_partners = load_previous_assignments()
    people, person_dept, person_available, slot_available = load_and_preprocess_survey(INPUT_FILE)
    
    fixed_matches = [p for p in people if FIXED_PERSON_NAME in p]
    if not fixed_matches:
        raise ValueError(f"고정 대상자 '{FIXED_PERSON_NAME}' 학우를 설문에서 찾을 수 없습니다.")
    fixed_person_disp = fixed_matches[0]
    print(f"[알림] 고정 대상자 확인: {fixed_person_disp} -> 고정 슬롯: {FIXED_TARGET_SLOT[0]}요일 {FIXED_TARGET_SLOT[1]}\n")

    export_all_available(slot_available)

    slot_assigned, assigned_global = optimize_shuffle(
        people, person_dept, person_available, slot_available,
        old_26_2_slot, old_26_2_partners, old_1st_slot, old_1st_partners,
        fixed_person_disp
    )

    validate_and_export_results(
        people, person_dept, slot_assigned, assigned_global,
        old_26_2_slot, old_26_2_partners, old_1st_slot, old_1st_partners,
        fixed_person_disp
    )


if __name__ == "__main__":
    main()

