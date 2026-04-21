import pandas as pd
from collections import defaultdict
import random
import math

# =========================
# 1. 파일 및 환경 설정
# =========================
INPUT_FILE = "input_final.xlsx"
MID_FILE = "mid_assignment.xlsx"
OUTPUT_FILE = "final_assignment.xlsx"

DAYS = ["월", "화", "수", "목", "금"]

# =========================
# 2. 데이터 로드 (입력 및 이전 배정)
# =========================
print("데이터를 불러오는 중입니다...")
df = pd.read_excel(INPUT_FILE)
# 엑셀의 시간대 헤더를 그대로 가져옵니다 (예: '11:00-12:20')
SLOTS = df.columns[1:].tolist()
people = df["이름"].dropna().tolist()

# 각 사람이 배정 가능한 시간대 (Set)
avail_dict = defaultdict(set)
for _, row in df.iterrows():
    name = row["이름"]
    if pd.isna(name): continue
    for slot in SLOTS:
        cell = row.get(slot)
        if pd.isna(cell): continue
        days = [d.strip() for d in str(cell).split(",")]
        for d in days:
            if d in DAYS:
                avail_dict[name].add((d, slot))

# 이전 배정(mid_assignment) 내역 파악
old_day = {}
old_slot = {}
old_partner_dict = defaultdict(list)

try:
    mid_df = pd.read_excel(MID_FILE, index_col=0)
    for slot in mid_df.index:
        for day in DAYS:
            if day not in mid_df.columns: continue
            cell = mid_df.at[slot, day]
            if pd.isna(cell): continue
            
            assigned = [p.strip() for p in str(cell).split(",")]
            for i, p1 in enumerate(assigned):
                if p1.startswith("(없음"): continue
                old_day[p1] = day
                old_slot[p1] = slot
                for j, p2 in enumerate(assigned):
                    if i != j and not p2.startswith("(없음"):
                        old_partner_dict[p1].append(p2)
except Exception as e:
    print(f"이전 배정 파일({MID_FILE})을 불러오는 중 문제가 발생했습니다: {e}")

# =========================
# 3. 페널티 (에너지) 계산 함수
# =========================
def calc_energy(assign):
    """배정 상태에 대한 페널티(에너지)를 계산합니다. 점수가 낮을수록 좋습니다."""
    energy = 0
    slot_people = defaultdict(list)
    
    for p, (d, s) in assign.items():
        slot_people[(d, s)].append(p)
        
    for (d, s), p_list in slot_people.items():
        # [Hard Constraint] 각 슬롯은 무조건 2명
        if len(p_list) > 2:
            energy += 100000 * (len(p_list) - 2)
            
        # 2명이 배정된 경우 파트너 페널티 확인
        elif len(p_list) == 2:
            p1, p2 = p_list
            # [Soft Constraint] 이전과 파트너가 같으면 페널티
            if p2 in old_partner_dict[p1]:
                energy += 5000 
                
    for p, (d, s) in assign.items():
        old_d = old_day.get(p)
        old_s = old_slot.get(p)
        
        # [Soft Constraint] 이전 요일과의 차이 (원형 차이)
        if old_d and old_d in DAYS:
            d1_idx = DAYS.index(old_d)
            d2_idx = DAYS.index(d)
            # 월(0)과 목(3)의 차이: min(|0-3|, 5-|0-3|) = min(3, 2) = 2
            dist_d = min(abs(d1_idx - d2_idx), 5 - abs(d1_idx - d2_idx))
            if dist_d == 0:
                energy += 200  # 같은 요일
            elif dist_d == 1:
                energy += 100  # 하루 차이
            else: # dist_d == 2
                energy += 0    # 가장 먼 요일 (최상)
        
        # [Soft Constraint] 이전 시간대(슬롯)와의 차이
        if old_s and old_s in SLOTS:
            if s in SLOTS:
                s1_idx = SLOTS.index(old_s)
                s2_idx = SLOTS.index(s)
                dist_s = abs(s1_idx - s2_idx)
                # 차이가 클수록 페널티 감소 (최대 차이 5)
                energy += (5 - dist_s) * 10
                
    return energy

# =========================
# 4. 초기 해 (Initial Solution) 생성
# =========================
def get_initial_assignment():
    """조건이 적은 사람부터 차례대로 최대한 빈 자리에 구겨 넣습니다."""
    assign = {}
    slot_counts = defaultdict(int)
    sorted_people = sorted(people, key=lambda p: len(avail_dict[p]))
    
    for p in sorted_people:
        # 2명이 다 안 찬 가능한 자리 찾기
        available_slots = [ds for ds in avail_dict[p] if slot_counts[ds] < 2]
        if available_slots:
            chosen = random.choice(available_slots)
        else:
            # 다 찼으면 일단 가능한 곳 중 그나마 덜 찬 곳으로 배치
            if avail_dict[p]:
                chosen = min(avail_dict[p], key=lambda ds: slot_counts[ds])
            else:
                # 아예 가능한 곳이 없는 경우 (예외)
                chosen = (DAYS[0], SLOTS[0])
                
        assign[p] = chosen
        slot_counts[chosen] += 1
        
    return assign

# =========================
# 5. 시뮬레이티드 어닐링 최적화
# =========================
def optimize(initial_assign, iterations=40000):
    current_assign = initial_assign.copy()
    current_energy = calc_energy(current_assign)
    
    best_assign = current_assign.copy()
    best_energy = current_energy
    
    T = 2000.0  # 초기 온도
    cooling_rate = (0.01 / T) ** (1 / iterations) 
    
    for _ in range(iterations):
        new_assign = current_assign.copy()
        action = random.choice([1, 2])
        
        if action == 1: 
            # Action 1: 두 사람의 스케줄 맞교환 (Swap)
            p1, p2 = random.sample(people, 2)
            d1, s1 = new_assign[p1]
            d2, s2 = new_assign[p2]
            
            # 교환하려는 자리가 각자에게 허용된 자리일 때만 수행
            if (d2, s2) in avail_dict[p1] and (d1, s1) in avail_dict[p2]:
                new_assign[p1] = (d2, s2)
                new_assign[p2] = (d1, s1)
            else:
                continue
                
        else: 
            # Action 2: 한 사람을 다른 빈 자리(혹은 다른 자리)로 이동 (Move)
            p1 = random.choice(people)
            if not avail_dict[p1]: continue
            target_ds = random.choice(list(avail_dict[p1]))
            new_assign[p1] = target_ds
            
        new_energy = calc_energy(new_assign)
        
        # 더 좋거나(낮은 에너지), 일정 확률로 덜 좋은 조건도 수용
        if new_energy < current_energy or random.random() < math.exp((current_energy - new_energy) / T):
            current_assign = new_assign
            current_energy = new_energy
            
            if current_energy < best_energy:
                best_energy = current_energy
                best_assign = current_assign.copy()
        
        T *= cooling_rate
        
    return best_assign, best_energy

# =========================
# 6. 메인 실행 루프
# =========================
print("최적의 배정을 찾기 위해 시뮬레이션을 여러 번 수행합니다...")
NUM_RESTARTS = 500  # 여러 번 돌려서 가장 좋은 결과를 찾음
best_overall_assign = None
best_overall_energy = float('inf')

for i in range(NUM_RESTARTS):
    init_assign = get_initial_assignment()
    # 4만번씩 스왑 시도
    result_assign, result_energy = optimize(init_assign, iterations=40000)
    print(f"시도 {i+1}/{NUM_RESTARTS} - 완료 (페널티 점수: {result_energy})")
    
    if result_energy < best_overall_energy:
        best_overall_energy = result_energy
        best_overall_assign = result_assign

print(f"\n최종 선택된 최고 배정안 페널티 점수: {best_overall_energy}")

# =========================
# 7. 최종 결과 엑셀로 저장
# =========================
result_df = pd.DataFrame("", index=SLOTS, columns=DAYS)
slot_assigned = defaultdict(list)

for p, (d, s) in best_overall_assign.items():
    slot_assigned[(d, s)].append(p)

for slot in SLOTS:
    for day in DAYS:
        assigned = slot_assigned.get((day, slot), [])
        
        # 정원이 모자란 경우 빈칸으로 표시
        while len(assigned) < 2:
            assigned.append("(없음)")
            
        result_df.at[slot, day] = ", ".join(assigned)

result_df.to_excel(OUTPUT_FILE)
print(f"\n완료! 결과가 '{OUTPUT_FILE}' 에 저장되었습니다.")