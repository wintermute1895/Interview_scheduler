import datetime
import random

import pandas as pd


def generate_time_slots():
    time_slots = []
    start = datetime.datetime.strptime("10:00", "%H:%M")
    end = datetime.datetime.strptime("20:00", "%H:%M")
    delta = datetime.timedelta(minutes=15)
    current = start

    while current <= end:
        current_time = current.time()
        if (current_time >= datetime.time(10, 0) and current_time < datetime.time(12, 0)) or \
                (current_time >= datetime.time(14, 0) and current_time <= datetime.time(20, 0)):
            time_slots.append(current.strftime("%H:%M"))
        current += delta

    # 生成时间段区间
    formatted = []
    for i in range(len(time_slots) - 1):
        start_t = time_slots[i]
        end_t = time_slots[i + 1]
        formatted.append(f"{start_t}-{end_t}")
    return formatted

def load_candidates(file_path):
    # 读取Excel文件（如果是CSV使用 pd.read_csv）
    df = pd.read_excel(file_path)

    candidates = []
    for _, row in df.iterrows():
        # 转换时间段格式
        time_slots = [ts.strip() for ts in row['可用时间段'].split(',')]

        # 转换组长标识
        is_leader = row['是否组长'] in ['是', True, 'Yes', 'Y']

        candidates.append({
            'id': row['成员ID'],
            'group': row['报名组别'],
            'available_times': time_slots,
            'is_leader': is_leader
        })
    return candidates


# 使用真实数据文件
candidates = load_candidates("测试用例5_综合问题.xlsx")  # 修改为你的文件路径
groups = ['宣传', '教务', '后勤', '调研']  # 保持原有组别定义
time_slots = generate_time_slots()  # 保持原有时间段生成逻辑
# 编排面试
schedule = []
scheduled_ids = set()

for slot in time_slots:
    # 获取当前时间段可用且未安排的候选人
    available = [c for c in candidates if c['id'] not in scheduled_ids and slot in c['available_times']]

    if len(available) < 10:
        continue

    # 筛选组长
    leaders = [c for c in available if c['is_leader']]
    if not leaders:
        continue

    # 随机选择1个组长
    leader = random.choice(leaders)
    scheduled_ids.add(leader['id'])
    group_members = [leader]
    current_groups = {leader['group']}

    # 准备剩余可用人员
    remaining_available = [c for c in available if c['id'] != leader['id']]

    # 选择剩余9人
    for _ in range(9):
        if not remaining_available:
            break

        # 优先选择未覆盖组别
        needed_groups = set(groups) - current_groups
        candidates_needed = [c for c in remaining_available if c['group'] in needed_groups]

        if candidates_needed:
            selected = random.choice(candidates_needed)
        else:
            selected = random.choice(remaining_available)

        group_members.append(selected)
        scheduled_ids.add(selected['id'])
        current_groups.add(selected['group'])
        remaining_available.remove(selected)

    # 记录安排
    for member in group_members:
        schedule.append({
            '时间段': slot,
            '成员ID': member['id'],
            '报名组别': member['group'],
            '是否组长': '是' if member['is_leader'] else '否'
        })

# 导出到Excel
df = pd.DataFrame(schedule)
df = df.sort_values('时间段')
df.to_excel('面试安排表6_综合测试.xlsx', index=False)

print("面试安排已生成到 面试安排表.xlsx")