from collections import defaultdict

import pandas as pd

# 读取数据
df = pd.read_excel("卫津路面试数据(已清洗).xlsx", sheet_name="详细数据统计（按提交顺序排序）")
df = df.dropna(subset=['成员ID', '报名组别'])  # 清理空行


# 数据预处理 - 增强时间段处理逻辑
def parse_time_slots(time_str):
    if pd.isna(time_str):
        return []
    try:
        cleaned = str(time_str).replace(" ", "")
        slots = list(map(int, cleaned.split(',')))
        # 标记仅包含1-5时间段的成员
        is_only_core = all(1 <= s <= 5 for s in slots)
        # 优先级排序：1、2优先，6+其次，3-5最后
        priority_1_2 = [s for s in slots if s in {1, 2}]
        others = [s for s in slots if s >= 6]
        priority_3_5 = [s for s in slots if 3 <= s <= 5]
        return {
            'slots': priority_1_2 + others + priority_3_5,
            'only_core': is_only_core
        }
    except:
        return {'slots': [], 'only_core': False}


# 构建成员数据结构
members = []
for _, row in df.iterrows():
    time_data = parse_time_slots(row['可用时间段'])
    members.append({
        'id': row['成员ID'],
        'group': row['报名组别'].replace('组', '').strip(),
        'time_slots': time_data['slots'],
        'is_leader': row['是否组长'] == '是',
        'only_core': time_data['only_core']
    })

# 初始化数据结构
time_slot_status = defaultdict(list)
used_members = set()
group_variety = defaultdict(set)

# 时间段分类（按处理优先级）
time_priority_groups = [
    {'slots': [1, 2], 'max_cap': 5},  # 核心优先时段（最大容量5）
    {'slots': range(6, 24), 'max_cap': 5},  # 其他非3-5时段
    {'slots': [3, 4, 5], 'max_cap': 5}  # 尽量避免使用的时段
]


# 辅助函数：获取时间段最大容量
def get_max_cap(t):
    for group in time_priority_groups:
        if t in group['slots']:
            return group['max_cap']
    return 5


# 成员分类处理
def classify_members(member_list):
    classified = {
        'core_leaders': [],
        'core_non_leaders': [],
        'other_leaders': [],
        'other_non_leaders': []
    }
    for m in member_list:
        if m['only_core']:
            if m['is_leader']:
                classified['core_leaders'].append(m)
            else:
                classified['core_non_leaders'].append(m)
        else:
            if m['is_leader']:
                classified['other_leaders'].append(m)
            else:
                classified['other_non_leaders'].append(m)
    return classified


# 优化分配策略：优先填满高优先级时段
def allocate_members():
    member_pools = classify_members(members)

    # 分配阶段（按优先级顺序）
    allocation_order = [
        ('core_leaders', [1, 2]),  # 核心时段组长优先
        ('core_non_leaders', [1, 2]),  # 核心时段普通成员
        ('other_leaders', []),  # 其他组长
        ('other_non_leaders', [])  # 其他成员
    ]

    for pool_type, preferred_slots in allocation_order:
        pool = member_pools[pool_type]
        for member in pool[:]:  # 使用副本遍历
            if member['id'] in used_members:
                continue

            # 寻找可用时间段
            allocated = False
            available_times = []

            # 收集所有符合条件的时间段
            for t in member['time_slots']:
                current = time_slot_status[t]
                max_cap = get_max_cap(t)

                # 优先匹配preferred_slots
                if preferred_slots and t not in preferred_slots:
                    continue

                if len(current) < max_cap:
                    available_times.append(t)

            # 按时段剩余容量和优先级排序
            available_times.sort(key=lambda x: (
                # 排序规则：
                # 1. 剩余容量小的优先（更接近满员）
                get_max_cap(x) - len(time_slot_status[x]),
                # 2. 时段优先级（1-2 > 6+ > 3-5）
                (0 if x in [1, 2] else 1 if x >= 6 else 2)
            ))

            # 尝试分配
            for t in available_times:
                current = time_slot_status[t]
                if len(current) < get_max_cap(t):
                    current.append(member)
                    used_members.add(member['id'])
                    group_variety[t].add(member['group'])
                    allocated = True
                    break

            # 如果未分配，尝试所有可用时间段
            if not allocated:
                available_times_all = []
                for t in member['time_slots']:
                    current = time_slot_status[t]
                    if len(current) < get_max_cap(t):
                        available_times_all.append(t)

                # 按剩余容量和优先级排序
                available_times_all.sort(key=lambda x: (
                    get_max_cap(x) - len(time_slot_status[x]),
                    (0 if x in [1, 2] else 1 if x >= 6 else 2)
                ))

                for t in available_times_all:
                    current = time_slot_status[t]
                    if len(current) < get_max_cap(t):
                        current.append(member)
                        used_members.add(member['id'])
                        group_variety[t].add(member['group'])
                        allocated = True
                        break

            # 从池中移除已分配成员
            if allocated:
                member_pools[pool_type].remove(member)


# 执行分配
allocate_members()

# 后处理优化：强制迁移3-5时段成员到其他可用时段
for t in [3, 4, 5]:
    current_members = time_slot_status[t].copy()
    for member in current_members:
        # 寻找其他可用时段
        for alt_t in member['time_slots']:
            if alt_t == t:
                continue
            if len(time_slot_status[alt_t]) < get_max_cap(alt_t):
                # 执行迁移
                time_slot_status[alt_t].append(member)
                time_slot_status[t].remove(member)
                break

# 二次后处理：确保3-5时段最少化
for t in [3, 4, 5]:
    current_members = time_slot_status[t].copy()
    for member in current_members:
        # 尝试迁移到任意其他可用时段
        for alt_t in range(1, 24):
            if alt_t == t:
                continue
            if alt_t in member['time_slots']:
                if len(time_slot_status[alt_t]) < get_max_cap(alt_t):
                    time_slot_status[alt_t].append(member)
                    time_slot_status[t].remove(member)
                    break

# 生成结果
output_data = []
for t in sorted(time_slot_status.keys()):
    group = time_slot_status[t]
    if len(group) < 1:
        continue

    output_data.append({
        '时间段': t,
        '成员列表': '、'.join([m['id'] for m in group]),
        '组长存在': '是' if any(m['is_leader'] for m in group) else '否',
        '人数': len(group),
        '组别数': len({m['group'] for m in group}),
        '是否核心': '是' if t in [1, 2] else '否'
    })

# 未安排成员分析
unassigned = []
for m in members:
    if m['id'] not in used_members:
        available = m['time_slots'] if m['time_slots'] else ['无']
        unassigned.append({
            '成员ID': m['id'],
            '是否组长': '是' if m['is_leader'] else '否',
            '报名组别': m['group'],
            '可用时间段': ','.join(map(str, available)),
            '原因': '无可用时段' if not m['time_slots'] else '时段冲突'
        })

# 保存结果
with pd.ExcelWriter('卫津路面试安排.xlsx') as writer:
    pd.DataFrame(output_data).to_excel(writer, sheet_name='正式安排', index=False)

    if unassigned:
        pd.DataFrame(unassigned).to_excel(writer, sheet_name='未安排成员', index=False)

    # 生成统计信息
    stats = {
        '总人数': [len(members)],
        '已安排': [len(used_members)],
        '未安排': [len(unassigned)],
        '3-5时段使用数': [sum(len(time_slot_status[t]) for t in [3, 4, 5])],
        '平均每组人数': [
            f"{sum(len(g) for g in time_slot_status.values()) / len(output_data):.1f}" if output_data else 'N/A']
    }
    pd.DataFrame(stats).to_excel(writer, sheet_name='统计概览', index=False)

print("安排完成！")
print(f"总人数: {len(members)}")
print(f"已安排: {len(used_members)}")
print(f"未安排: {len(unassigned)}")
print(f"3-5时段使用人数: {sum(len(time_slot_status[t]) for t in [3, 4, 5])}")
print("详细结果请查看：卫津路面试安排.xlsx")