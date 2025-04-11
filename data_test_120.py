import random
from datetime import datetime, timedelta

import pandas as pd


def generate_test_data():
    # 基础设置
    total_candidates = 120
    groups = ['宣传', '教务', '后勤', '调研']
    time_slots = []

    # 生成所有合法时间段
    start = datetime.strptime("10:00", "%H:%M")
    end = datetime.strptime("20:00", "%H:%M")
    delta = timedelta(minutes=15)
    current = start
    while current <= end:
        if not (12 <= current.hour < 14):
            end_time = current + delta
            time_slots.append(f"{current.strftime('%H:%M')}-{end_time.strftime('%H:%M')}")
        current += delta

    # 特殊设置
    special_config = {
        # 组别分布（宣传、教务、后勤、调研）
        'group_dist': [45, 40, 30, 5],  # 调研组只有5人

        # 特殊时间段设置
        'problem_slots': {
            '10:15-10:30': {'total': 7, 'leaders': 0},  # 人数不足时间段
            '14:00-14:15': {'total': 15, 'leaders': 1},  # 组长不足时间段
            '15:30-15:45': {'total': 8, 'leaders': 0}  # 无组长时间段
        },

        # 全局组长比例
        'global_leader_ratio': 0.15  # 15%的全局组长比例,可以降低一点
    }

    data = []
    candidate_id = 1000
    group_counter = {g: 0 for g in groups}

    # 生成特殊时间段候选人
    problem_candidates = []
    for slot, config in special_config['problem_slots'].items():
        for _ in range(config['total']):
            # 强制分配到调研组（制造组别不足）
            group = '调研' if group_counter['调研'] < 5 else random.choice(groups[:3])
            is_leader = False

            # 分配组长身份
            if config['leaders'] > 0:
                is_leader = True
                config['leaders'] -= 1
            else:
                is_leader = random.random() < special_config['global_leader_ratio']

            problem_candidates.append({
                'id': candidate_id,
                'group': group,
                'times': [slot],
                'is_leader': is_leader
            })
            candidate_id += 1
            group_counter[group] += 1

    # 生成常规候选人
    remaining = total_candidates - len(problem_candidates)
    for _ in range(remaining):
        # 按预设权重分配组别
        group = random.choices(groups, weights=special_config['group_dist'], k=1)[0]
        if group == '调研' and group_counter['调研'] >= 5:
            group = random.choice(groups[:3])  # 确保调研组只有5人

        # 分配时间段（3-5个随机时间段）
        available_slots = random.sample(time_slots[2:-2], k=random.randint(3, 5))

        # 避开问题时间段
        available_slots = [s for s in available_slots if s not in special_config['problem_slots']]

        # 分配组长身份
        is_leader = random.random() < special_config['global_leader_ratio']

        data.append({
            'id': candidate_id,
            'group': group,
            'times': available_slots,
            'is_leader': is_leader
        })
        candidate_id += 1
        group_counter[group] += 1

    # 合并所有候选人
    all_candidates = problem_candidates + data

    # 转换为DataFrame
    df = pd.DataFrame([{
        "成员ID": c['id'],
        "报名组别": c['group'],
        "可用时间段": ", ".join(c['times']),
        "是否组长": "是" if c['is_leader'] else "否"
    } for c in all_candidates])

    # 保存文件
    df.to_excel("综合测试数据120人.xlsx", index=False)

    # 打印统计信息
    print("=== 测试数据统计 ===")
    print(f"总人数: {len(df)}")
    print("\n组别分布:")
    print(df['报名组别'].value_counts())
    print("\n组长数量:")
    print(df['是否组长'].value_counts())
    print("\n特殊时间段人数:")
    for slot in special_config['problem_slots']:
        count = df[df['可用时间段'].str.contains(slot)].shape[0]
        leaders = df[(df['可用时间段'].str.contains(slot)) & (df['是否组长'] == '是')].shape[0]
        print(f"{slot}: {count}人 (组长{leaders}人)")


generate_test_data()