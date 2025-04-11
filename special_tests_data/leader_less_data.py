import pandas as pd
import random
from datetime import datetime, timedelta


def generate_leader_insufficient_data():
    groups = ['宣传', '教务', '后勤', '调研']
    time_slots = ["10:00-10:15", "14:00-14:15", "15:30-15:45"]  # 3个时间段需要3个组长
    candidates = []

    # 只生成2个组长（比时间段数少1）
    leader_count = 0
    for i in range(1, 31):  # 总人数刚好3组×10=30人
        group = random.choice(groups)
        is_leader = False

        # 严格控制组长数量
        if leader_count < 2:  # 只允许2个组长
            is_leader = random.choice([True] + [False] * 9)  # 10%概率

        if is_leader:
            leader_count += 1

        candidates.append({
            "成员ID": 1000 + i,
            "报名组别": group,
            "可用时间段": ", ".join(time_slots),  # 所有人都可以参加所有时间段
            "是否组长": "是" if is_leader else "否"
        })

    df = pd.DataFrame(candidates)
    df.to_excel("测试用例1_组长不足.xlsx", index=False)
    print(f"生成测试用例1：组长数 {leader_count}，时间段数 3")


generate_leader_insufficient_data()