import random

import pandas as pd


def generate_candidate_insufficient_data():
    groups = ['宣传', '教务', '后勤', '调研']
    time_slots = ["10:00-10:15", "10:15-10:30", "10:30-10:45"]  # 3个时间段需要30人
    candidates = []

    # 只生成25人（比需要少5人）
    for i in range(1, 26):
        group = random.choice(groups)
        candidates.append({
            "成员ID": 2000 + i,
            "报名组别": group,
            "可用时间段": ", ".join(time_slots),
            "是否组长": "是" if random.random() < 0.3 else "否"
        })

    df = pd.DataFrame(candidates)
    df.to_excel("测试用例2_总人数不足.xlsx", index=False)
    print(f"生成测试用例2：总人数 25，需要人数 30")


generate_candidate_insufficient_data()