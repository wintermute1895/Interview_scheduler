import random

import pandas as pd


def generate_group_insufficient_data():
    groups = ['宣传', '教务', '后勤', '调研']
    time_slots = ["10:00-10:15", "14:00-14:15"]
    candidates = []

    # 特别设置调研组只有3人
    group_dist = {
        '宣传': 15,
        '教务': 15,
        '后勤': 15,
        '调研': 3  # 明显不足
    }

    for group, count in group_dist.items():
        for i in range(count):
            candidates.append({
                "成员ID": 3000 + len(candidates) + 1,
                "报名组别": group,
                "可用时间段": ", ".join(time_slots),
                "是否组长": "是" if random.random() < 0.3 else "否"
            })

    df = pd.DataFrame(candidates)
    df.to_excel("测试用例3_单组人数不足.xlsx", index=False)
    print("生成测试用例3：调研组仅有3人")


generate_group_insufficient_data()