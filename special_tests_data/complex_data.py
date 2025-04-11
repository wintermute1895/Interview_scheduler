import random

import pandas as pd


def generate_complex_case():
    groups = ['宣传', '教务', '后勤', '调研']
    candidates = []

    # 生成临界值数据
    total_groups = 5  # 需要5组×10=50人
    leader_count = 4  # 比组数少1
    problem_slot = "10:00-10:15"

    # 生成刚好49人（少1人）
    for i in range(1, 50):
        group = '调研' if i < 3 else random.choice(groups)  # 调研组仅有2人
        is_leader = False

        # 控制组长数量
        if leader_count > 0 and random.random() < 0.5:
            is_leader = True
            leader_count -= 1

        # 控制问题时间段人数
        if i < 9:  # 问题时间段只有8人
            available = [problem_slot, "14:00-14:15"]
        else:
            available = ["14:00-14:15", "15:00-15:15"]

        candidates.append({
            "成员ID": 5000 + i,
            "报名组别": group,
            "可用时间段": ", ".join(available),
            "是否组长": "是" if is_leader else "否"
        })

    df = pd.DataFrame(candidates)
    df.to_excel("测试用例5_综合问题.xlsx", index=False)
    print("生成测试用例5：综合所有问题")


generate_complex_case()