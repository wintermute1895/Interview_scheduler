import random

import pandas as pd


def generate_timeslot_insufficient_data():
    groups = ['宣传', '教务', '后勤', '调研']
    special_slot = "10:00-10:15"  # 问题时间段

    candidates = []
    # 生成足够的总人数（50人）
    for i in range(1, 51):
        group = random.choice(groups)
        # 只有前8人可以参加问题时间段
        if i <= 8:
            available = [special_slot, "14:00-14:15"]
        else:
            available = ["14:00-14:15", "15:00-15:15"]

        candidates.append({
            "成员ID": 4000 + i,
            "报名组别": group,
            "可用时间段": ", ".join(available),
            "是否组长": "是" if random.random() < 0.3 else "否"
        })

    df = pd.DataFrame(candidates)
    df.to_excel("测试用例4_时间段人数不足.xlsx", index=False)
    print(f"生成测试用例4：{special_slot}时段仅有8人可用")


generate_timeslot_insufficient_data()