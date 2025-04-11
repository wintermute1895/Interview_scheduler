import random
import pandas as pd
from datetime import datetime, timedelta


def generate_random_time_slots():
    """生成随机可用的时间段（每人至少2个）"""
    all_slots = []
    current = datetime.strptime("10:00", "%H:%M")
    end = datetime.strptime("20:00", "%H:%M")

    # 生成所有可能时间段（跳过午休）
    while current < end:
        if not (12 <= current.hour < 14):  # 跳过12:00-14:00
            end_time = current + timedelta(minutes=15)
            all_slots.append(f"{current.strftime('%H:%M')}-{end_time.strftime('%H:%M')}")
        current += timedelta(minutes=15)

    # 每人随机选择2-4个时间段
    return random.sample(all_slots, k=random.randint(2, 4))


def generate_random_data(num_candidates=200):
    """生成随机报名数据"""
    groups = ['宣传', '教务', '后勤', '调研']
    data = []

    for i in range(1, num_candidates + 1):
        # 随机组别（加权随机，使各组人数不均）
        group = random.choices(
            groups,
            weights=[0.3, 0.25, 0.25, 0.2],  # 宣传组稍多
            k=1
        )[0]

        # 随机是否组长（约25%概率）
        is_leader = random.random() < 0.25

        # 生成时间段
        time_slots = generate_random_time_slots()

        data.append({
            "成员ID": 1000 + i,
            "报名组别": group,
            "可用时间段": ", ".join(time_slots),
            "是否组长": "是" if is_leader else "否"
        })

    return pd.DataFrame(data)


# 生成数据
df = generate_random_data()

# 保存Excel（自动调整列宽）
with pd.ExcelWriter("随机报名数据.xlsx", engine='xlsxwriter') as writer:
    df.to_excel(writer, index=False)

    # 自动调整列宽
    worksheet = writer.sheets['Sheet1']
    for i, col in enumerate(df.columns):
        max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
        worksheet.set_column(i, i, max_len)

print("已生成 随机报名数据.xlsx")
print("数据概览：")
print(f"总人数：{len(df)}")
print("各组人数：")
print(df['报名组别'].value_counts())
print("\n组长人数：")
print(df['是否组长'].value_counts())