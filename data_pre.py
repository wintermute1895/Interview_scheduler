import pandas as pd

# 定义替换规则字典（键为原字符串，值为替换后的字符串）
replace_rules = {
    r'4月18日（周五） 20:30-22:00(,|$)': r'1,2,3,4,5\1',
    r'4月19日（周六） 9:00-11:00(,|$)': r'6,7,8,9,10,11\1',
    r'4月19日（周六） 14:00-16:00(,|$)': r'12,13,14,15,16,17\1',
    r'4月19日（周六） 16:00-18:00(,|$)': r'18,19,20,21,22,23\1'
}


def clean_time_columns(df):
    """清洗时间列数据"""
    # 处理可用时间段列（D列）
    df['可用时间段'] = df['可用时间段'].astype(str)
    for pattern, replacement in replace_rules.items():
        df['可用时间段'] = df['可用时间段'].str.replace(pattern, replacement, regex=True)

    # 清理多余的逗号和空值
    df['可用时间段'] = df['可用时间段'].str.replace(r',+', ',', regex=True)
    df['可用时间段'] = df['可用时间段'].str.strip(',')
    df['可用时间段'] = df['可用时间段'].replace({'--': '', 'nan': ''})
    return df


# 读取原始数据
raw_df = pd.read_excel("医路赞歌五期一面时间统计_预处理1(4_18零点).xlsx", sheet_name=0)

# 清洗数据
cleaned_df = clean_time_columns(raw_df)

# 保存新文件
with pd.ExcelWriter("卫津路面试数据(已清洗).xlsx") as writer:
    cleaned_df.to_excel(writer, index=False, sheet_name='详细数据统计（按提交顺序排序）')

    # 保留原统计分析表
    stats_df = pd.read_excel("医路赞歌五期一面时间统计_预处理1(4_18零点).xlsx", sheet_name=1)
    stats_df.to_excel(writer, index=False, sheet_name='统计分析')

print("数据清洗完成，已保存为：卫津路面试数据(已清洗).xlsx")