import os
import re

import pandas as pd


def extract_chinese_names(root_path):
    # 定义中文姓名正则模式（匹配开头2-4个汉字）
    pattern = re.compile(r'^[\u4e00-\u9fa5]{2,4}')
    names = set()

    try:
        # 递归遍历所有子目录
        for foldername, subfolders, filenames in os.walk(root_path):
            for filename in filenames:
                # 匹配文件名开头的姓名
                match = pattern.match(filename)
                if match:
                    names.add(match.group())

        # 处理结果
        if names:
            df = pd.DataFrame(sorted(names), columns=["姓名"])
            output_path = os.path.join(root_path, "姓名列表.xlsx")
            df.to_excel(output_path, index=False)
            print(f"成功生成文件：{output_path}")
            print(f"唯一姓名数量：{len(names)}")
            print(f"扫描范围包含：{sum(len(files) for _, _, files in os.walk(root_path))}个文件")
        else:
            print("未检测到符合要求的姓名")

    except Exception as e:
        print(f"程序异常: {str(e)}")

if __name__ == "__main__":
    path = r"F:\医路赞歌五期报名表"
    extract_chinese_names(path)