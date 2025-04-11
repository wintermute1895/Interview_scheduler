import datetime
import random
import re

import pandas as pd


def generate_time_slots():
    time_slots = []
    start = datetime.datetime.strptime("10:00", "%H:%M")
    end = datetime.datetime.strptime("20:00", "%H:%M")
    delta = datetime.timedelta(minutes=15)
    current = start

    while current <= end:
        current_time = current.time()
        if (current_time >= datetime.time(10, 0) and current_time < datetime.time(12, 0)) or \
                (current_time >= datetime.time(14, 0) and current_time <= datetime.time(20, 0)):
            time_slots.append(current.strftime("%H:%M"))
        current += delta

    formatted = []
    for i in range(len(time_slots) - 1):
        formatted.append(f"{time_slots[i]}-{time_slots[i + 1]}")
    return formatted


def load_candidates(file_path):
    df = pd.read_excel(file_path)
    candidates = []

    for _, row in df.iterrows():
        time_str = str(row['可用时间段']).strip()
        available_times = [t.strip() for t in time_str.split(',')]

        is_leader = str(row['是否组长']).strip() in ['是', 'True', 'Yes', 'Y', '1']

        candidates.append({
            'id': row['成员ID'],
            'group': row['报名组别'],
            'available_times': available_times,
            'is_leader': is_leader,
            'scheduled': False  # 添加安排状态标记
        })

    # 数据校验
    valid_groups = {'宣传', '教务', '后勤', '调研'}
    time_pattern = re.compile(r"\d{2}:\d{2}-\d{2}:\d{2}")

    for c in candidates:
        if c['group'] not in valid_groups:
            raise ValueError(f"无效组别：{c['group']} (成员ID: {c['id']})")
        for t in c['available_times']:
            if not time_pattern.match(t):
                raise ValueError(f"无效时间段格式：{t} (成员ID: {c['id']})")

    return candidates


class InterviewScheduler:
    def __init__(self, candidates):
        self.candidates = candidates
        self.time_slots = generate_time_slots()
        self.schedule = []
        self.debug_info = {
            'time_slot_status': {},
            'unassigned': [],
            'canceled_slots': [],
            'special_groups': []
        }
        self.groups = ['宣传', '教务', '后勤', '调研']

    def schedule_interviews(self):
        # 创建可修改的副本
        candidates_pool = [c.copy() for c in self.candidates]

        # 按时间段处理
        for slot in self.time_slots:
            # 获取本时段可用且未安排的候选人
            slot_candidates = [c for c in candidates_pool
                               if not c['scheduled']
                               and slot in c['available_times']]

            # 记录时间段状态
            slot_status = {
                'total': len(slot_candidates),
                'scheduled': 0,
                'has_leader': False,
                'missing_groups': [],
                'is_canceled': False
            }

            # 决定是否安排本时间段
            if len(slot_candidates) < 8:
                self.debug_info['canceled_slots'].append(slot)
                slot_status['is_canceled'] = True
                self.debug_info['time_slot_status'][slot] = slot_status
                continue

            # 组建面试组（8-12人）
            group_size = min(12, len(slot_candidates))
            group, leader_flag, missing_groups = self._form_group(slot_candidates, group_size)

            # 更新成员安排状态
            for member in group:
                member['scheduled'] = True
                # 从池中移除已安排成员
                candidates_pool = [c for c in candidates_pool if c['id'] != member['id']]

            # 记录组信息
            self.schedule.append({
                '时间段': slot,
                '成员': group,
                'has_leader': leader_flag,
                'missing_groups': missing_groups
            })

            # 更新调试信息
            slot_status.update({
                'scheduled': len(group),
                'has_leader': leader_flag,
                'missing_groups': missing_groups
            })
            self.debug_info['time_slot_status'][slot] = slot_status

            if not leader_flag:
                self.debug_info['special_groups'].append(f"{slot}：无组长")
            if missing_groups:
                self.debug_info['special_groups'].append(f"{slot}：缺少组别 {', '.join(missing_groups)}")

        # 处理剩余人员
        self._handle_remaining(candidates_pool)

        return self._format_output(), self.debug_info

    def _form_group(self, candidates, target_size):
        temp_candidates = [c.copy() for c in candidates if not c['scheduled']]
        group = []
        has_leader = False
        required_groups = set(self.groups)

        # 优先选择组长
        leaders = [c for c in temp_candidates if c['is_leader']]
        if leaders:
            leader = random.choice(leaders)
            group.append(leader)
            temp_candidates.remove(leader)
            has_leader = True
            required_groups.discard(leader['group'])

        # 补充其他成员
        while len(group) < target_size and temp_candidates:
            # 优先补充缺少的组别
            candidate = None
            for g in required_groups:
                candidates_in_group = [c for c in temp_candidates if c['group'] == g]
                if candidates_in_group:
                    candidate = random.choice(candidates_in_group)
                    required_groups.discard(g)
                    break

            # 没有需要补充的组别则随机选择
            if not candidate:
                candidate = random.choice(temp_candidates)

            group.append(candidate)
            temp_candidates.remove(candidate)

        # 检查缺少的组别（允许缺少1个）
        present_groups = {m['group'] for m in group}
        missing = set(self.groups) - present_groups
        allowed_missing = list(missing)[:1] if len(missing) > 1 else list(missing)

        return group[:target_size], has_leader, allowed_missing

    def _handle_remaining(self, remaining_candidates):
        # 尝试将剩余人员插入已有小组
        for candidate in remaining_candidates:
            if candidate['scheduled']:
                continue  # 安全校验

            # 查找可加入的时间段
            for group in self.schedule:
                slot = group['时间段']
                if (len(group['成员']) < 12
                        and slot in candidate['available_times']
                        and not candidate['scheduled']):
                    group['成员'].append(candidate)
                    candidate['scheduled'] = True
                    break

        # 记录最终未安排人员
        self.debug_info['unassigned'] = [c['id'] for c in remaining_candidates if not c['scheduled']]

    def _format_output(self):
        output = []
        member_ids = set()  # 用于重复检查

        for group in self.schedule:
            for member in group['成员']:
                # 重复性校验
                if member['id'] in member_ids:
                    raise ValueError(f"成员 {member['id']} 被重复安排")
                member_ids.add(member['id'])

                output.append({
                    '时间段': group['时间段'],
                    '成员ID': member['id'],
                    '报名组别': member['group'],
                    '是否组长': '是' if member['is_leader'] else '否'
                })

        return pd.DataFrame(output)

    def generate_report(self):
        report = [
            "=== 面试安排报告 ===",
            f"总候选人数量: {len(self.candidates)}",
            f"已安排人数: {len(self.candidates) - len(self.debug_info['unassigned'])}",
            f"未安排人数: {len(self.debug_info['unassigned'])}",
            "\n=== 时间段安排详情 ==="
        ]

        for slot in self.time_slots:
            status = self.debug_info['time_slot_status'].get(slot, {})
            if status.get('is_canceled', False):
                report.append(f"{slot}：取消安排（可用人数：{status['total']}）")
            else:
                info = [
                    f"{slot}：",
                    f"  - 安排人数：{status.get('scheduled', 0)}",
                    f"  - 包含组长：{'是' if status.get('has_leader', False) else '否'}",
                    f"  - 缺少组别：{', '.join(status.get('missing_groups', [])) or '无'}"
                ]
                report.append("\n".join(info))

        report.append("\n=== 特殊情况汇总 ===")
        report.append(f"取消时间段总数：{len(self.debug_info['canceled_slots'])}")
        report.append(f"特殊小组数量：{len(self.debug_info['special_groups'])}")
        if self.debug_info['special_groups']:
            report.extend(["具体特殊情况："] + self.debug_info['special_groups'])

        report.append("\n=== 未安排人员 ===")
        if self.debug_info['unassigned']:
            report.append(", ".join(map(str, self.debug_info['unassigned'])))
        else:
            report.append("所有候选人均已安排")

        return "\n\n".join(report)

if __name__ == "__main__":
    try:
        candidates = load_candidates("综合测试数据120人.xlsx")
    except Exception as e:
        print(f"数据加载失败：{str(e)}")
        exit(1)

    scheduler = InterviewScheduler(candidates)

    try:
        schedule_df, debug_info = scheduler.schedule_interviews()
    except ValueError as e:
        print(f"安排错误：{str(e)}")
        exit(2)

    # 保存结果
    schedule_df.to_excel("五期面试安排表_test2.xlsx", index=False)
    with open("面试安排报告.txt", "w", encoding="utf-8") as f:
        f.write(scheduler.generate_report())

    print("安排完成，结果已保存")