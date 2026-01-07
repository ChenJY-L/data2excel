import xlwings as xw
import csv
from pathlib import Path


class Temp2Data:
    dat_columns = [
        '时间',  # 👈 新增：时间列
        'Default1','Default2','Default3','Default4','Default5','Default6','Default7','Default8',
        'TC1控制温度','TC1实际温度','TC1实际输出电压','TC1实际输出电流','TC1实际输出功率',
        'TC2控制温度','TC2实际温度','TC2实际输出电压','TC2实际输出电流','TC2实际输出功率',
        '压力(电压值)','压力(g)','步进电机当前位置','温度','湿度',
        'LD设定温度(度)','LD实际温度(度)','LD设定温度(脉冲值)',
        'LD实际温度(脉冲值)','LD PWM(功率)'
    ]

    def __init__(self, out_dir: str, excel_files: list[str]):
        self.out_dir = Path(out_dir)
        self.files = excel_files
        self.xwapp = xw.App(visible=False, add_book=False)
        self.xwapp.display_alerts = False
        self.xwapp.screen_updating = False

    def _read_excel(self, file):
        """读取Excel：奇数列为时间，偶数列为温度"""
        wb = self.xwapp.books.open(file)
        sht = wb.sheets[0]
        data = sht.range(sht.used_range).value
        wb.close()

        rows = data[1:] if isinstance(data[0][0], str) else data  # 去掉表头行（如有）

        result = []
        for row in rows:
            for i in range(0, len(row), 2):  # 奇数列时间，偶数列温度
                t_val = row[i]
                temp_val = row[i + 1] if i + 1 < len(row) else None
                if t_val is None or temp_val is None:
                    continue
                try:
                    temp = float(temp_val)
                    result.append((t_val, temp))
                except Exception:
                    continue
        return result

    def align_and_generate(self):
        """时间对齐并生成 dat 文件"""
        data1 = self._read_excel(self.files[0])
        data2 = self._read_excel(self.files[1])

        # 转成 dict 方便查找
        dict1 = dict(data1)
        dict2 = dict(data2)

        # 找交集（时间完全一致）
        common_times = sorted(set(dict1.keys()).intersection(set(dict2.keys())))
        if not common_times:
            raise ValueError("❌ 未找到可对齐的时间点")

        aligned = []
        for t in common_times:
            v1, v2 = dict1[t], dict2[t]
            room_temp = min(v1, v2)
            skin_temp = max(v1, v2)
            aligned.append((t, room_temp, skin_temp))

        # 输出 .dat 文件
        out_path = self.out_dir / "温度合并.dat"
        with open(out_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter='\t')
            writer.writerow(self.dat_columns)
            for t, room, skin in aligned:
                row = [0]*len(self.dat_columns)
                row[0] = t  # 👈 第一列写入 Excel 原始时间数值
                row[self.dat_columns.index('Default2')] = room
                row[self.dat_columns.index('Default3')] = room
                row[self.dat_columns.index('TC1实际温度')] = skin
                writer.writerow(row)

        print(f"✅ 已生成文件: {out_path}")


if __name__ == '__main__':
    files = ["C:/Users/Dell/Desktop/20251204-空腹-75M24-小臂皮包骨-3MPET胶-cyh/温度数据.xlsx",
             "C:/Users/Dell/Desktop/20251204-空腹-75M24-小臂皮包骨-3MPET胶-cyh/皮肤温度.xlsx",]
    out_dir = '../'
    t2d = Temp2Data(out_dir, files)
    t2d.align_and_generate()
