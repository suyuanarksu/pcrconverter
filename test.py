from sys import argv, stdin
from pandas import read_excel, to_numeric
from os.path import basename, splitext

def load_axis_data(excel_file):
    """加载轴模板数据"""
    try:
        # 读取轴模板数据，从第9行开始作为表头
        df_axis = read_excel(excel_file, sheet_name='轴模板', header=13)  # 第14行当作标题，即第15行当作正文
        df_axis = df_axis.iloc[:, :4]  # 取所有行的前4列数据 .iloc[行数, 列数]
        df_axis.columns = ['帧数', '秒数', '角色', '操作']

        # 过滤有效数据行（帧数不为空且为数字）
        valid_rows = df_axis.dropna(subset=['帧数'])
        valid_rows = valid_rows[to_numeric(valid_rows['帧数'], errors='coerce').notna()]

        return valid_rows[['帧数', '秒数', '角色', '操作']]
    except Exception as e:
        print(f"Error loading axis data: {e}")
        return None

def load_chara_list(excel_file):
    try:
        df_chara = read_excel(excel_file, sheet_name='基础数据', header=1)
        df_chara = df_chara.iloc[0:5, 1:2]
        chara_list = df_chara['角色名字'].tolist()
        return chara_list
    except Exception as e:
        print(f"Error loading chara data: {e}")
        return None

def lframe_to_rframe(excel_file, chara_list, chara_frames):
    try:
        df_tp = read_excel(excel_file, sheet_name='TP变化', header=2)
        df_tp = df_tp.iloc[:, :19]  # 取所有行的前19列数据 .iloc[行数, 列数]
        df_tp.columns = ['逻辑帧', '渲染帧', 'col2', 'col3', 'col4', 'tp0', '原因0', 'col7', 'tp1', '原因1', 'col10', 'tp2',
                         '原因2', 'col13', 'tp3', '原因3', 'col16', 'tp4', '原因4']
        tp_maps = {}
        for i in range(5):
            valid_row = df_tp[['逻辑帧', '渲染帧', f'tp{i}', f'原因{i}']]
            valid_row = valid_row.dropna(subset=[f'tp{i}'])
            tp_map = {}
            for _, row in valid_row.iterrows():
                lframe = int(row["逻辑帧"])
                rframe = int(row["渲染帧"])
                tp = int(row[f'tp{i}'])
                reason = row[f'原因{i}']
                if lframe in chara_frames[chara_list[i]] and reason == "放UB":
                    tp_map[lframe] = rframe
            tp_maps[chara_list[i]] = tp_map
        stoprow = df_tp[['渲染帧']]
        stopframe = 0
        for _, row in stoprow.iterrows():
            rframe = int(row["渲染帧"])
            if rframe >= stopframe:
                stopframe = rframe
        tp_maps["暂停"] = stopframe
        return tp_maps
    except Exception as e:
        print(f"Error loading chara data: {e}")
        return None

excel_file = argv[1]
#excel_file = "D5-海中二水兔水莉莉水优妮圣千-16955w.xlsx"
# "D3-611火猫星猫克总els-15084w.xlsx"
filename = basename(excel_file)
file, extension = splitext(filename)
axis_data = load_axis_data(excel_file)
chara_list = load_chara_list(excel_file)
update_records = """
更新日志:
v25.12.28:
1.新增自动识别模拟器功能

v25.12.24:
1.修复了mzq2.4.0及以上版本导出的excel中假auto操作识别错误的问题
"""
# 脚本头部
script_content = """from autotimeline import *
import sys
sys.path.append('.')
import ctypes, os
from ctypes import wintypes

"""
pos_name = ["暂停", "SET", "AUTO", "SPEED", chara_list[0], chara_list[1], chara_list[2], chara_list[3], chara_list[4]]
script_content += f"pos_name = {pos_name}\n"
script_content += """pos_dnplayer = ([0.94,0.95,0.95,0.95,0.74,0.62,0.50,0.38,0.26],[0.05,0.64,0.76,0.90,0.80,0.80,0.80,0.80,0.80])
pos_mumu = ([0.95,0.36,0.24,0.10,0.20,0.20,0.20,0.20,0.20],pos_dnplayer[0])
GetModuleFileNameExA = ctypes.windll.psapi.GetModuleFileNameExA

def _get_name(h):
    buf = ctypes.create_string_buffer(1024)
    GetModuleFileNameExA(wintypes.HANDLE(h), None, buf, 1024)
    return os.path.basename(buf.value.decode(errors="ignore"))

def detect(h):
    n = _get_name(h)
    if "MuMu" in n: return "MUMU", pos_mumu
    if "Ld9Box" in n or "LdBox" in n: return "雷电", pos_dnplayer

handle_name, (pos_x, pos_y) = detect(Program.hwnd)
print("当前模拟器" + handle_name)
"""

script_content += """
print("minitouch 连接中")
minitouch.connect("127.0.0.1", 1111)
max_x = minitouch.getMaxX()
max_y = minitouch.getMaxY()
"""
script_content += f"print(\"正在运行: {file}\")"

script_content += """
for i, name in enumerate(pos_name):
	print(f"{name} 定位中")
	minitouch.setPos(name, int(max_x * pos_x[i]), int(max_y * pos_y[i]))

print(f"解除暂停，塔塔开!")

autopcr.setOffset(2, 0); # offset calibration
"""

grouped_operations = {}  # 空字典
chara_frames = {}
preframe = 0
for _, row in axis_data.iterrows():
    frame = int(row['帧数'])
    time = row['秒数']
    character = row['角色']
    operation = row['操作']
    grouped_operation = {}
    if frame > preframe:
        grouped_operations[frame] = []
        grouped_operation[character] = [time, operation]
        grouped_operations[frame].append(grouped_operation)  # 添加以逻辑帧为键，角色和操作为值的元素
        preframe = frame
    elif frame == preframe:
        grouped_operation[character] = [time, operation]
        grouped_operations[frame].append(grouped_operation)  # 添加以逻辑帧为键，角色和操作为值的元素
    if character in chara_list:
        if character not in chara_frames:
            chara_frames[character] = []
            chara_frames[character].append(frame)
        else:
            chara_frames[character].append(frame)
sorted_lframes = sorted(grouped_operations.keys())  # 逻辑帧排序
tp_maps = lframe_to_rframe(excel_file, chara_list, chara_frames)
print(f"请选择战斗速度（1/2/4）")
while True:
    try:
        user_input = int(stdin.readline().rstrip())
    except ValueError:
        print("输入有误，请重新输入")
        continue
    if user_input in [1, 2, 4]:
        break
    elif user_input == "v":
        print(update_records)
        continue
    else:
        print("输入有误，请重新输入")
if user_input == 2:
    script_content += f"autopcr.waitFrame(autopcr.getFrame() + 60); minitouch.press(\"SPEED\") # 加速\n"
elif user_input == 4:
    script_content += f"autopcr.waitFrame(autopcr.getFrame() + 60); minitouch.press(\"SPEED\") # 2倍速\n"
    script_content += f"autopcr.waitFrame(autopcr.getFrame() + 120); minitouch.press(\"SPEED\") # 2倍速\n"
for lframe in sorted_lframes:
    operations = grouped_operations[lframe]
    # 分析操作类型
    for i in range(len(operations)):
        chara = (list(operations[i].keys()))[0]
        if chara not in chara_list:
            script_content += f"# {chara} time {operations[i][chara][0]}\n"
        else:
            if user_input == 1:
                if tp_maps[chara][lframe]:
                    if operations[i][chara][1] == '连点':
                        script_content += f"autopcr.waitFrame({tp_maps[chara][lframe]} - {45}); minitouch.press(\"{chara}\") # 连点 time {operations[i][chara][0]}\n"
                    elif operations[i][chara][1] == 'AUTO':
                        script_content += f"autopcr.waitFrame({tp_maps[chara][lframe]} - {45}); minitouch.press(\"AUTO\") # AUTO开\n"
                        script_content += f"# AUTO time {operations[i][chara][0]}\n"
                        script_content += f"autopcr.waitFrame({tp_maps[chara][lframe]} + {5}); minitouch.press(\"AUTO\") # AUTO关\n"
                    elif '容错' in operations[i][chara][1]:
                        script_content += f"autopcr.waitFrame({tp_maps[chara][lframe]} - {5}); minitouch.press(\"AUTO\") # AUTO开\n"
                        script_content += f"# AUTO time {operations[i][chara][0]}\n"
                        script_content += f"autopcr.waitFrame({tp_maps[chara][lframe]} + {5}); minitouch.press(\"AUTO\") # AUTO关\n"
                    else:
                        script_content += f"autopcr.waitFrame({tp_maps[chara][lframe]}); minitouch.press(\"{chara}\") # lframe {lframe}//time {operations[i][chara][0]}\n"
            if user_input == 2:
                if tp_maps[chara][lframe]:
                    if operations[i][chara][1] == '连点':
                        script_content += f"autopcr.waitFrame({tp_maps[chara][lframe]} - {45}); minitouch.press(\"{chara}\") # 连点 time {operations[i][chara][0]}\n"
                    elif operations[i][chara][1] == 'AUTO':
                        script_content += f"autopcr.waitFrame({tp_maps[chara][lframe]} - {45}); minitouch.press(\"AUTO\") # AUTO开\n"
                        script_content += f"# AUTO time {operations[i][chara][0]}\n"
                        script_content += f"autopcr.waitFrame({tp_maps[chara][lframe]} + {5}); minitouch.press(\"AUTO\") # AUTO关\n"
                    elif '容错' in operations[i][chara][1]:
                        script_content += f"autopcr.waitFrame({tp_maps[chara][lframe]} - {60}); minitouch.press(\"SPEED\") # 减速\n"
                        script_content += f"autopcr.waitFrame({tp_maps[chara][lframe]} - {5}); minitouch.press(\"AUTO\") # AUTO开\n"
                        script_content += f"# AUTO time {operations[i][chara][0]}\n"
                        script_content += f"autopcr.waitFrame({tp_maps[chara][lframe]} + {5}); minitouch.press(\"AUTO\") # AUTO关\n"
                        script_content += f"autopcr.waitFrame({tp_maps[chara][lframe]} + {30}); minitouch.press(\"SPEED\") # 加速\n"
                    else:
                        script_content += f"autopcr.waitFrame({tp_maps[chara][lframe]} - {60}); minitouch.press(\"SPEED\") # 减速\n"
                        script_content += f"autopcr.waitFrame({tp_maps[chara][lframe]}); minitouch.press(\"{chara}\") # lframe {lframe}//time {operations[i][chara][0]}\n"
                        script_content += f"autopcr.waitFrame({tp_maps[chara][lframe]} + {30}); minitouch.press(\"SPEED\") # 加速\n"
            if user_input == 4:
                if tp_maps[chara][lframe]:
                    if operations[i][chara][1] == '连点':
                        script_content += f"autopcr.waitFrame({tp_maps[chara][lframe]} - {60}); minitouch.press(\"{chara}\") # 连点 time {operations[i][chara][0]}\n"
                    elif operations[i][chara][1] == 'AUTO':
                        script_content += f"autopcr.waitFrame({tp_maps[chara][lframe]} - {60}); minitouch.press(\"AUTO\") # AUTO开\n"
                        script_content += f"# AUTO time {operations[i][chara][0]}\n"
                        script_content += f"autopcr.waitFrame({tp_maps[chara][lframe]} + {5}); minitouch.press(\"AUTO\") # AUTO关\n"
                    elif '容错' in operations[i][chara][1]:
                        script_content += f"autopcr.waitFrame({tp_maps[chara][lframe]} - {60}); minitouch.press(\"SPEED\") # 1倍速\n"
                        script_content += f"autopcr.waitFrame({tp_maps[chara][lframe]} - {5}); minitouch.press(\"AUTO\") # AUTO开\n"
                        script_content += f"# AUTO time {operations[i][chara][0]}\n"
                        script_content += f"autopcr.waitFrame({tp_maps[chara][lframe]} + {5}); minitouch.press(\"AUTO\") # AUTO关\n"
                        script_content += f"autopcr.waitFrame({tp_maps[chara][lframe]} + {30}); minitouch.press(\"SPEED\") # 2倍速\n"
                        script_content += f"autopcr.waitFrame({tp_maps[chara][lframe]} + {60}); minitouch.press(\"SPEED\") # 4倍速\n"
                    else:
                        script_content += f"autopcr.waitFrame({tp_maps[chara][lframe]} - {60}); minitouch.press(\"SPEED\") # 1倍速\n"
                        script_content += f"autopcr.waitFrame({tp_maps[chara][lframe]}); minitouch.press(\"{chara}\") # lframe {lframe}//time {operations[i][chara][0]}\n"
                        script_content += f"autopcr.waitFrame({tp_maps[chara][lframe]} + {30}); minitouch.press(\"SPEED\") # 2倍速\n"
                        script_content += f"autopcr.waitFrame({tp_maps[chara][lframe]} + {60}); minitouch.press(\"SPEED\") # 4倍速\n"
stopframe = int(tp_maps["暂停"])
script_content += f"autopcr.waitFrame({stopframe} - 30); minitouch.press(\"暂停\") # 暂停\n"

output_file = f"{file}-{user_input}倍速.py"
with open(output_file, 'w', encoding='utf-8') as f:
    f.write(script_content)
print(f'已生成 {file}-{user_input}倍速.py')
