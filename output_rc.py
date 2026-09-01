import win32com.client
import os
import time

# ================= 配置 =================
SC_FILE = r"F:\bo1\paper\iridium\iridium_stk\iridium.sc"
OUTPUT_TXT = r".\\output\\access_intervals.txt"
USE_EXISTING_STK = False

# 指定要额外加入的对象名称
SPECIAL_NAMES = {"1-G357","2--G1591","3-G1579","4--G339","5-G501","6-G309","7-G339","8-G696","9-G979","10-G689","BeijingXi", "ZhengzhouDong"}

# 指定这些名字可能对应的对象类型（卫星为常用的一般类型，所以将satellite除外）
#此处列车groundvehi和地面站beijingxi和zhengzhoudong的类型由collect_objects输出得到
SPECIAL_CLASSES = {"GroundVehicle", "Place"}
# ======================================

def get_dataset_values(result, dataset_name, fallback_index=None):

    try:
        return result.DataSets.GetDataSetByName(dataset_name).GetValues()
    except Exception:
        if fallback_index is not None:
            return result.DataSets.Item(fallback_index).GetValues()
        raise

def load_scenario(root, sc_file):
    """
    加载场景
    """
    print(f"打开场景文件：{sc_file}")

    try:
        if root.CurrentScenario is not None:
            print("关闭当前场景...")
            root.CloseScenario()
            time.sleep(2)
    except Exception:
        pass

    try:
        root.LoadScenario(sc_file)
        time.sleep(2)
    except Exception as e:
        print(f"LoadScenario 失败：{e}")
        print("尝试命令方式加载场景...")
        root.ExecuteCommand(f'Load / "{sc_file}"')
        time.sleep(2)

    scenario = root.CurrentScenario
    if scenario is None:
        raise RuntimeError("场景加载失败，CurrentScenario 为空。")

    return scenario

def collect_objects(scenario):
    """
    收集对象：
    1. 所有 Satellite
    2. 名称在 SPECIAL_NAMES 中，且类型在 SPECIAL_CLASSES 中的对象
    """
    objects = []

    print("\n场景一级对象列表：")
    for i in range(scenario.Children.Count):
        obj = scenario.Children.Item(i)
        print(f"   - ClassName={obj.ClassName}, InstanceName={obj.InstanceName}")

        if obj.ClassName == "Satellite":
            objects.append(obj)
        elif obj.ClassName in SPECIAL_CLASSES and obj.InstanceName in SPECIAL_NAMES:
            objects.append(obj)

    return objects

def main():
    # 输出目录
    os.makedirs(os.path.dirname(os.path.abspath(OUTPUT_TXT)), exist_ok=True)

    # 场景文件检查
    sc_file = os.path.abspath(SC_FILE)
    if not os.path.isfile(sc_file):
        raise FileNotFoundError(f"场景文件不存在：{sc_file}")

    print("连接 STK11 桌面应用...")

    # 启动或连接 STK
    try:
        if USE_EXISTING_STK:
            stk_app = win32com.client.GetActiveObject("STK11.Application")
            print("已连接到正在运行的 STK11")
        else:
            raise Exception("强制启动新实例")
    except Exception:
        print("启动新的 STK11 实例...")
        stk_app = win32com.client.DispatchEx("STK11.Application")
        time.sleep(8)
        print("STK11 已启动")

    stk_app.Visible = True
    stk_app.UserControl = True
    root = stk_app.Personality2

    # 加载场景
    scenario = load_scenario(root, sc_file)
    print(f"场景已加载：{scenario.InstanceName}")

    # 收集对象
    objects = collect_objects(scenario)

    print(f"\n 参与计算对象数：{len(objects)}")
    for obj in objects:
        print(f"   - {obj.ClassName}: {obj.InstanceName}")

    if len(objects) < 2:
        raise RuntimeError("对象数量不足 2 个，无法计算连接关系。")

    # 计算 access
    pair_count = 0
    interval_count = 0

    with open(OUTPUT_TXT, "w", encoding="utf-8") as f:
        f.write(f"Scenario: {scenario.InstanceName}\n")
        f.write(f"Time Range: {scenario.StartTime} -> {scenario.StopTime}\n")
        f.write("=" * 100 + "\n\n")

        for i in range(len(objects)):
            obj1 = objects[i]
            for j in range(i + 1, len(objects)):
                obj2 = objects[j]

                type1 = obj1.ClassName
                name1 = obj1.InstanceName
                type2 = obj2.ClassName
                name2 = obj2.InstanceName

                pair_count += 1
                print(f"->计算: {type1}/{name1} <-> {type2}/{name2}")

                f.write(f"{type1}/{name1} <-> {type2}/{name2}\n")

                try:
                    access = obj1.GetAccessToObject(obj2)
                    access.ComputeAccess()

                    dp = access.DataProviders.Item("Access Data")
                    result = dp.Exec(scenario.StartTime, scenario.StopTime)

                    start_times = get_dataset_values(result, "Start Time", 0)
                    stop_times = get_dataset_values(result, "Stop Time", 1)

                    if len(start_times) == 0:
                        f.write("  No Access\n\n")
                        continue

                    for k, (st, et) in enumerate(zip(start_times, stop_times), start=1):
                        f.write(f"  Access {k}: Start = {st}, Stop = {et}\n")
                        interval_count += 1

                    f.write("\n")

                except Exception as e:
                    f.write(f"  Error: {e}\n\n")
                    print(f"   失败: {e}")

    print("\n 全部完成！")
    print(f" 输出文件：{os.path.abspath(OUTPUT_TXT)}")
    print(f" 对象对数量：{pair_count}")
    print(f" 连接区间数量：{interval_count}")

if __name__ == "__main__":
    main()