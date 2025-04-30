import time
import os
import pandas as pd
import random
from DrissionPage import Chromium, ChromiumOptions
from DrissionPage._functions.keys import Keys
from datetime import datetime, timedelta


# 计算时间的函数，我希望23 - 6点不发布笔记，可以按照分钟间隔来增加
def add_minutes_with_idle_time(start_time_str, minutes, idle_start=23, idle_end=6):
    """
    按照指定分钟数增加时间，同时考虑空闲时间（晚上23点到早上6点）
    :param start_time_str: 起始时间，格式为'YYYY-MM-DD HH:MM'
    :param minutes: 要增加的分钟数
    :param idle_start: 空闲时间开始的小时数，默认为23
    :param idle_end: 空闲时间结束的小时数，默认为6
    :return: 增加时间后的时间，格式为'YYYY-MM-DD HH:MM'
    """
    # 将输入的时间字符串转换为datetime对象
    start_time = datetime.strptime(start_time_str, '%Y-%m-%d %H:%M')
    remaining_minutes = minutes

    while remaining_minutes > 0:
        current_hour = start_time.hour
        if idle_start <= current_hour or current_hour < idle_end:
            # 如果当前时间处于空闲时间内
            next_non_idle_time = start_time.replace(hour=idle_end, minute=0, second=0, microsecond=0)
            if current_hour >= idle_start:
                next_non_idle_time += timedelta(days=1)
                # 计算到空闲时间结束还需要多久
            time_to_non_idle = (next_non_idle_time - start_time).total_seconds() / 60
            if remaining_minutes <= time_to_non_idle:
                start_time += timedelta(minutes=remaining_minutes)
                remaining_minutes = 0
            else:
                start_time = next_non_idle_time
                remaining_minutes -= time_to_non_idle
        else:
            # 如果当前时间不在空闲时间内
            time_to_idle = None
            if current_hour < idle_start:
                next_idle_time = start_time.replace(hour=idle_start, minute=0, second=0, microsecond=0)
                time_to_idle = (next_idle_time - start_time).total_seconds() / 60

            if time_to_idle is not None and remaining_minutes > time_to_idle:
                start_time += timedelta(minutes=time_to_idle)
                remaining_minutes -= time_to_idle
            else:
                start_time += timedelta(minutes=remaining_minutes)
                remaining_minutes = 0

                # 将结果转换为字符串格式
    return start_time.strftime('%Y-%m-%d %H:%M')


# 从Excel文件随机读取标签的函数
def read_random_tags_from_excel(excel_path, num_tags=3, sheet_name=0, tag_column='标签'):
    """
    从Excel文件中随机读取指定数量的标签
    :param excel_path: Excel文件路径
    :param num_tags: 要读取的标签数量，默认为3
    :param sheet_name: 工作表名称或索引，默认为第一个工作表
    :param tag_column: 包含标签的列名，默认为'标签'
    :return: 随机选择的标签列表
    """
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
        if tag_column in df.columns:
            # 获取标签列，过滤掉空值，并转换为列表
            tags_raw = df[tag_column].dropna().tolist()

            # 处理标签格式，确保每个标签以#开头
            all_tags = []
            for tag in tags_raw:
                tag = str(tag).strip()
                if not tag.startswith('#'):
                    tag = '#' + tag
                all_tags.append(tag)

                # 如果标签总数少于请求的数量，返回所有标签
            if len(all_tags) <= num_tags:
                print(f"警告：Excel中只有{len(all_tags)}个标签，少于请求的{num_tags}个")
                return all_tags

                # 随机选择指定数量的标签
            selected_tags = random.sample(all_tags, num_tags)
            return selected_tags
        else:
            print(f"警告：在Excel文件中未找到'{tag_column}'列")
            return []
    except Exception as e:
        print(f"读取Excel文件时出错：{e}")
        return []

    # 从Excel文件随机读取内容的函数


def read_random_content_from_excel(excel_path, sheet_name=0, content_column='内容'):
    """
    从Excel文件中随机读取一条笔记内容
    :param excel_path: Excel文件路径
    :param sheet_name: 工作表名称或索引，默认为第一个工作表
    :param content_column: 包含内容的列名，默认为'内容'
    :return: 随机选择的内容字符串，如果读取失败则返回默认内容
    """
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
        if content_column in df.columns:
            # 获取内容列，过滤掉空值，并转换为列表
            content_values = df[content_column].dropna().tolist()
            if content_values:
                # 随机选择一条内容
                return str(random.choice(content_values))
            else:
                print(f"警告：Excel中'{content_column}'列没有内容")
                return get_default_content()
        else:
            print(f"警告：在Excel文件中未找到'{content_column}'列")
            return get_default_content()
    except Exception as e:
        print(f"读取Excel文件内容时出错：{e}")
        return get_default_content()

    # 从Excel文件随机读取标题的函数


def read_random_title_from_excel(excel_path, sheet_name=0, title_column='标题'):
    """
    从Excel文件中随机读取一条笔记标题
    :param excel_path: Excel文件路径
    :param sheet_name: 工作表名称或索引，默认为第一个工作表
    :param title_column: 包含标题的列名，默认为'标题'
    :return: 随机选择的标题字符串，如果读取失败则返回默认标题
    """
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
        if title_column in df.columns:
            # 获取标题列，过滤掉空值，并转换为列表
            title_values = df[title_column].dropna().tolist()
            if title_values:
                # 随机选择一条标题，并确保不超过20个字符（小红书标题限制）
                return str(random.choice(title_values))[:20]
            else:
                print(f"警告：Excel中'{title_column}'列没有内容")
                return get_default_title()
        else:
            print(f"警告：在Excel文件中未找到'{title_column}'列")
            return get_default_title()
    except Exception as e:
        print(f"读取Excel文件标题时出错：{e}")
        return get_default_title()

    # 获取默认内容的函数


def get_default_content():
    """
    当无法从Excel读取内容时，返回默认内容
    :return: 默认内容字符串
    """
    return """📍包住=荒郊野岭？北京求职地理指南：这些地段公司慎重选！通勤3小时警告⏰  
    💰看到"工资面议"快跑！北京hr不会说的真相：面议=低到不敢写！应届生必看防坑手册"""


# 获取默认标题的函数
def get_default_title():
    """
    当无法从Excel读取标题时，返回默认标题
    :return: 默认标题字符串
    """
    return "北京公司避雷指南"


# 获取多个文件夹中的图片路径
def get_images_from_multiple_folders(folder_paths, extensions=('.jpg', '.jpeg', '.png')):
    """
    获取多个文件夹中的所有图片文件路径
    :param folder_paths: 图片文件夹路径列表
    :param extensions: 图片文件扩展名元组
    :return: 每个文件夹中图片路径的字典，格式为 {文件夹路径: [图片路径列表]}
    """
    folder_images = {}

    for folder_path in folder_paths:
        if not os.path.exists(folder_path):
            print(f"警告：文件夹路径 {folder_path} 不存在")
            continue

        images = []
        for file_name in sorted(os.listdir(folder_path)):
            if file_name.lower().endswith(extensions):
                full_path = os.path.join(folder_path, file_name)
                images.append(full_path)

        if images:
            folder_images[folder_path] = images
            print(f"在文件夹 {folder_path} 中找到 {len(images)} 张图片")
        else:
            print(f"警告：文件夹 {folder_path} 中没有找到图片")

    return folder_images


# 从多个文件夹中获取指定索引的图片
def get_images_by_index(folder_images, index, images_per_folder=1):
    """
    从多个文件夹中获取指定索引的图片
    :param folder_images: 每个文件夹中图片路径的字典
    :param index: 当前索引
    :param images_per_folder: 每个文件夹选择的图片数量
    :return: 选择的图片路径列表
    """
    selected_images = []

    for folder, images in folder_images.items():
        if not images:
            continue

        # 计算当前文件夹应该使用的图片索引
        start_idx = (index * images_per_folder) % len(images)

        # 选择指定数量的图片，如果到达文件夹末尾则循环到开头
        for i in range(images_per_folder):
            img_idx = (start_idx + i) % len(images)
            selected_images.append(images[img_idx])

    return selected_images


# 发布笔记的函数
def pub_note(tab, home_url, f_str, title, content_str, tags, g_id, use_time):
    """
    发布小红书笔记
    :param tab: 浏览器窗口中最后一个活动的tab页面
    :param home_url: 小红书创作平台URL
    :param f_str: 拼接号的图片路径
    :param title: 笔记标题
    :param content_str: 笔记内容
    :param tags: 笔记标签
    :param g_id: 商品id，从商品列表可以找到
    :param use_time: 定时时间，一般需要大于当前时间+1小时
    """
    # 点击跳转到创作服务平台
    tab.get(home_url)

    # 等待页面加载
    print("等待页面加载...")
    time.sleep(5)

    # 点击发布笔记
    tab.ele('.btn el-tooltip__trigger el-tooltip__trigger').click()

    #  点击 上传图文
    tab.ele('.creator-tab').click()

    # 文件上传 ，多个图片使用'\n'换行符分割
    tab.ele('@type=file').click.to_upload(f_str)

    # 标题
    tab.ele('@placeholder=填写标题会有更多赞哦～').input(title[:20])

    # 正文输入
    content = tab.ele('@data-placeholder=输入正文描述，真诚有价值的分享予人温暖')
    content.input(content_str[:800])

    # 输入标签，可以跳转的
    time.sleep(1)
    content.focus()
    tab.actions.key_down(Keys.END)
    tab.actions.key_down(Keys.ENTER)
    tab.actions.key_down(Keys.ENTER)

    for tag in tags:
        content.input(tag)
        time.sleep(0.7)
        content.input(Keys.ENTER)
        time.sleep(0.5)
        # tab.actions.key_down(Keys.ENTER)

    # 页面向下滚动，避免元素不可见导致不能点击
    tab.ele('.content').scroll.to_bottom()

    # 是否有商品id
    if g_id:
        tab.ele(
            '.d-button d-button-small d-button-with-content --color-static bold --color-bg-fill-light-opaque d-button-stroke --color-text-paragraph custom-stroke-button').click()
        # tab.ele('.media-commodity').ele('tag:span@@text():添加商品').click()

        tab.ele('@placeholder=搜索商品ID 或 商品名称').input(g_id)
        time.sleep(1)

        tab.ele('.good-card-container').ele('@type=checkbox').click()

        tab.ele(
            '.d-button d-button-default d-button-with-content --color-static bold --color-bg-primary --color-white').click()

        # 是否定时
    if use_time:
        tab.ele('.content').scroll.to_bottom()
        tab.ele('tag:span@@text():定时发布').click()
        tab.ele('@placeholder=选择日期和时间').clear()
        tab.ele('@placeholder=选择日期和时间').input(use_time)

    tab.ele(
        '.d-button d-button-large --size-icon-large --size-text-h6 d-button-with-content --color-static bold --color-bg-fill --color-text-paragraph custom-button red publishBtn').click()


# 主函数
def main():
    # 设置循环次数
    total_cycles = int(input("请输入要循环的任务次数: "))

    # 设置发布间隔（分钟）
    post_interval = int(input("请输入发布间隔（分钟）: "))

    # 设置每个文件夹使用的图片数量
    images_per_folder = int(input("请输入每个文件夹使用的图片数量（默认1张）: ") or "1")

    # 设置图片文件夹路径
    folder_count = int(input("请输入要使用的图片文件夹数量: "))
    folder_paths = []

    for i in range(folder_count):
        folder_path = input(f"请输入第{i + 1}个图片文件夹路径: ").strip()
        while not folder_path or not os.path.exists(folder_path):
            print(f"错误：文件夹路径 {folder_path} 不存在或为空")
            folder_path = input(f"请重新输入第{i + 1}个图片文件夹路径: ").strip()
        folder_paths.append(folder_path)

    # 获取所有文件夹中的图片
    folder_images = get_images_from_multiple_folders(folder_paths)
    if not folder_images:
        print("错误：所有指定文件夹中都没有找到图片，程序退出")
        return

    # Excel文件路径
    excel_path = input("请输入Excel文件路径（例如：D:\\小红书批量创作\\北京\\长尾词\\长尾词.xlsx）: ")
    while not os.path.exists(excel_path):
        print(f"错误：Excel文件路径 {excel_path} 不存在")
        excel_path = input("请重新输入Excel文件路径: ")

    # 商品ID
    g_id = input("请输入商品ID（如果没有请直接回车）: ") or None

    # 创建Chromium选项对象
    edge_options = ChromiumOptions()
    # 设置浏览器类型为Edge
    edge_options.browser_type = 'edge'
    # 设置使用Edge浏览器的二进制文件路径
    edge_options.set_browser_path(r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe')
    # 添加选项，禁用安全特性
    edge_options.set_argument('--disable-web-security')
    edge_options.set_argument('--disable-features=IsolateOrigins,site-per-process')

    # 初始化Chromium对象并传入Edge选项
    driver = Chromium(edge_options)
    tab = driver.new_tab()

    # 小红书创作平台URL
    new_url = "https://creator.xiaohongshu.com/new/home"

    # 当前时间
    current_time = datetime.now()

    # 循环发布笔记
    for i in range(total_cycles):
        print(f"\n开始执行第 {i + 1}/{total_cycles} 次任务")

        # 从多个文件夹中获取指定索引的图片
        selected_images = get_images_by_index(folder_images, i, images_per_folder)

        if not selected_images:
            print("警告：没有获取到图片，跳过本次任务")
            continue

        # 打印将要使用的图片
        print("将使用以下图片：")
        for img_path in selected_images:
            print(img_path)

        # 从Excel随机读取标题
        title = read_random_title_from_excel(excel_path)
        print(f"将使用以下标题：{title}")

        # 从Excel随机读取内容
        content_str = read_random_content_from_excel(excel_path)
        print(f"将使用以下内容：\n{content_str}")

        # 从Excel文件随机读取标签，指定读取5个标签
        num_tags_to_use = 5  # 可以根据需要修改这个数字，控制使用的标签数量
        tags = read_random_tags_from_excel(excel_path, num_tags=num_tags_to_use)

        # 如果没有从Excel读取到标签，使用默认标签
        if not tags:
            tags = ["#标签1", "#标签2"]

        print(f"将使用以下{len(tags)}个标签：")
        for tag in tags:
            print(tag)

            # 计算发布时间
        if i == 0:
            # 第一篇设置为当前时间加1小时
            use_time_dt = current_time + timedelta(hours=1)
            use_time = use_time_dt.strftime('%Y-%m-%d %H:%M:%S')
        else:
            # 后续的笔记按照间隔时间递增，并避开23点到6点
            previous_time_str = use_time.split('.')[0]  # 移除可能的毫秒部分
            # 从previous_time_str中提取日期和时间部分（去掉秒）
            previous_time_parts = previous_time_str.split(":")
            if len(previous_time_parts) >= 2:
                previous_time_no_seconds = ":".join(previous_time_parts[0:2])
                # 计算下一个发布时间
                next_time_str = add_minutes_with_idle_time(previous_time_no_seconds, post_interval)
                if next_time_str:
                    use_time_dt = datetime.strptime(next_time_str, '%Y-%m-%d %H:%M')
                    use_time = use_time_dt.strftime('%Y-%m-%d %H:%M:%S')
                else:
                    # 如果add_minutes_with_idle_time返回None，使用默认值
                    print("警告：计算下一发布时间失败，使用默认时间")
                    use_time_dt = current_time + timedelta(hours=i + 1)
                    use_time = use_time_dt.strftime('%Y-%m-%d %H:%M:%S')
            else:
                # 时间格式不正确，使用默认值
                print("警告：时间格式不正确，使用默认时间")
                use_time_dt = current_time + timedelta(hours=i + 1)
                use_time = use_time_dt.strftime('%Y-%m-%d %H:%M:%S')

        print(f"计划发布时间：{use_time}")

        # 调用函数发布笔记
        try:
            pub_note(tab, new_url, tuple(selected_images), title, content_str, tags, g_id, use_time)
            print(f"第 {i + 1}/{total_cycles} 次任务发布成功！")
        except Exception as e:
            print(f"发布笔记时出错：{e}")
            print("继续下一次任务...")

        # 如果不是最后一次任务，等待一段时间再继续
        if i < total_cycles - 1:
            wait_time = 5  # 每次发布之间等待5秒
            print(f"等待 {wait_time} 秒后继续下一次任务...")
            time.sleep(wait_time)

    # 所有任务完成后提示用户
    print("\n所有任务已安排发布完成！")

    # 询问用户是否关闭浏览器
    close_browser = input("是否关闭浏览器？(y/n): ").strip().lower()
    if close_browser == 'y':
        driver.quit()
        print("浏览器已关闭")
    else:
        print("浏览器保持打开状态")


if __name__ == "__main__":
    main()
