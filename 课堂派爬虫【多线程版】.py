# ###############################################################
# #######
# #                       课堂派文件下载【多线程版】
# #                        作者：Ignorant
# #                 使用声明：本程序仅用于学习交流，请勿用于商业用途。
# #######
# ###############################################################
#                           修订版本
# ###############################################################
#
# 5.4   改进：添加back、retry的指令（利用嵌套函数实现的循环）
#       修复：修复了重复创建output文件夹的bug（该问题会引起360报毒查杀）
#       修复：修复了重复运行重新时未初始化参数导致页码重复的bug
#
# 5.3   改进：提供exit的输入指令，可以退出函数【但是提供back指令以达到返回上一级函数失败了，python没有goto语法】
#       改进：中文对齐函数通过传递多个元组参数，实现一次实现多次居中对齐的功能
#
# 5.2   改进：重命名部分变量和函数的名称，使其更切合其定义；
#       改进：中英文混杂的字符串的居中对齐输出，可实现快速修改分隔符
#       修复：修复打包后程序出现错误自动退出而不提示错误信息的bug,
#       修复：修复打包后，正常退出报错：NameError: name ‘exit‘ is not defined的bug
#
# 5.0   改进：弃用GPT的代码建议，采用CSDN的代码（见参考资料），提高代码的优雅性和效率
#       功能：添加时间统计的功能
#       内测：【同一文件（370页）同一网络（300Mbps）的下载速度对比：单线程版：35+min > 多线程（GPT版）：3+min > 多线程(CSDN版)：2-min】
#
# 4.4   功能：实现程序内输入解析网址【抛弃urllib.parse中的urlencode()函数，自己重写，同时构建其逆函数，避免网址的参数变化后需要重构字典】
#       改进：预计时间采取1位小数
#       内测：【目前可下载pdf、docx文件；pptx文件不行】
#
# 4.2   改进：动态线程分配，根据页码数自动分配线程
#       改进：重新利用pyinstaller打包程序（因为发送给舍友运行不了）【单线程报毒，多线程则正常】
#
# 4.0   功能：实现多线程下载【为防止对目标网站造成过大的压力，最高16线程】
#       分叉：分离单线程版本和多线程版本，单线程版本逐渐废弃
#
# 3.0   重构：代码重新优化
#       修复：修复时间计算的bug、合成pdf页码混乱等bug
#       改进：使用cmd脚本作为启动器【杀毒软件会对pyinstaller打包的程序报毒】
#
# 2.2   功能：增加检查是否存在重复图片的功能
#       修复：修复合并全部图片的bug
#
# 2.0   功能：增加合并图片为PDF的功能
#
# 1.2   功能：增加进度提示的功能，估算剩余时间
#       改进：整理代码结构，拆分代码为多个功能函数（模块化编程）
#
# 1.0   实现最基础的爬取功能，没有花里胡巧的纠错功能
#       历程：定位请求url->模拟请求url->提取json内容，在浏览器打开图片->识别json内容，爬取单张图片->爬取多张图片->保存到指定的文件夹
#
# ###############################################################
#                           参考资料
# ###############################################################
#
# 1.    [课堂派资料 PDF 文件下载_课堂派不允许下载的 word 怎么下载-CSDN 博客](https://blog.csdn.net/sucr233/article/details/114501594)
# 2.    [Python 异步加载XHR数据抓取——GET与POST请求方式对比](https://blog.csdn.net/qq_17249717/article/details/84326555)
# 3.    [python几行代码，把图片转换、合并为PDF文档](https://blog.csdn.net/snrxian/article/details/108916764)
# 4.    [Python全局变量和局部变量（超详细，纯干货，保姆级教学）](https://blog.csdn.net/Kristen_jiang/article/details/129780846)
# 5.    [Python多线程详解_python 多线程-CSDN博客](https://blog.csdn.net/ifhuke/article/details/128619653)
# 6.    其他：知乎{全角空格输出\u3000}
# 7.    GPT：在当前目录下合并图片为PDF【因为给出的代码存在PDF页码混乱的bug，所以魔改：利用for循环为合并指定范围内的图片】
#           提供的多线程代码【太拉跨了，导致文件下载是一块（多线程数同时）进行的，当线程同时结束才进行下一次【代码臭长，不如CSDN的优雅高效】
#
# ###############################################################
# 导入库
import sys
from concurrent.futures import ThreadPoolExecutor
from os import mkdir, chdir, getcwd, startfile
from os.path import exists as path_exists
from time import time

from PIL import Image
from requests import get as request_get

"""
全局变量信息
"""
# 默认的待解析的xhr请求网址、自定义的待解析的xhr请求网址、xhr请求网址的前缀、网址的参数部分（字典形式表示）、需要爬取图片的网址前缀
default_url_needToAnalyze = "https://document.ketangpai.com/PW/GetPage?f=RDpcXGZpbGVjYWNoZTIwMjNcXGtldGFuZ3BhaS10ZXN0Lm9zcy1jbi1oYW5nemhvdS1pbnRlcm5hbC5hbGl5dW5jcy5jb20uODBcNjVkYjQ1OWIyYWU0MS5wZGY=&img=EnvigofzSy2a1z7QH2udMI776DeYeSt1FKS4b5RR3INpxzB_hCHH3YbO_T*vdWBEs276jiqR8pw-&isMobile=false&hd=&readLimit=&sn=1&furl=&srv=0&revision=-1&comment=-1"
custom_url_needToAnalyze = ""
xhr_base_url = "https://document.ketangpai.com/PW/GetPage?"
param_dict = dict()
picture_base_url = "https://document.ketangpai.com/img?img="
# 用户信息
user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36 Edg/123.0.0.0'
request_headers = {'User-Agent': user_agent}
# 时间估计(时间估算时钟，爬虫开始时钟，合并开始时钟)，过程时间（爬取经历时间，合并经历时间，总耗时）对应（0，1，2）
estimate_clock, spider_clock, combine_clock = 0, 0, 0
time_statistics_list = [0, 0, 0]
# 最大线程数，多线程下载文件计数器
max_threads, downloadCounter = 1, 0
# 开始页码、结束页码、文件总页数、爬取文件页数
pageStart, pageEnd, pageCount, page_count = 1, 1, 1, 1
# 已完成页数、剩余页数、进度百分比、预计等待时间
finished_pages, rest_pages, percent, wait_time = 0, 0, "【0.00%】", "正在评估中"
# 指定图片的集合
image_list = list()
# 段落分隔符
sign = '='
paragraph = sign * 64

"""
初始化参数
"""


# 程序再次运行时的重新初始化参数
def init_variables():
    global estimate_clock, spider_clock, combine_clock
    estimate_clock, spider_clock, combine_clock = 0, 0, 0
    global time_statistics_list, max_threads, downloadCounter
    time_statistics_list = [0, 0, 0]
    max_threads, downloadCounter = 1, 0
    global pageStart, pageEnd, pageCount, page_count
    pageStart, pageEnd, pageCount, page_count = 1, 1, 1, 1
    global finished_pages, rest_pages, percent, wait_time
    finished_pages, rest_pages, percent, wait_time = 0, 0, "【0.00%】", "正在评估中"
    global image_list
    image_list = list()


"""
打印提示信息
"""


# 全角字符转化【GPT的代码提示】
def to_full_width(text):
    full_width_text = ''
    for char in text:
        char_code = ord(char)
        # 如果是半角空格，则转换为全角空格
        if char_code == 32:
            full_width_text += chr(12288)
        # 如果是半角字符，根据Unicode编码的规律进行转换
        elif 33 <= char_code <= 126:
            full_width_text += chr(char_code + 65248)
        else:
            full_width_text += char
    return full_width_text


# 多层：文本居中对齐
def multi_center_align(text="", *args):
    for arg in args:
        text = text.center(arg[0], arg[1])
    return text


# 打印信息
def print_info():
    # 统一行长度、含文本时sign数量
    p_length = 40
    sign_count = 1 * 2
    # 需要打印的信息
    product_info = "课堂派文件下载【多线程版】"
    writer_info = "作者：Ignorant"
    use_info = "使用声明：本程序仅用于学习交流，请勿用于商业用途。"
    paragraph_a = sign * p_length
    paragraph_b = " " * (p_length - sign_count * 3)
    # 文本居中对齐
    product_info = multi_center_align(product_info, (p_length - sign_count, ' '), (p_length, '='))
    writer_info = multi_center_align(writer_info, (p_length - sign_count, ' '), (p_length, '='))
    use_info = multi_center_align(use_info, (p_length - sign_count, ' '), (p_length, '='))
    paragraph_a = multi_center_align(paragraph_a, (p_length, sign))
    paragraph_b = multi_center_align(paragraph_b, (p_length, sign))
    # 转化为全角字符
    product_info = to_full_width(product_info)
    writer_info = to_full_width(writer_info)
    use_info = to_full_width(use_info)
    paragraph_a = to_full_width(paragraph_a)
    paragraph_b = to_full_width(paragraph_b)
    # 打印信息
    print("\n" + paragraph_a)
    print(paragraph_b)
    print(product_info)
    print(writer_info)
    print(use_info)
    print(paragraph_b)
    print(paragraph_a + "\n")
    print("\n" + paragraph + "\n")


"""
文件夹操作
"""


# 创建输出文件夹
def create_output_dir():
    if not getcwd().endswith("output"):
        if not path_exists("output"):
            mkdir("output")
        chdir("output")


# 打开输出文件夹
def open_output_dir():
    file_path = getcwd()
    startfile(file_path)


'''
程序的暂停和退出
'''


# 确认退出
def press_to_pause():
    print("\n" + paragraph + "\n")
    order = input("\n提示：\t按Enter键确认退出;\n\t输入retry重新运行程序\n\t")
    print("\n" + paragraph + "\n")
    if order == "retry":
        init_variables()
        main()


# 退出程序
def exit_program():
    sys.exit()


"""
进度条
"""


# 时间转换

def time_convert(elapsed_time):
    elapsed_time_h = elapsed_time // 3600
    elapsed_time_min = (elapsed_time - elapsed_time_h * 3600) // 60
    elapsed_time_s = elapsed_time - elapsed_time_h * 3600 - elapsed_time_min * 60
    elapsed_time_h = int(elapsed_time_h)
    elapsed_time_min = int(elapsed_time_min)
    wait_time_str = ""
    if elapsed_time_h != 0:
        wait_time_str += f"{elapsed_time_h}小时"
        wait_time_str += f"{elapsed_time_min}分钟"
    else:
        if elapsed_time_min != 0:
            wait_time_str += f"{elapsed_time_min}分钟"
            wait_time_str += "{:.1f}秒".format(elapsed_time_s)
        else:
            wait_time_str += "{:.1f}秒".format(elapsed_time_s)
    return wait_time_str


# 进度估计
def estimate_progress(index):
    global estimate_clock, rest_pages, percent, wait_time, finished_pages
    # 完成页数
    finished_pages = index - pageStart
    # 剩余页数
    rest_pages = pageEnd - index + 1
    # 进度百分比
    percent = "【" + str(round(finished_pages / page_count * 100, 2)) + "%】"
    # 时间差
    elapsed_time = time() - estimate_clock
    # 预计等待时间
    if finished_pages != 0:
        wait_time = time_convert(elapsed_time / finished_pages * rest_pages)


# 时间统计
def time_statistics():
    print("\n" + paragraph + "\n")
    print("时间统计：")
    print(
        f"\t爬取用时：{time_statistics_list[0]}\n\t合并用时：{time_statistics_list[1]}\n\t总用时：{time_statistics_list[2]}")
    print("\n" + paragraph + "\n")


"""
检查（文件是否重复\图片是否重复、页码是否合法）
"""


# 存在同名文件的处理方法
def check_file_repeat():
    if path_exists("output(" + str(pageStart) + "~" + str(pageEnd) + ").pdf"):
        print("\t警告：已存在同名文件，默认退出程序！")
        exit_program()


# 检查图片是否重复
def check_picture_repeat(page_num):
    global estimate_clock
    picture_name = str(page_num) + ".png"
    if path_exists(picture_name):
        print(f"\r\t警告：跳过第{page_num}页，原因：图片已存在！")
        # 重新记录时间（否则时间存在偏差）
        estimate_clock = time()
        return True
    else:
        return False


# 检查页码合法性
def check_page_range():
    # pageCount 为总页数，page_count 为需要下载页数
    global pageStart, pageEnd, pageCount, page_count, param_dict
    # 检查页码范围是否合法
    print("\n" + paragraph + "\n")
    print("正在解析文件...")
    param_dict['img'] = ''
    data = request_get(request_url(1), headers=request_headers).json()
    pageCount = data.get('PageCount', None)
    if pageEnd == "all":
        pageEnd = pageCount
    if pageStart < 1:
        print(f"\t警告：起始页码不规范，重置起始页码为最小页码1！")
        pageStart = 1
    if pageStart > pageCount:
        pageStart = 1
        print(f"\t警告：起始页码大于结束页码，重置起始页码为最小页码1！")
    if pageEnd > pageCount:
        pageEnd = pageCount
        print(f"\t警告：结束页码超出范围，重置结束页码为最大页码{pageCount}！")
    print(f"正在获取页码信息：共{pageEnd - pageStart + 1}/{pageCount}页（第{pageStart}页~第{pageEnd}页）...")
    # 初始化页码相关信息
    page_count = pageEnd - pageStart + 1
    init_merge_variables()


"""
线程设置
"""


# 设置最大线程数
def threads_settings():
    global max_threads
    pages_level = [100, 50, 25, 10, 5, 1]
    thread_list = [16, 12, 8, 4, 2, 1]
    for index in range(len(thread_list)):
        if page_count >= pages_level[index]:
            max_threads = thread_list[index]
            print(f"正在设置线程数：{max_threads}...")
            break


# 多线程爬虫【参考自CSDN】
def spiders_multi_threads():
    task = []  # 存放线程
    with ThreadPoolExecutor(max_threads) as pool:
        for page_num in range(pageStart, pageEnd + 1):
            task.append(pool.submit(crawl_picture, page_num))


"""
xhr请求网址的解析和构建
"""


# 获取文件的解析网址，可选参数【是否返回到函数get_pages_range()】
def get_file_url():
    global custom_url_needToAnalyze
    order = input("\n提示：是否改变爬取文件：【输入0/1/exit，代表否/是/退出】（默认文件解析地址可能会失效）\n\t")
    if order == '1':
        custom_url_needToAnalyze = input("\n提示：请输入目标文件的解析网址：\n")
    elif order == '0':
        custom_url_needToAnalyze = default_url_needToAnalyze
        print(f"目标文件的解析网址:\n{custom_url_needToAnalyze}\n")
    elif order == 'exit':
        exit_program()
    else:
        print(f"\t警告：错误的命令，请重新输入...\n")
        get_file_url()
        return 0


# url拆解
def url_split(src_url):
    global param_dict, xhr_base_url
    try:
        char_count = src_url.count('&')
        # 提取网址的基址和参数
        xhr_base_url = src_url.split("?")[0]
        params_str_list = src_url.split("?")[1].split("&")
        # 提取参数到字典
        for index in range(char_count + 1):
            dict_key = params_str_list[index].split("=")[0]
            dict_value = params_str_list[index].split(dict_key + "=")[1]
            param_dict[dict_key] = dict_value
    except:
        print("\t警告：URL解析失败，请检查输入的网址是否正确...\n")
        get_file_url()
        url_split(custom_url_needToAnalyze)


# url合并
def url_joint(dictionary):
    target_url = xhr_base_url + "?"
    for key in dictionary:
        target_url = target_url + str(key) + "=" + str(dictionary[key]) + "&"
    target_url = target_url[:-1]
    return target_url


# 构建XHR请求的网址
def request_url(page_num):
    global param_dict
    param_dict['sn'] = page_num - 1
    target_url = url_joint(param_dict)
    return target_url


"""
爬取文件
"""


# 格式化输入信息(将输入的num1和num1间的间隔符替换为',')
def input_format(input_str):
    new_str = input_str
    for char in input_str:
        if not char.isdigit():
            new_str = new_str.replace(char, ',')
            break
    return new_str


#

# 输入起始和结束页码
def get_pages_range():
    global pageStart, pageEnd, custom_url_needToAnalyze
    input_str = input("\n提示：请输入页码范围【例如：1或2,3或all(全部页码)或back(返回上一步)或exit（退出）】:\n\t")
    try:
        # 格式化输入内容
        input_format_str = input_format(input_str)
        pageStart, pageEnd = eval(input_format_str)
        # 检查文件是否重复
        check_file_repeat()
    except Exception:
        if input_str == "all":
            pageStart, pageEnd = 1, "all"
        elif input_str.isnumeric():
            pageStart, pageEnd = int(input_str), int(input_str)
        elif input_str == "back":
            get_file_url()
            url_split(custom_url_needToAnalyze)
            get_pages_range()
            return 0
        elif input_str == "exit":
            exit_program()
        else:
            print("\t警告：输入页码格式错误...\n")
            get_pages_range()
            return 0
        # 检查文件是否重复
        check_file_repeat()


# 爬取多页图片
def crawl_pictures():
    global estimate_clock, spider_clock
    # 输入解析地址
    get_file_url()
    # 解析文件地址
    url_split(custom_url_needToAnalyze)
    # 输入页码范围
    get_pages_range()
    # 检查页码合法性
    check_page_range()
    # 时间预估时钟、爬虫时钟开始计时
    estimate_clock, spider_clock = time(), time()
    # 设置最大线程数量
    threads_settings()
    # 打印爬虫开始前的信息
    print("\n" + paragraph + "\n")
    print("正在爬取图片，请耐心等待...")
    # 开始多线程爬虫
    spiders_multi_threads()
    # 统计爬虫时间：
    global time_statistics_list
    time_statistics_list[0] = time_convert(time() - spider_clock)
    # 打印爬虫完成后的信息
    print(
        f"\r\t进度：正在爬取第{pageEnd}/{pageCount}页，已完成{page_count}/{page_count}页【100.00%】，预计等待时间：0.0秒...",
        end="")
    print(f"\n\t爬取完成！\n输出文件夹为：\n\t{getcwd()}")


# 爬取单页图片
def crawl_picture(page_num):
    global downloadCounter
    if not check_picture_repeat(page_num):
        # 正常情况下执行：
        try:
            # 打印进度
            print(
                f"\r\t进度：正在爬取第{downloadCounter}/{pageCount}页，已完成{finished_pages}/{page_count}页{percent}，预计等待时间：{wait_time}...",
                end="")
            # 发起请求并获取响应内容
            data = request_get(request_url(page_num), headers=request_headers).json()
            # 保存图片
            save_picture(data)
            # 进度评估、统计下载的图片数量
            downloadCounter += 1
            estimate_progress(downloadCounter + pageStart)
        # 出现异常情况执行：
        except Exception as e:
            print(f"错误：跳过第{page_num}页，原因：{e}...")


# 保存图片
def save_picture(data):
    # 构建图片的网址
    picture_url = picture_base_url + data.get('NextPage', None)
    # 解析图片的页码并命名图片
    picture_index = data.get('PageIndex', None)
    picture_name = str(picture_index) + ".png"
    # 下载图片
    picture_data = request_get(picture_url).content
    # 保存图片
    with open(picture_name, 'wb') as picture:
        picture.write(picture_data)


"""
合并图片为PDF
"""


# 获取当前文件夹中的指定图片
def get_image_files():
    for page_num in range(pageStart, pageEnd + 1):
        # 转化为图片对象,并转化为RGB模式
        img = (Image.open(str(page_num) + '.png')).convert('RGB')
        # 输出进度
        print(
            f"\r\t进度：正在合并第{page_num}/{pageCount}页，已完成{finished_pages}/{page_count}页{percent}，预计等待时间：{wait_time}...",
            end="")
        # 将图片加入列表
        image_list.append(img)
        # 进度评估
        estimate_progress(page_num)


# 将图片列表中的图片合并为pdf
def image_list_to_pdf(pdf_name):
    if image_list:
        image_list[0].save(pdf_name, save_all=True, append_images=image_list[1:])


# 合并图片为PDF
def image_to_pdf():
    global estimate_clock, combine_clock, finished_pages, percent, rest_pages, wait_time
    # 时间预估时钟、合并时钟开始计时
    estimate_clock, combine_clock = time(), time()
    # 重置页码信息
    init_merge_variables()
    # PDF文件的名称
    pdf_name = f"output({pageStart}~{pageEnd}).pdf"
    # 打印合并开始前的信息
    print("\n" + paragraph + "\n")
    print("正在合并中，请耐心等待...")
    # 获取当前文件夹中的指定图片
    get_image_files()
    # 将图片列表保存为PDF文件
    image_list_to_pdf(pdf_name)
    # 统计爬虫时间：
    global time_statistics_list
    time_statistics_list[1] = time_convert(time() - combine_clock)
    time_statistics_list[2] = time_convert(time() - spider_clock)
    # 打印合并完成后的信息
    print(
        f"\r\t进度：正在合并第{pageEnd}/{pageCount}页，已完成{page_count}/{page_count}页【100.00%】，预计等待时间：0.0秒...",
        end="")
    print(f"\n\t合并完成：\nPDF文件保存为: \n\t{pdf_name}")


# 重新初始化合并pdf文件的参数
def init_merge_variables():
    # 已完成页数、剩余页数、进度百分比、预计等待时间
    global finished_pages, rest_pages, percent, wait_time
    finished_pages, rest_pages, percent, wait_time = 0, page_count, "【0.00%】", "正在评估中"


# 主函数
def main():
    try:
        print_info()  # 打印程序信息
        create_output_dir()  # 创建输出文件夹
        crawl_pictures()  # 爬取图片
        image_to_pdf()  # 合并PDF
        time_statistics()  # 输出时间统计信息
        open_output_dir()  # 打开输出文件夹
    except Exception as error_reasons:
        # 异常结束程序：
        print("\n" + paragraph + "\n")
        print(f"错误：程序异常结束!!!\n\t原因：{error_reasons}")
    else:
        # 正常结束程序：
        print("程序正常结束...".center(64, " "))
    finally:
        # 运行完成按下任意键后退出（用于pyinstaller打包后的程序，避免自动退出）
        press_to_pause()

# 主程序
if __name__ == '__main__':
    main()
