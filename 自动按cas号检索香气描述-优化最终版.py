import requests
from bs4 import BeautifulSoup
from deep_translator import GoogleTranslator
import pandas as pd
import time
from tqdm import tqdm
import os


# 进行网络请求时，捕获可能的SSL错误并自动重试
def make_request_with_retry(url, data=None, retries=3, delay=5):
    for attempt in range(retries):
        try:
            if data:
                response = requests.post(url, data=data)
            else:
                response = requests.get(url)

            response.raise_for_status()  # 如果响应码不是200会抛出异常
            return response
        except requests.exceptions.SSLError as ssl_error:
            print(f"SSL错误发生，正在尝试第 {attempt + 1} 次重试...")
            time.sleep(delay)  # 延迟后重新尝试
        except requests.exceptions.RequestException as e:
            print(f"请求错误: {e}")
            return None
    print("多次尝试后仍无法连接，跳过该请求。")
    return None


def search_cas_odor(cas_number):
    # 使用POST请求并将CAS号填入qName字段
    url = 'http://www.perflavory.com/search.php'
    data = {'qName': cas_number}

    # 调用带重试机制的请求函数
    response = make_request_with_retry(url, data)

    if not response:
        return {
            'CAS 编号': cas_number,
            '化合物名称 (英文)': "请求失败",
            '化合物名称 (中文)': "请求失败",
            '香气描述 (中文)': "请求失败",
            '香气描述 (英文)': "请求失败"
        }

    # 解析HTML内容
    soup = BeautifulSoup(response.text, 'html.parser')

    # 检查是否包含“抱歉，您的搜索：...返回零结果”
    if soup.find(string=lambda text: text and f"抱歉，您的搜索：“{cas_number}”返回零结果" in text):
        return {
            'CAS 编号': cas_number,
            '化合物名称 (英文)': "请求失败",
            '化合物名称 (中文)': "请求失败",
            '香气描述 (中文)': "请求失败",
            '香气描述 (英文)': "请求失败"
        }

    # 查找所有包含CAS号的标签，class为lstw10的span标签
    cas_tags = soup.find_all('span', class_='lstw10')

    for cas_tag in cas_tags:
        # 提取检索结果中的CAS号
        retrieved_cas = cas_tag.get_text(strip=True)

        # 如果检索结果中的CAS号与输入的CAS号匹配
        if retrieved_cas == cas_number:
            # 获取相邻的香气描述标签，class为lstw11的span标签
            odor_info = cas_tag.find_next('span', class_='lstw11')

            # 获取化合物英文名称的<a>标签
            compound_name_tag = soup.find('a', onclick=True)
            compound_name = compound_name_tag.text.strip() if compound_name_tag else "请求失败"
            translated_compound_name = GoogleTranslator(source='en', target='zh-CN').translate(
                compound_name) if compound_name_tag else "请求失败"

            if odor_info:
                # 获取香气描述的英文内容
                odor_description = odor_info.get_text(strip=True)
                translated_odor = GoogleTranslator(source='en', target='zh-CN').translate(odor_description)

                return {
                    'CAS 编号': cas_number,
                    '化合物名称 (英文)': compound_name,
                    '化合物名称 (中文)': translated_compound_name,
                    '香气描述 (中文)': translated_odor,
                    '香气描述 (英文)': odor_description
                }
            else:
                # 如果没有香气描述英文，仍然输出CAS编号和化合物名称（英文及中文翻译）
                return {
                    'CAS 编号': cas_number,
                    '化合物名称 (英文)': compound_name,
                    '化合物名称 (中文)': translated_compound_name,
                    '香气描述 (中文)': "请求失败",
                    '香气描述 (英文)': "请求失败"
                }

    return {
        'CAS 编号': cas_number,
        '化合物名称 (英文)': "请求失败",
        '化合物名称 (中文)': "请求失败",
        '香气描述 (中文)': "请求失败",
        '香气描述 (英文)': "请求失败"
    }


# 动态获取Excel文件路径
input_file = input("请输入Excel文件的路径（例如：C:\\Users\\ymx20\\Desktop\\化合物数据.xlsx）：").strip()

# 打开文件并获取列名
df = pd.read_excel(input_file)

# 显示所有列名，并让用户选择需要查询的列
print(f"Excel 文件中包含的列为：{list(df.columns)}")
column_to_query = input("请输入要爬取数据的列名称（例如：'CAS 编号'）：")

if column_to_query not in df.columns:
    print(f"列 {column_to_query} 不存在，请检查列名。")
    exit(1)

# 确定输出文件路径
output_dir = os.path.dirname(input_file)

# 获取输入文件的文件名（不带扩展名）
file_name_without_extension = os.path.splitext(os.path.basename(input_file))[0]

# 修改输出文件名称
output_file = os.path.join(output_dir, f"{file_name_without_extension}_香气描述爬虫.xlsx")

# 获取CAS编号列数据
total_cas_numbers = len(df[column_to_query])
results = []

# 循环遍历CAS号并显示进度条
for index, cas_number in tqdm(enumerate(df[column_to_query]), total=total_cas_numbers, desc="进度", unit="CAS"):
    print(f"进度: {((index + 1) / total_cas_numbers) * 100:.2f}% 完成")
    print(f"正在处理 {index + 1}/{total_cas_numbers}：{cas_number}")

    # 获取查询结果
    result = search_cas_odor(cas_number)

    # 打印查询结果中的详细信息
    print(f"CAS 编号: {result['CAS 编号']}")
    print(f"化合物名称 (英文): {result['化合物名称 (英文)']}")
    print(f"化合物名称 (中文): {result['化合物名称 (中文)']}")
    print(f"香气描述 (英文): {result['香气描述 (英文)']}")
    print(f"香气描述 (中文): {result['香气描述 (中文)']}")

    # 将查询结果添加到结果列表
    results.append(result)

    # 请求间隔5秒
    time.sleep(5)

# 将结果保存为新的Excel文件，按指定列顺序排列
output_df = pd.DataFrame(results, columns=[
    'CAS 编号', '化合物名称 (英文)', '化合物名称 (中文)', '香气描述 (中文)', '香气描述 (英文)'
])
output_df.to_excel(output_file, index=False)
print("已成功保存到:", output_file)
