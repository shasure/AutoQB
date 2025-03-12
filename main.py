import sys

import yaml
from playwright.sync_api import sync_playwright
from datetime import datetime, timedelta
import time
import random
import subprocess
from urllib.parse import urlparse

import pandas as pd


def csv_to_html(csv_file_path, html_file_path):
    """
    将 CSV 文件内容转换为具有翻页功能的 HTML 表格并保存为 HTML 文件。
    支持选择每页展示 10、20、30、50、100 条记录，link 列将显示为超链接。
    页面下方添加“上一页”和“下一页”按钮。

    :param csv_file_path: 输入的 CSV 文件路径
    :param html_file_path: 输出的 HTML 文件路径
    """
    # 读取 CSV 文件
    df = pd.read_csv(csv_file_path)

    # 将 link 列转换为超链接
    df['link'] = df['link'].apply(lambda x: f'<a href="{x}" target="_blank">{x}</a>' if pd.notna(x) else x)

    # 将 DataFrame 转换为 HTML 表格，不包含索引
    html_table = df.to_html(na_rep='nan', escape=False)

    # 生成完整的 HTML 页面，包含翻页功能和每页记录数选择
    html_content = f"""
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>搜索结果</title>
    <style>
        table {{
            border-collapse: collapse;
            width: 100%;
        }}
        th, td {{
            border: 1px solid #ddd;
            padding: 8px;
            text-align: left;
        }}
        th {{
            background-color: #f2f2f2;
        }}
    </style>
</head>
<body>
    <h1>搜索结果</h1>
    <label for="pageSize">每页记录数: </label>
    <select id="pageSize" onchange="changePageSize()">
        <option value="10">10</option>
        <option value="20">20</option>
        <option value="30">30</option>
        <option value="50" selected>50</option>
        <option value="100">100</option>
    </select>
    <div id="table-container">
        {html_table}
    </div>
    <div id="pagination">
        <button id="prevPage" onclick="prevPage()" disabled>上一页</button>
        <button id="nextPage" onclick="nextPage()">下一页</button>
    </div>

    <script>
        const table = document.querySelector('table');
        const tableContainer = document.getElementById('table-container');
        const pagination = document.getElementById('pagination');
        const pageSizeSelect = document.getElementById('pageSize');
        let currentPage = 1;
        let pageSize = parseInt(pageSizeSelect.value);

        function changePageSize() {{
            pageSize = parseInt(pageSizeSelect.value);
            currentPage = 1;
            renderTable();
        }}

        function renderTable() {{
            const rows = table.rows;
            const totalPages = Math.ceil((rows.length - 1) / pageSize);
            const startIndex = (currentPage - 1) * pageSize + 1;
            const endIndex = Math.min(startIndex + pageSize, rows.length);

            for (let i = 1; i < rows.length; i++) {{
                if (i >= startIndex && i < endIndex) {{
                    rows[i].style.display = 'table-row';
                }} else {{
                    rows[i].style.display = 'none';
                }}
            }}

            renderPagination(totalPages);
        }}

        function renderPagination(totalPages) {{
            const prevPageButton = document.getElementById('prevPage');
            const nextPageButton = document.getElementById('nextPage');

            prevPageButton.disabled = currentPage === 1;
            nextPageButton.disabled = currentPage === totalPages;
        }}

        function prevPage() {{
            if (currentPage > 1) {{
                currentPage--;
                renderTable();
            }}
        }}

        function nextPage() {{
            const rows = table.rows;
            const totalPages = Math.ceil((rows.length - 1) / pageSize);
            if (currentPage < totalPages) {{
                currentPage++;
                renderTable();
            }}
        }}

        renderTable();
    </script>
</body>
</html>
    """
    # 保存 HTML 文件
    with open(html_file_path, 'w', encoding='utf-8') as file:
        file.write(html_content)
    print(f"CSV 文件已成功转换为 HTML 并保存到 {html_file_path}")

def is_domain_blacklisted(link, blacklist):
    """
    检查链接的域名是否在黑名单中。

    :param link: 输入的链接
    :param blacklist: 黑名单域名列表
    :return: 如果域名在黑名单中返回 True，否则返回 False
    """
    try:
        parsed_url = urlparse(link)
        domain = parsed_url.netloc
        for black_domain in blacklist:
            if domain.endswith(black_domain):
                return True
    except ValueError:
        return False

def main():
    # 读取配置文件
    with open('config.yaml', 'r', encoding='utf-8') as file:
        config = yaml.safe_load(file)
    print("成功读取配置文件")

    # 读取 blacklist.txt 文件中的域名
    blacklist = []
    with open(config['blacklist_file'], 'r', encoding='utf-8') as file:
        blacklist = file.read().splitlines()

    # 读取 query.txt 文件中的关键词
    with open(config['query_file'], 'r', encoding='utf-8') as file:
        keywords = file.read().splitlines()
    print(f"从 {config['query_file']} 中读取到 {len(keywords)} 个关键词")

    # 获取当前日期
    current_date = datetime.now()
    # 根据 days 变量计算起始日期
    days = config.get('days', 0)
    start_date = (current_date - timedelta(days=days)).strftime('%m/%d/%Y')
    end_date = current_date.strftime('%m/%d/%Y')
    print(f"查询时间段：从 {start_date} 到 {end_date}")

    # open Chrome browser
    chrome_path = r'Google Chrome'
    debugging_port = r"--remote-debugging-port=9999"
    try:
        # 启动 Chrome 浏览器
        if sys.platform.startswith('darwin'):  # macOS
            subprocess.run(['open', '-a', chrome_path, '--args', debugging_port], check=True)
        elif sys.platform.startswith('win'):  # Windows
            chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
            command = [chrome_path, debugging_port]
            subprocess.Popen(command)
    except subprocess.CalledProcessError as e:
        print(f"启动 Chrome 时出错: {e}")
    # 随机化请求间隔
    sleep_time = random.uniform(2, 5)
    time.sleep(sleep_time)

    # 存储结果的列表
    results = []

    with sync_playwright() as p:
        for keyword in keywords:
            # 启动浏览器，禁用无头模式
            browser = p.chromium.connect_over_cdp("http://localhost:9999")
            content = browser.contexts[0]
            page = content.new_page()

            print(f"开始搜索关键词: {keyword}")
            # 构建 Google 搜索 URL，包含时间段筛选
            base_url = f'https://www.google.com/search?q={keyword}&tbs=cdr:1,cd_min:{start_date},cd_max:{end_date}'
            url = base_url
            for page_num in range(config['search_pages']):
                print(f"正在搜索第 {page_num + 1} 页")
                if page_num > 0:
                    # 计算后续页面的 URL
                    start_index = page_num * 10
                    url = f'{base_url}&start={start_index}'

                page.goto(url)
                # 随机化请求间隔
                sleep_time = random.uniform(2, 5)
                time.sleep(sleep_time)

                search_results = page.query_selector_all('div.g')
                result_count = 0
                for result in search_results:
                    try:
                        title = result.query_selector('h3').text_content()
                        link = result.query_selector('a').get_attribute('href')
                        if not is_domain_blacklisted(link, blacklist):
                            # 提取摘要
                            snippet = result.query_selector('div.VwiC3b').text_content() if result.query_selector(
                                'div.VwiC3b') else ""
                            results.append({'keyword': keyword, 'title': title, 'link': link, 'snippet': snippet})
                            result_count += 1
                    except Exception as e:
                        print(f"Error extracting result: {e}")
                print(f"从第 {page_num + 1} 页获取到 {result_count} 条结果")

            browser.close()
            print("关闭浏览器连接")

    # 将结果转换为 DataFrame
    df = pd.DataFrame(results)
    print(f"共获取到 {len(results)} 条搜索结果")

    # 保存结果到 CSV 文件，使用配置中的输出文件名
    output_file = config['output_file']
    df.to_csv(output_file, index=False, encoding='utf-8')
    print(f"搜索结果已保存到 {output_file}")

    # 生成 HTML 文件
    html_output_file = output_file.replace('.csv', '.html')
    csv_to_html(output_file, html_output_file)


if __name__ == "__main__":
    main()
