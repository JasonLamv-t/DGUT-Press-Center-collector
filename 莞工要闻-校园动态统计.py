import time
import requests
from bs4 import BeautifulSoup
import xlrd
import xlwt
import re

link_head = 'http://www.dgut.edu.cn'
link = [
    'http://www.dgut.edu.cn/xwzx/ggyw',
    'http://www.dgut.edu.cn/xwzx/xydt'
]


def get_author(art_uri):
    art_url = link_head + '/' + art_uri.replace('../', '')
    art = requests.get(art_url)
    art.encoding = art.apparent_encoding
    art_soup = BeautifulSoup(art.text, 'html.parser')
    return art_soup.find('li', {'class': 'unit'}).text[5:-6].split(' ')[0]


def is_img(art_uri):
    art_url = link_head + '/' + art_uri.replace('../', '')
    art = requests.get(art_url)
    art.encoding = art.apparent_encoding
    art_img = BeautifulSoup(art.text, 'html.parser').find('form', {'name': '_newscontent_fromname'}).find_all('img')
    return True if len(art_img) != 0 else False


if __name__ == '__main__':
    list_of_art = []
    continue_flag = True

    # 输入时间段和统计类型
    regex = '20[0-9]{2}-[0-9]{2}-[0-9]{2}'
    while True:
        start_date = input('请输入开始日期(含)，例:2020-01-01：')
        if re.match(regex, start_date):
            break
        else:
            print('输入错误，请重新输入')

    while True:
        end_date = input('请输入结束日期(含)，例:2020-03-31：')
        if re.match(regex, end_date):
            break
        else:
            print('输入错误，请重新输入')

    while True:
        try:
            count_type = int(input('请输入统计类型，1-莞工要闻，2-校园动态：'))
            if count_type == 1 or count_type == 2:
                break
            else:
                print('输入错误，请重新输入')
        except:
            print('输入错误，请重新输入')

    result_file_name = '莞工要闻' if count_type == 1 else '校园动态'
    result_file_name = start_date + '-' + end_date + result_file_name + '统计.xls'

    # 获取soup
    print('正在分析第1页')
    res = requests.get(link[count_type - 1] + '.htm')
    res.encoding = res.apparent_encoding
    soup = BeautifulSoup(res.text, 'html.parser')

    # 获取总页数
    pages = soup.find('div', {'class': 'pb_sys_common'}).find_all('span', {'class': 'p_no'})
    total_page = int(pages[len(pages) - 1].text)

    # 获取首页的列表
    art_list = soup.find('div', {'class': 'listList'}).find_all('li')

    # 遍历首页列表
    for item in art_list:
        # 获取文章日期
        art_date = item.find('span', {'class': 'time'}).text

        # 如果在统计范围内，获取对应信息
        if start_date <= art_date <= end_date:
            art_title = item.find('a').text
            art_uri = item.find('a').get('href')
            art_author = get_author(art_uri)
            art_author = '新闻中心' if art_author == '' else art_author
            art_fee = 30 if is_img(art_uri) else 20

            art = [art_date, art_author, art_title, art_fee]
            list_of_art.append(art)
            print(art)
        # 如果在统计范围后，继续遍历
        elif art_date > end_date:
            continue
        # 如果在统计范围前，停止遍历
        else:
            continue_flag = False
            break

    # 开始从第二页遍历
    for page in reversed(range(1, total_page)):
        # 如果已经遍历到在范围前到文章，停止遍历
        if not continue_flag:
            break

        time.sleep(3)

        # 获取列表页面，处理得到列表
        print('正在分析第{}页'.format(total_page - page + 1))
        res = requests.get(link[count_type - 1] + '/{}.htm'.format(page))
        res.encoding = res.apparent_encoding
        soup = BeautifulSoup(res.text, 'html.parser')
        art_list = soup.find('div', {'class': 'listList'}).find_all('li')

        # 对该页面列表进行遍历
        for item in art_list:
            # 获取文章日期
            art_date = item.find('span', {'class': 'time'}).text

            # 如果在统计范围内，获取对应信息
            if start_date <= art_date <= end_date:
                art_title = item.find('a').text
                art_uri = item.find('a').get('href')
                art_author = get_author(art_uri)
                art_author = '新闻中心' if art_author == '' else art_author
                art_fee = 30 if is_img(art_uri) else 20

                art = [art_date, art_author, art_title, art_fee]
                list_of_art.append(art)
                print(art)

            # 如果在统计范围后，继续遍历
            elif art_date > end_date:
                continue
            # 如果在统计范围前，停止遍历
            else:
                continue_flag = False
                break

    print('写入文件中')
    # 读取配置文件
    config_data = xlrd.open_workbook(r'配置文件_勿删.xls')
    config_table = config_data.sheet_by_name('Sheet1')

    config_sheets = []
    for i in range(config_table.nrows):
        row = config_table.row_values(i)
        # 去除配置文件内容未对齐导致的空行
        while '' in row:
            row.remove('')
        config_sheets.append(row)

    # 初始化要写入的列表和找不到的列表
    res_tables = [[] for i in range(config_table.nrows)]
    no_found = []

    # 根据配置文件找到文章的分页，并插入到对应的写入位置
    for art in list_of_art:
        no_in_row = True
        author = art[1]
        for row in range(len(config_sheets)):
            if author in config_sheets[row]:
                res_tables[row].append(art)
                no_in_row = False
                break

        if no_in_row:
            no_found.append(art)

    # 新建文件
    res_file = xlwt.Workbook()
    # 设置样式
    style = xlwt.XFStyle()  # 初始化样式
    font = xlwt.Font()  # 创建字体
    font.name = u'微软雅黑'  # 字体类型
    font.height = 280  # 字体大小   200等于excel字体大小中的10
    style.font = font  # 设定样式

    # 遍历要写入的三维列表
    for index, sheet in enumerate(res_tables):
        # 跳过没有文章的分页
        if len(sheet) == 0:
            continue
        # 创建sheet
        res_sheet = res_file.add_sheet(config_sheets[index][0])
        # 将表头插入尾部并反转
        sheet.append(['日期', '单位', '标题', '稿费'])
        for row, art in enumerate(reversed(sheet)):
            for col, text in enumerate(art):
                res_sheet.write(row, col, art[col], style)

    res_file.save(result_file_name)
    print('写入完成！请检查文件！')
    print('文件名：{}'.format(result_file_name))
    if len(no_found) != 0:
        print('以下结果未找到合适分页，请手动加入文件中，并调整配置')
        for art in reversed(no_found):
            print(art)

    input('输入任意键退出')
