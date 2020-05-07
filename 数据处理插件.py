
import pandas as pd
import time
import datetime
import os
import re
import dateutil  # 识别日期格式
import chardet  # 文件编码格式判断
import numpy as np

# import plotly.express as px
# import plotly.graph_objects as go
# from plotly.subplots import make_subplots
# import plotly.offline as po


def is_wrong():
    while True:
        print('程序运行出错，请确认表格格式正确，包含指定字段')
        mode = input('输入q, 退出提示并继续：')
        if re.search('q', mode, re.IGNORECASE):
            break


def read_file(file_path):
    if file_path.endswith('csv'):
        try:
            df = pd.read_csv(file_path)
            return df
        except:
            try:
                f = open(file_path, 'rb')  # 先用二进制打开
                data = f.read()  # 读取文件内容
                file_encoding = chardet.detect(data).get('encoding')
                f.close()
                df = pd.read_csv(file_path, encoding=file_encoding, engine='python')
                return df
            except Exception as e:
                print(e)
                return None

    elif file_path.endswith('xlsx'):
        df = pd.read_excel(file_path)
        return df
    else:
        print("{}不是excel或者csv文件".format(file_path))
        is_wrong()
        return None


def df_to_excel(df, file_name):
    """
    excel生成过程中 url会被自动识别，但是URL过长会导致报错，“不将字符串转化为链接即可”
    :param df: pandas DataFrame
    :param file_name: 输出文件名字
    :return: 将df按照Excel格式保存到指定路径 不返回内容
    """
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter', options={'strings_to_urls': False})
    df.to_excel(writer, index=False)
    writer.close()
    return "{}已被保存到{}".format(df, file_name)


def in_all(data_path, fill_name=True):
    """
    合并Excel文件，在原路径下创建文件夹data/保存合并后的文件
    :param data_path: 多个Excel文件的保存路径
    :return:
    """
    while True:
        if not os.path.isdir(data_path):
            print('未找到此文件夹，请重新输入')
            data_path = input('文件夹路径：').strip()
            is_wrong()
        else:
            break

    row_df = pd.DataFrame()
    for each in os.listdir(data_path):
        file = data_path + "/" + each
        if each.endswith('xlsx'):
            df = pd.read_excel(file)
            if fill_name:
                df['fname'] = each.split('.')[0]
            row_df = pd.concat([row_df, df], sort=False)
        if each.endswith('csv'):
            df = read_file(file)
            if fill_name:
                df['fname'] = each.split('.')[0]
            row_df = pd.concat([row_df, df], sort=False)

    aft = datetime.datetime.strftime(datetime.datetime.now(), '_%m%d_%H%M')
    out_path = data_path + '/data/'
    if not os.path.exists(out_path):
        os.makedirs(out_path)

    out_file = out_path + '/合并' + aft + '.xlsx'
    df_to_excel(df=row_df, file_name=out_file)
    print("文件合并保存在目录：{}".format(os.path.realpath(out_path)))


def get_pic(file_path, height='140', width=''):
    """
    :param file_path: Excel文件路径
    :param key_words: 图片名称字段
    :param pic_path:  图品存储路径
    :param height:  图片高度
    :param width:  图片宽度 不建议同时设置图片高度 宽度
    :return:
    """

    while True:
        if not os.path.exists(file_path):
            print('未找到该文件...', )
            file_path = input('请确认需要处理的文件路径：').strip()
            is_wrong()
        else:
            break
    try:
        print(pd.read_excel(file_path).columns.to_list())
    except Exception as e:
        print(e)
        print('Excel文件选择错误')
        is_wrong()
    key_words = input('请在上述字段中选择，作为图片命名的字段(默认‘asin’,可直接回车)：').strip()
    if not key_words:
        key_words = 'asin'
    pic_path = input('请输入图片存储路径：').strip()
    print('选择的图片路径为：{}'.format(pic_path))
    print('添加辅助列中...')
    data = pd.read_excel(file_path)

    data['pic_url'] = pic_path + r'/' + data[key_words] + ".jpg"
    data['pic_table_url'] = '<table> <img src=' + '\"' + data['pic_url'] + '\"' + ' height=' + height + \
                            ' width=' + width + '>'
    # data.drop(columns=['pic_url'], inplace=True)
    data.to_excel(file_path.replace('.xlsx', '_添加图片辅助列.xlsx'), engine='xlsxwriter', encoding='utf-8')
    print("图片辅助列已添加，在原Excel文件路径下查看")


# def daily_stock(stock_file, sale_file):
#     df_stock = pd.read_excel(stock_file)
#     df_sales = pd.read_excel(sale_file)
#     if "" not in df_stock.columns:
#         print('请确认库存数据包含字段：')
#         is_wrong()
#     if "" not in df_sales.columns:
#         print('请确认销量数据包含字段：')
#         is_wrong()


def marketplace_choose(x):
    if re.search(r'(com|us)', x, re.IGNORECASE):
        return 'USA'
    if re.search(r'ca', x, re.IGNORECASE):
        return 'CA'
    if re.search(r'jp', x, re.IGNORECASE):
        return 'JP'
    return x


def get_deliver(sales_file, stock_file, on_the_way=6, mini_security_day=7, security_day=30):
    print('如需默认参数，请直接回车')
    # on_the_way = input('请输入在途运输时间(默认6天)：') or 6
    # mini_security_day = input('输入安全库存天数：(默认7天)') or 7
    # security_day = input('输入安全库存天数：(默认30天)') or 30
    #
    while True:
        try:
            on_the_way = int(input('请输入在途运输时间(默认6天)：') or 6)
            mini_security_day = int(input('输入安全库存天数：(默认7天)') or 7)
            security_day = int(input('输入安全库存天数：(默认30天)') or 30)

        except Exception as e:
            print(e)

        print('在途时间{}，最小库存天数{}，安全库存天数为{}'.format(on_the_way,mini_security_day,security_day))
        key_out = input('确认参数，直接回车或者输入q, 其他输入重新填写：',)
        if not key_out or (key_out == 'q'):
            break

    def get_stock_df(stock_file):
        df_stock = read_file(stock_file)
        date = max(df_stock['snapshot-date'])
        # df_stock = df_stock[df_stock['currency'] == 'USD']
        df_stock_res = df_stock[['fnsku', 'sku', 'asin', 'sellable-quantity', 'in-bound-quantity']]
        df_stock_res.columns = ['fnsku', 'sku', 'asin', '可售数量', '在途数量']
        sku_unique = df_stock_res['sku'].value_counts()[0]
        print(sku_unique)
        df_stock_res['库存记录时间'] = dateutil.parser.parse(date.replace('PDT', '').replace('PST', '')).date()
        if sku_unique > 1:
            print('输入表格中sku列有重复，请核实')
            print('仅支持单个站点操作，请筛选数据后重新尝试')
            is_wrong()
        try:
            df_stock_res.set_index(keys='sku', inplace=True)
        except Exception as e:
            print(e)

        return df_stock_res

    def get_sales_df(sales_file):

        df = read_file(sales_file)
        df['quantity'] = df['quantity'].mask(df['type'] == 'Refund', -df['quantity'])  # 退回数量转化为负数
        df_order = df[df['type'].map(lambda x: x in ['Order', 'Refund'])]  # 得到仅仅包含 Order Refund 的订单
        df_order_res = df_order[['sku', 'quantity', 'type', 'total', 'date/time']]
        df_order_res['date'] = df_order_res['date/time'].map(
            lambda x: dateutil.parser.parse(x.replace('PDT', '').replace('PST', '')).date())

        df_order_res['date_diff'] = df_order_res['date'].apply(lambda x: (df_order_res['date'].max() - x).days)
        df_in_7 = pd.DataFrame(df_order_res[df_order_res['date_diff'] < 7].groupby('sku')['quantity'].agg(np.sum))
        df_in_7.columns = ['7天销量']
        df_in_15 = pd.DataFrame(df_order_res[df_order_res['date_diff'] < 15].groupby('sku')['quantity'].agg(np.sum))
        df_in_15.columns = ['15天销量']
        df_in_30 = pd.DataFrame(df_order_res[df_order_res['date_diff'] < 30].groupby('sku')['quantity'].agg(np.sum))
        df_in_30.columns = ['30天销量']

        df_sum_sale = pd.concat([df_in_7, df_in_15, df_in_30], axis=1, sort=True)

        df_sum_sale.fillna(0, inplace=True)

        return df_sum_sale

    def time_cls(x):

        if not x:
            return '滞销'
        if re.search(r'inf', str(x)):
            return '滞销'
        if x == -6:
            return "断货-考察"
        elif x < 10:
            return '立即'
        elif x <= 15:
            return '弹性'
        elif 15 <= x:
            return '暂不'
        else:
            return '滞销'

    def get_last(df_sum_sale, df_stock_res):
        df_last = pd.concat([df_sum_sale, df_stock_res], sort=True, axis=1)

        df_last[['7天销量', '15天销量', '30天销量', '可售数量', '在途数量']] = df_last[
            ['7天销量', '15天销量', '30天销量', '可售数量', '在途数量']].fillna(0)

        df_last['平均日销量'] = (df_last['7天销量'] / 7 + df_last['15天销量'] / 15 + df_last['30天销量'] / 30) / 3
        df_last['平均日销量'] = df_last['平均日销量'].apply(lambda x: 0 if x < 0 else x)
        df_last['总库存量'] = df_last['可售数量'] + df_last['在途数量']
        df_last['可售天数'] = df_last['总库存量'] / df_last['平均日销量']
        df_last['最晚发货时间'] = df_last['可售天数'] - on_the_way
        df_last['最小安全库存'] = df_last['平均日销量'] * (on_the_way + mini_security_day)
        df_last['安全库存'] = df_last['平均日销量'] * security_day
        df_last['建议补货数量'] = df_last['安全库存'] - df_last['总库存量']
        df_last['应发状态'] = df_last['最晚发货时间'].apply(time_cls)
        # df_last['当天补货推荐'] = (security_day - df_last['最晚发货时间']) * df_last['平均日销量'] - df_last['总库存量']
        df_last['制表时间'] = datetime.datetime.now().date()
        # df_last['站点'] = 'US'
        # df_last['当天补货推荐'] = df_last['当天补货推荐'].fillna(0)
        for i in ['最晚发货时间', '最小安全库存', '安全库存', '建议补货数量']:
            try:
                df_last[i] = df_last[i].apply(lambda x: round(x, 2) if (x and not re.search('inf', str(x))) else x)
            except Exception as e:
                print(e)

        # 按照补货数量 对 补货状态进行调整
        df_last['应发状态'] = df_last['应发状态'].mask(pd.DataFrame([df_last['建议补货数量'].map(lambda x: 0 < x < 5),
                                                             df_last['最晚发货时间'] != -6]).all(), '考察')

        return df_last

    df_sales = get_sales_df(sales_file)
    df_stock = get_stock_df(stock_file)
    df_last = get_last(df_stock_res=df_stock, df_sum_sale=df_sales)
    aft = datetime.datetime.now().strftime('_%Y-%m-%d-%H-%M') + "_发货计划.xlsx"
    file_name = stock_file.replace('.xlsx', aft).replace('.csv', aft)
    df_to_excel(df=df_last, file_name=file_name)
    print('发货计划存储位置：{}'.format(file_name))


def get_daily(data_path):
    """
    仅仅适用于日常数据，需要具有下面这些字段
    :param file_path: 保存当月文件的路径
    :return: 当月的汇总表
    """
    while True:
        if not os.path.exists(data_path):
            print('未找到此文件夹，请重新输入')
            data_path = input('文件夹路径：').strip()
        else:
            break
    data_sum = pd.DataFrame()
    for each in os.listdir(data_path):
        file = data_path + "/" + each
        if each.endswith('xlsx'):
            df = read_file(file)
            data_time = each.split('.')[0]
            try:
                data_time = dateutil.parser.parse(data_time).date()
            except Exception as e:
                print(e)
                print('文件名并非可识别日期的格式，请更改文件名后重试')
                is_wrong()
            df['date'] = data_time
            data_sum = pd.concat([data_sum, df], sort=False)
        if each.endswith('csv'):
            df = read_file(file)
            data_time = each.split('.')[0]
            data_time = dateutil.parser.parse(data_time).date()
            df['date'] = data_time
            data_sum = pd.concat([data_sum, df], sort=False)

    for each in data_sum.columns:
        if re.search('父.*ASIN', each, re.I):
            data_sum['asin'] = data_sum[each]
        if re.search('子.*ASIN', each, re.I):
            data_sum['sub_asin'] = data_sum[each]

    data_info = data_sum[['date', 'asin', 'sub_asin', '买家访问次数', '订单商品数量转化率', '已订购商品数量', '已订购商品销售额']]
    # if 'date' in data_info.columns:
    #     try:
    #         data_info['date'] = data_info['date'].apply(lambda x: time.strftime('%Y-%m-%d', time.localtime(x / 1000)))
    #     except:
    #         pass
    data_info.head()

    def get_num(strr):
        strr = str(strr)
        if re.search('￥', strr):
            try:
                strr = float(strr.strip('￥').replace(",", ''))
                return strr
            except Exception as e:
                print(e)

        if re.search(r'US', strr):
            try:
                strr = float(strr.strip('US$').replace(",", ''))
                return strr
            except Exception as e:
                print(e)
        return strr

    data_info['已订购商品销售额'] = data_info['已订购商品销售额'].apply(get_num)
    data_info['买家访问次数'] = data_info['买家访问次数'].apply(lambda x: int(str(x).replace(',', '')))
    data_info['订单商品数量转化率'] = data_info['订单商品数量转化率'].apply(lambda x: float(str(x).strip('%')) * 0.01)
    print(data_info.columns)
    if 'Unnamed 0' in data_info.columns:
        data_info.drop(labels='Unnamed 0', inplace=True)
    output_path = data_path + '/data/'
    if not os.path.exists(output_path):
        os.makedirs(output_path)
    aft = datetime.datetime.strftime(datetime.datetime.now(), '%Y-%m-%d')
    df_to_excel(df=data_info, file_name=output_path + 'data_info_' + aft + '.xlsx')
    print('文件保存在：{}'.format(output_path))

#
# def get_html(file):
#     file_path, file_name = os.path.split(os.path.realpath(file))
#     html_path = file_path + '/html/'
#     if not os.path.exists(html_path):
#         os.makedirs(html_path)
#
#     df_row = read_file(file)
#     asin_list = list(df_row['asin'].unique())
#     date_list = list(df_row['date'].unique())
#     try:
#         date_mouth = datetime.datetime.strftime(dateutil.parser.parse(date_list[0]), '%Y-%m')
#     except Exception as e:
#         print(e)
#         date_mouth = str(date_list[0])
#
#     df_asin = df_row.groupby(by=['asin', 'date']).agg(
#         {'买家访问次数': np.sum, '已订购商品数量': np.sum}).reset_index()
#     fig_asin = px.scatter(df_asin, x='date', y='asin', size='买家访问次数', color='已订购商品数量')
#
#     asin_count = len(asin_list)
#     col_num = 3
#     row_num = asin_count // 3 + 1
#
#     fig_sub_asin = make_subplots(rows=row_num, cols=col_num)
#     for i in range(asin_count):
#         asin = asin_list[i]
#         df = df_row[df_row['asin'] == asin].groupby(by=['sub_asin', 'date']).agg(
#             {'买家访问次数': np.sum, '已订购商品数量': np.sum}).reset_index()
#         print(df.head())
#         fig_each = go.Scatter(x=df['sub_asin'],
#                               y=df['date'],
#                               # text=df['pic_url'],
#                               mode='markers',  # lines lines+markers
#                               marker=dict(
#                                   size=df['买家访问次数'],
#                                   color=df['已订购商品数量'],
#                               ),
#                               text=df[['买家访问次数', '已订购商品数量']].to_dict('row'),
#                               name=asin,
#                               # hovertemplate="%{x}<br>%{y}<br>%{text}",
#                               hovertemplate=
#                               "%{x}<br>%{y}<br>%{text}"
#                               )
#         fig_sub_asin.add_trace(fig_each, i // 3 + 1, i % 3 + 1)
#
#     fig_sub_asin.update_layout_images(dict(
#         # xref="paper",
#         # yref="paper",
#         sizex=0.1,
#         sizey=0.1,
#         xanchor="left",
#         yanchor="top"
#     ))
#     fig_sub_asin.update_layout(
#         showlegend=True,
#         title=date_mouth + '月度浏览量、销量比较',
#         width=600 * col_num,
#         height=600 * row_num,
#     )
#
#     aft = datetime.datetime.strftime(datetime.datetime.now(), '%H_%M')
#     asin_file_name = html_path + date_mouth + "_" + aft + '父ASIN_浏览量_销量汇总.html'
#     sub_asin_file_name = html_path + date_mouth + "_" + aft + '子ASIN_浏览量_销量汇总.html'
#     po.plot(fig_asin, filename=asin_file_name)
#     po.plot(fig_sub_asin, filename=sub_asin_file_name)


def main():
    """
    主函数，控制功能选择
    :return:
    """
    print('请选择要使用的功能:')
    print('-' * 30)
    for i in ['1: 表格合并', '2: 日常数据生成','3: 添加图片辅助列', '4: 生成发货表格']:
        print('-' * 8 + '{:-<20}'.format(i))
    print('-' * 30)

    while True:
        try:
            mode = int(input('输入对应功能的数字：'))
            if mode in [1, 2, 3, 4]:
                break
        except Exception as e:
            print(e)
            print('请确定输入1-4的数字')

    while True:

        if mode == 1:
            print('输入文件夹路径，自动合并Excel文件，需要文件字段相同')
            data_path = input('输入文件夹路径：').strip()
            fill_name = input('选择是否将文件名作为新的列(Y/N)：')
            if re.search('y', fill_name, re.IGNORECASE):
                fill_name = True
            elif re.search('n', fill_name, re.IGNORECASE):
                fill_name = False
            in_all(data_path=data_path, fill_name=fill_name)
            time.sleep(3)

        if mode == 2:
            print('输入日常数据保存的文件夹路径：')
            data_path = input('输入文件夹路径：').strip()
            get_daily(data_path)

        if mode == 3:
            print('输入excel文件路径，添加图片辅助列')
            file_path = input('输入Excel文件路径：').strip()
            get_pic(file_path)
            time.sleep(3)

        if mode == 4:
            print('依据所传表格生成发货计划, 仅支持单站点操作，多站点会出现sku重复')
            sales_file = input('销量表格：').strip()
            stock_file = input('现有库存表格：').strip()
            get_deliver(sales_file, stock_file)
            time.sleep(3)

        exit_mode = input('输入q,退出程序, 其他输入继续执行：')
        if re.search(r'q', exit_mode):
            break


if __name__ == '__main__':
    main()
