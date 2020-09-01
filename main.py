#!/usr/bin/python
# -*- coding:UTF-8 -*-
'''
Author: zhangpeiyu
Date: 2020-09-01 21:51:45
LastEditTime: 2020-09-02 01:14:56
Description: 我不是诗人，所以，只能够把爱你写进程序，当作不可解的密码，作为我一个人知道的秘密。
'''
import file_to_json

from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH


def main():
    document = Document()
    document.add_heading(u'接口文档', 0)

    json_data = file_to_json.getJson('api.json')

    root = '/root'
    tags = json_data['tags']
    paths = json_data['paths']
    definitions = json_data['definitions']

    temp = {}
    for url, url_info in paths.items():
        for method, req_info in url_info.items():
            for prop, val in req_info.items():
                if prop == 'tags':
                    if val[0] in temp:
                        arr = temp[val[0]]
                        item2 = {}
                        item2[method] = req_info
                        item = {}
                        item[url] = item2
                        arr.append(item)
                        temp[val[0]] = arr
                        break
                    else:
                        arr = []
                        item2 = {}
                        item2[method] = req_info
                        item = {}
                        item[url] = item2
                        arr.append(item)
                        temp[val[0]] = arr
                        break

    for index1, item1 in enumerate(tags):
        first_title = item1['name']
        # 一级标题
        document.add_heading(''.join([str(index1 + 1), '.', first_title]), 1)

        # 二级标题
        for tag, values in temp.items():
            if tag == first_title:
                for index2, item2 in enumerate(values):
                    ind = index2
                    for url, url_info in item2.items():
                        for method, method_info in url_info.items():
                            for prop, val in method_info.items():
                                if 'summary' == prop:
                                    document.add_heading(
                                        ''.join([str(index1 + 1), '.', str(ind + 1), '.', val]), 2)
                                    break
                            # 添加段落
                            paragraph = document.add_paragraph(u'请求方式：')
                            paragraph.add_run(method)
                            paragraph.add_run('\n')

                            paragraph.add_run(u'请求链接：')
                            paragraph.add_run(root + url)
                            paragraph.add_run('\n')

                            paragraph.add_run(u'功能描述：')
                            paragraph.add_run(method_info['description'])
                            paragraph.add_run('\n')

                            paragraph.add_run(u'请求类型（Content-Type）：')
                            paragraph.add_run(
                                '、'.join(method_info.get('consumes', [])))
                            paragraph.add_run('\n')

                            paragraph.add_run(u'响应类型（Content-Type）：')
                            paragraph.add_run(
                                '、'.join(method_info.get('produces', [])))
                            paragraph.add_run('\n')

                            parameters = method_info.get('parameters', [])
                            if len(parameters) > 0:
                                paragraph.add_run(u'请求参数：')
                                table = document.add_table(
                                    rows=1, cols=5)
                                hdr_cells = table.rows[0].cells
                                hdr_cells[0].text = u'字段名'
                                hdr_cells[1].text = u'参数位置'
                                hdr_cells[2].text = u'字段类型'
                                hdr_cells[3].text = u'是否必填'
                                hdr_cells[4].text = u'备注'
                                for row_index, row_data in enumerate(parameters):
                                    if 'body' != row_data.get('in', ''):
                                        table.add_row()
                                        hdr_cells = table.rows[row_index + 1].cells
                                        hdr_cells[0].text = row_data.get(
                                            'name', '')
                                        hdr_cells[1].text = row_data.get(
                                            'in', '')
                                        hdr_cells[2].text = row_data.get(
                                            'type', '').capitalize()
                                        hdr_cells[3].text = str(
                                            row_data.get('required', False))
                                        hdr_cells[4].text = row_data.get(
                                            'description', '')
                                    else:
                                        properties = {}
                                        if '$ref' in row_data['schema']:
                                            ref = row_data['schema']['$ref'].replace(
                                                '#/definitions/', '')
                                            print(ref)
                                            properties = definitions[ref]['properties']
                                        else:
                                            ref = row_data['schema']['items']['$ref'].replace(
                                                '#/definitions/', '')
                                            print(ref)
                                            properties = definitions[ref]['properties']

                                        for property_index, property_value in enumerate(properties):
                                            table.add_row()
                                            hdr_cells = table.rows[property_index + 1].cells
                                            print(property_value)
                                            if 'type' in properties[property_value]:
                                                hdr_cells[0].text = property_value
                                                hdr_cells[1].text = 'body'
                                                hdr_cells[2].text = properties[property_value].get(
                                                    'type', '').capitalize()
                                                hdr_cells[3].text = str(
                                                    properties[property_value].get('required', False))
                                                hdr_cells[4].text = properties[property_value].get(
                                                    'description', '')
                                            else:
                                                hdr_cells[0].text = property_value
                                                hdr_cells[1].text = 'body'
                                                hdr_cells[2].text = properties[property_value].get(
                                                    '$ref', '').replace('#/definitions/', '').capitalize()
                                                hdr_cells[3].text = str(
                                                    properties[property_value].get('required', False))
                                                hdr_cells[4].text = properties[property_value].get(
                                                    'description', '')
                            else:
                                paragraph.add_run(u'请求参数：')
                                paragraph.add_run(u'无')

                            paragraph = document.add_paragraph(u'响应参数：')
                            paragraph.add_run(u'未完成')
                            paragraph.add_run('\n')
    # 保存文档
    document.save('test.docx')


if __name__ == '__main__':
    main()
