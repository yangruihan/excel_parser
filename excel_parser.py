#!/usr/bin/env python3
# -*- coding:utf-8 -*-

import sys
import xlrd


LOOP_TYPE_KEYS = ['repeated']
STRUCT_TYPE_KEYS = ['required_struct', 'optional_struct']


class ExcelParser:

    def parse_class(self, excel_path: str, sheet_idx: str or int, start_col: int, end_col: int = -1, max_element: int = sys.maxsize) -> list:
        """
        用字典描述的 Excel 定义的结构

        Args:
            excel_path: excel 路径
            sheet_idx: sheet 索引或名称
            start_col: 整型，开始解析的列数
            end_col: 整型，结束解析的列数
            max_element: 整型，最大元素个数

        Returns:

            返回一个数组和一个整型

            数组：
            表示指定范围内所包含的类型的数组
            其中每个元素是一个字典
            字典字段如下:
                name: 成员名
                type: 类型名，其中数组有前缀'[]'，结构体为'struct'，结构体数组为'[]struct'
                col: 该成员开始列
                comment: 该成员注释
                struct_type（结构体特有）: 结构体包含的子成员集合，每一个子成员 是一个与该字典形式相同的字典

        形如：
        [
            {
                "name": "id",
                "type": "uint32",
                "col": 0,
                "comment": " @区域ID"
            },
            {
                "name": "scene_id",
                "type": "uint32",
                "col": 1,
                "comment": " @所属场景id"
            },
            {
                "name": "pos",
                "type": "[]struct",
                "col": 4,
                "struct_type": [
                    {
                        "name": "pos_x",
                        "type": "int32",
                        "col": 5,
                        "comment": " @x"
                    },
                    {
                        "name": "pos_y",
                        "type": "int32",
                        "col": 6,
                        "comment": " @y"
                    },
                    {
                        "name": "pos_z",
                        "type": "int32",
                        "col": 7,
                        "comment": " @z1"
                    }
                ],
                "comment": " @结构体声明"
            },
            {
                "name": "pos_empty",
                "type": "[]struct",
                "col": 114,
                "struct_type": [
                    {
                        "name": "pos_empty_x",
                        "type": "int32",
                        "col": 115,
                        "comment": " @x"
                    },
                    {
                        "name": "pos_empty_y",
                        "type": "int32",
                        "col": 116,
                        "comment": " @y"
                    },
                    {
                        "name": "pos_empty_z",
                        "type": "int32",
                        "col": 117,
                        "comment": " @z"
                    }
                ],
                "comment": " @结构体声明"
            }
        ]
        """

        book = xlrd.open_workbook(excel_path)
        sheet = None
        if isinstance(sheet_idx, int):
            sheet = book.get_sheet(sheet_idx)
        elif isinstance(sheet_idx, str):
            for s in book.sheets():
                if s.name.strip() == sheet_idx.strip():
                    sheet = s
                    break

        if sheet is None:
            print('Sheet not found')
            return

        return self.parse_class_with_sheet(sheet, start_col, end_col, max_element)

    def parse_class_with_sheet(self, sheet: xlrd.sheet.Sheet, start_col: int, end_col: int = -1, max_element: int = sys.maxsize) -> list:
        """
        用字典描述的 Excel 定义的结构

        Args:
            sheet: xlrd sheet 对象
            start_col: 整型，开始解析的列数
            end_col: 整型，结束解析的列数
            max_element: 整型，最大元素个数

        Returns:

            返回一个数组和一个整型

            数组：
            表示指定范围内所包含的类型的数组
            其中每个元素是一个字典
            字典字段如下:
                name: 成员名
                type: 类型名，其中数组有前缀'[]'，结构体为'struct'，结构体数组为'[]struct'
                col: 该成员开始列
                comment: 该成员注释
                struct_type（结构体特有）: 结构体包含的子成员集合，每一个子成员 是一个与该字典形式相同的字典

        形如：
        [
            {
                "name": "id",
                "type": "uint32",
                "col": 0,
                "comment": " @区域ID"
            },
            {
                "name": "scene_id",
                "type": "uint32",
                "col": 1,
                "comment": " @所属场景id"
            },
            {
                "name": "pos",
                "type": "[]struct",
                "col": 4,
                "struct_type": [
                    {
                        "name": "pos_x",
                        "type": "int32",
                        "col": 5,
                        "comment": " @x"
                    },
                    {
                        "name": "pos_y",
                        "type": "int32",
                        "col": 6,
                        "comment": " @y"
                    },
                    {
                        "name": "pos_z",
                        "type": "int32",
                        "col": 7,
                        "comment": " @z"
                    }
                ],
                "comment": " @结构体声明"
            },
            {
                "name": "pos_empty",
                "type": "[]struct",
                "col": 114,
                "struct_type": [
                    {
                        "name": "pos_empty_x",
                        "type": "int32",
                        "col": 115,
                        "comment": " @x"
                    },
                    {
                        "name": "pos_empty_y",
                        "type": "int32",
                        "col": 116,
                        "comment": " @y"
                    },
                    {
                        "name": "pos_empty_z",
                        "type": "int32",
                        "col": 117,
                        "comment": " @z1"
                    }
                ],
                "comment": " @结构体声明"
            }
        ]
        """

        if end_col == -1:
            end_col = sheet.ncols - 1

        class_define = []

        i = start_col
        solved_element = 0

        while i <= end_col and solved_element < max_element:
            define, i = self._parse_col(sheet, i)
            class_define.append(define)

        return class_define

    def _parse_col(self, sheet: xlrd.sheet.Sheet, col: int) -> (dict, int):
        # 兼容 proto（required, optional, repeated）
        proto_type, define_type, name, comment = self._get_sheet_data(
            sheet, col)

        # 处理数组
        if proto_type in LOOP_TYPE_KEYS:
            # 判断是否为单列数组
            if isinstance(define_type, int) or isinstance(define_type, float):
                next_col = self._get_next(sheet, col)

                next_proto_type, next_define_type, next_name, next_comment = self._get_sheet_data(
                    sheet, next_col)

                arr_count = int(define_type)

                # 如果是结构体
                if next_proto_type in STRUCT_TYPE_KEYS:
                    struct_element_count = int(next_define_type)

                    ret = {
                        'name': next_name,
                        'type': '[]struct',
                        'col': next_col,
                        'struct_type': [],
                        'comment': next_comment
                    }

                    temp_col = next_col + 1
                    handle_count = 0

                    while handle_count < struct_element_count:
                        sub_type, temp_col = self._parse_col(sheet, temp_col)
                        ret['struct_type'].append(sub_type)
                        handle_count += 1

                    item_cnt = temp_col - col - 2

                    return ret, self._get_next(sheet, col + 1 + item_cnt * arr_count)
                else:
                    ret = {
                        'name': next_name,
                        'type': f'[]{next_define_type}',
                        'col': next_col,
                        'comment': next_comment
                    }

                    return ret, self._get_next(sheet, col + arr_count)
            else:
                return {
                    'name': name,
                    'type': f'[]{define_type}',
                    'col': col,
                    'comment': comment
                }, self._get_next(sheet, col)

        elif proto_type in STRUCT_TYPE_KEYS:
            struct_element_count = int(define_type)

            ret = {
                'name': name,
                'type': 'struct',
                'struct_type': [],
                'col': col,
                'comment': comment
            }

            temp_col = col + 1
            handle_count = 0

            while handle_count <= struct_element_count:
                sub_type, temp_col = self._parse_col(sheet, temp_col)
                ret['struct_type'].append(sub_type)
                handle_count = handle_count + 1

            return ret, temp_col

        elif self._is_skip_col(proto_type):
            return None

        else:
            return {
                'name': name,
                'type': define_type,
                'col': col,
                'comment': comment
            }, self._get_next(sheet, col)

    def _get_sheet_data(self, sheet: xlrd.sheet.Sheet, col: int) -> (str, str, str, str):
        proto_type = sheet.cell_value(0, col)
        define_type = sheet.cell_value(1, col)  # 定义的类型
        name = sheet.cell_value(2, col)  # 字段名
        comment = str(sheet.cell_value(4, col)).replace(
            '\n', '').replace('\r', '')  # 注释

        if comment != '':
            comment = f' @{comment}'

        return proto_type, define_type, name, comment

    def _get_next(self, sheet: xlrd.sheet.Sheet, col: int, max: int = -1) -> int:
        col = col + 1
        if max == -1:
            max = sheet.ncols

        if col >= max:
            return max

        proto_type = sheet.cell_value(0, col)

        while self._is_skip_col(proto_type):
            col = col + 1
            if col >= max:
                break

            proto_type = sheet.cell_value(0, col)

        if col >= max:
            return max
        else:
            return col

    def _is_skip_col(self, proto_type: str) -> bool:
        proto_type = proto_type.strip()

        if proto_type == "*" or proto_type == "":
            return True

        return False
