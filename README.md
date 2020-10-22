# Excel 解析器
Excel 特定格式解析器

## 依赖

- xlrd>=1.2.0

## 用法

Excel 形如：

|required|optional||repeated|optional_struct|optional|optional|optional|
|:---|:---|:---|:---|:---|:---|:---|:---|
|uint32|uint32||2|3|int32|int32|int32|
|id|scene_id|||pos|pos_x|pos_y|pos_z|
|||||||||
|ID|区域ID|备注|数组声明|结构体声明|x|y|z|
|1|1|大厅|7||100|200|300|

```python
from excel_parser import ExcelParser

ExcelParser = ExcelParser()

ret = ExcelParser.parse_class_with_sheet(sheet, 0, 7)
```

ret 内容如下：

```python
"""
[
    {
        "name": "id",
        "type": "uint32",
        "col": 0,
        "comment": " @ID"
    },
    {
        "name": "scene_id",
        "type": "uint32",
        "col": 1,
        "comment": " @区域id"
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
    }
]       
"""
```

