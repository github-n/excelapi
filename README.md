# Excel 导出模块

## node_modules

`excel4node`
JSON -> Excel

## 传入数据说明

* 使用 `post` 请求，请求头类型 `Content-Type=application/json`
* 输出表格第一行为表头，其他为数据。

## 字段说明

```js
{
    "exceldata": { // {}, 传入的数据
        "filename": "", // str, 文件名,不含.xlsx
        "sheets": [ // [], sheet信息
            {
                "sheetname": "", // str, sheet名称
                "column": [ // [], 表头信息
                    {
                        "dataKey":"", // str, 列标识
						"ref":1, // num, 列数
						"title":"", // str, 显示文字
						"width":25 // num, 列宽, 可省略
                    }
                ],
                "data"：[ // [], 所有行数据
                    {
                        "row": [ // [], 行数据
                            {
								"dataKey": "", // str, 匹配列标识
                				"text":"", // str|num, 显示文字
                				"type":"s|n", // str, 数据类型,"s":字符,"n":数字, 可省略
                        	}
                        ]
                	}
                ]
            }            
        ]
    }
}
```

## 数据样例

```js
{
	"exceldata":{
		"filename":"testoutput.xlsx",
		"sheets":[
			{
				"sheetname":"sheetname1",
				"column":[
					{
						"dataKey":"col1",
						"ref":1,
						"title":"第一列",
						"width":25
					},
					{
						"dataKey":"col2",
						"ref":2,
						"title":"第二列"
					},
					{
						"dataKey":"col3",
						"ref":3,
						"title":"第三列"
					}
				],
				"data":[
					{
						"row":[
							{
								"dataKey":"col1",
								"text":"第一行第一列",
								"type":"s"
							},
							{
								"dataKey":"col2",
								"text":"123456",
								"type":"n"
							},
							{
								"dataKey":"col3",
								"text":67890,
								"type":"s"
							}
						]
					},
					{
						"row":[
							{
								"dataKey":"col3",
								"text":"第二行第三列",
								"type":"s"
							},
							{
								"dataKey":"col2",
								"text":111111
							}
						]
					}
				]
			},
			{
				"sheetname":"sheetname2",
				"column":[],
				"data":[]
			},
            {
                
            }
		]
	}
}
```

## TODO

* JSON 判断
* 日期格式
