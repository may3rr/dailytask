# dailytask

# 使用百度情感分析API对Excel文件中的评论进行情感分析

## 环境准备

首先，确保你已经安装了Python和pip。然后，根据你使用的开发工具，选择合适的方式来安装所需的库。

### 如果你使用PyCharm、VSCode或其他非Jupyter Notebook的编辑器

在命令行中输入以下命令：

```bash
pip install baidu-aip
pip install openpyxl
```

### 如果你使用Jupyter Notebook

在代码单元中输入以下命令：

```bash
%pip install baidu-aip
%pip install openpyxl
```

## 代码说明

### 导入必要的库

```python
from aip import AipNlp
from openpyxl import load_workbook
import time
from openpyxl import Workbook
```

### 设置百度API

你需要创建你的百度智能云账户，并在百度只能云控制台创建一个自然语言处理的应用，参考老师在群里发的`PDF`文件，获取`APP_ID`、`API_KEY`和`SECRET_KEY`。

```python
APP_ID = 'your_app_id'
API_KEY = 'your_api_key'
SECRET_KEY = 'your_secret_key'
client = AipNlp(APP_ID, API_KEY, SECRET_KEY)
```

### 加载Excel文件

确保你的Excel文件中包含了评论内容，并且你知道评论内容所在的列名。

```python
wb = load_workbook("your_file.xlsx")
ws = wb.active
```

### 进行情感分析

代码会遍历Excel文件中的所有评论，并使用百度API进行情感分析。

```python
comments_column = None

from aip import AipNlp
from openpyxl import load_workbook
import time
from openpyxl import Workbook

# 创建你的 AppID
# 此处参考老师发的pdf中，自行注册百度账户
# 获得自己的api
APP_ID = 'your_app_id'
API_KEY = 'your_api_key'
SECRET_KEY = 'your_secret_key'
client = AipNlp(APP_ID, API_KEY, SECRET_KEY)

# 加载 Excel 文件
wb = load_workbook("your_file.xlsx") # 这里替换成自己爬取的xlsx信息
ws = wb.active

# 获取所有评论
comments_column = None
for cell in ws[1]:
    if cell.value == "评论内容": # 把这里换成爬取xlsx内容评论部分的列名
        comments_column = cell.column
        break

if comments_column is None:
    print("找不到评论内容列")
else:
    result_data = [["评论内容", "情感积极", "情感消极", "置信度"]]

    for row in range(2, ws.max_row + 1):
        comment = ws.cell(column=comments_column, row=row).value

        # 使用百度API进行情感分析
        result = client.sentimentClassify(comment)

        if 'items' in result:
            emotion = result['items'][0]
            positive_prob = round(emotion['positive_prob'] * 100, 2)
            negative_prob = round(emotion['negative_prob'] * 100, 2)
            confidence = round(emotion['confidence'] * 100, 2)
        else:
            positive_prob = 0
            negative_prob = 0
            confidence = 0

        result_data.append([comment, positive_prob, negative_prob, confidence])

        # 在每次循环后等待一段时间
        time.sleep(1)  # 这里等待 1 秒

    # 创建一个新的 Excel 文件来保存结果
    result_wb = Workbook()
    result_ws = result_wb.active

    for row in result_data:
        result_ws.append(row)

    result_wb.save("emotion_result.xlsx") # 把这个修改成你想保存的名字
    print("情感分析结果已保存到 emotion_result.xlsx")

```

### 保存结果

情感分析的结果将会被保存到一个新的Excel文件中。

```python
result_wb.save("emotion_result.xlsx")
print("情感分析结果已保存到 emotion_result.xlsx")
```

## 运行代码

将上述代码复制到你的Python编辑器中，确保你已经准备好了Excel文件，并替换掉代码中的占位符（如`your_file.xlsx`、`your_app_id`等）。

运行代码，并检查生成的`emotion_result.xlsx`文件以查看结果。
