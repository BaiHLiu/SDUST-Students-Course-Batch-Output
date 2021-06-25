# SDUST-Students-Course-Batch-Output

> 批量导出山东科技大学指定学号范围的学生课表

## 学号模板格式

仅支持xls文件，从第一行第一列开始，一行一个学号。

## 使用方法

### 环境

Python3 (>=3.6)

### 安装依赖

```shell
python3 -m pip install -r requirements.txt
```

### 修改配置

只有一个main.py，修改文件开头的相关配置并运行。

```
#查询接口地址（必须校内网或走代理）
API_END_POINT = "http://192.168.111.167:8081/api/v3/students"
#配置关注词（导出时优先放在前面）
IMPORTANT_COURSES = [
    "体育"
]
#开始时间
START_TIME = "2020-09-10"
#截止时间
END_TIME = "2021-02-01"
#输入文件名
SRC_FILE_NAME = 'templete.xls'
#输出文件名
DEST_FILE_NAME = 'out.xls'
```

### 运行

```shell
python3 main.py
```



