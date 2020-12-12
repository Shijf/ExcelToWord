# ExcelToWord

## 温馨提示
刚刚再用工具的时候，发现一个问题：当自定义文件名称的时候，请不要出现 windows 不支持的文件名称，尤其是空格，当前写的时候应该得捕获一下异常的，哎，太年轻呀。大家用的时候，记得避免一下，程序年久失修，我看了下代码，居然不会。

由Excel作为数据源，批量生成基于Word模板的文档，

## 特色

无需安装 office 即可转换，速度快递飞起。

## 使用说明(Use)

### 1.建立数据源（Build a datasource）

### 2.在数据源sheet1中填入需要批量生成的数据，在sheet2中填入每个生成文件的文件名 （注意：sheet1的第一行填写的为模板中对应数据的书签名称，sheet2中需要在第二行开始填写，第一行可以作说明处理）
![Image text](https://github.com/Shijf/ExcelToWord/blob/master/img/datasource1.png)

![Image text](https://github.com/Shijf/ExcelToWord/blob/master/img/datasource2.png)

![Image text](https://github.com/Shijf/ExcelToWord/blob/master/img/datasource3.png)

![Image text](https://github.com/Shijf/ExcelToWord/blob/master/img/datasource4.png)

### 3.建立模板文件（在需要填入数据的地方插入书签）

![Image text](https://github.com/Shijf/ExcelToWord/blob/master/img/templete.png)

### 4.运行本软件进行批量生成文档

![Image text](https://github.com/Shijf/ExcelToWord/blob/master/img/output.png)







