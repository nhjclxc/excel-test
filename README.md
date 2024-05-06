# excel-test
excel模板导出练习

## 介绍
- 鉴于easyExcel在填充list模板之后不能在添加数据这一缺陷。本人基于poi实现复杂Excel模板实现导出，在填充了List数据之后还允许填充模板数据。
- 使用方法直接看工具类 [ExcelExporterMultSheetUtils.java](src%2Fmain%2Fjava%2Fcom%2Fexample%2Fexceltest%2FExcelExporterMultSheetUtils.java)
- 使用之前应当先导入[pom.xml](pom.xml)里面的相关依赖。

## 例子
### 模板sheet1
![./img.png](img.png)
### 模板sheet2
![./img_1.png](img_1.png)

### sheet1结果
![./img_2.png](img_2.png)
### sheet2结果
![./img_3.png](img_3.png)

### 支持列表数据的导出导入，该工具类的请看[ExcelUtils.java](src%2Fmain%2Fjava%2Fcom%2Fexample%2Fexceltest%2FExcelUtils.java)
![./img_4.png](img_4.png)

## Author：[LuoXianchao]()
