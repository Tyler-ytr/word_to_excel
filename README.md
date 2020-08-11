## word批量提取信息到excel的小工具

这个项目源于我实习期间帮忙从word里面copy-paste到Excel的需求,手动做完之后感觉可能不太准确,因此同python写了一个脚本;

注意,仅仅适用于特定的格式的文件,对于其他文件还需要使用者进一步修改

#### 使用方式

将word_to_excel.py拷贝到装有word文件的文件夹的外面,

python word_to_excel.py

出现: 告诉我文件夹名称,输入装有word文件的文件夹名称;

结果在result_of_excel.xlsx里面

#### 开发环境

````
conda 4.8.3
````

其中:

````
pandas 1.0.5
python-docx 0.8.10
regex 2020.6.8
````

#### 测试

````
python3 word_to_excel.py
输入:test
输出:result_of_excel.xlsx
````

可以与我做的结果result_of_excel_for_test.xlsx比较
