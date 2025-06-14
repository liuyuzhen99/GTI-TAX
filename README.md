# GTI-TAX

## Business background业务背景
SAP GTI invoice format and Tax Bureau format not match, need a script to transform
SAP GTI模块下载下来的发票数据和上传税局开票所用数据格式不匹配，需要一个脚本转换下格式

## Function功能
upload tables to the application上传以下表格到应用:
1. GTI导出文件
2. 税局模版
3. 税局导出数据

based on above file to generate clock info for each staff, and support to check the result by using employee id根据以上信息生成每个员工每月的打卡记录，并可根据员工号进行筛选

![image](https://github.com/user-attachments/assets/456e37c2-d92f-4161-867a-c0af2d113d12)

![image](https://github.com/user-attachments/assets/275f19c1-5886-4b70-852e-e5c54c4241c9)


## logic逻辑
1. main function: excelConverter.py 主函数
2. readExcelClass.py: read excel and generate dataFrame 使用pandas读取excel文件
3. exception_handler.py: catch the exception and emit signal 用于捕获异常并发送信号
4. bindClass.py(main logic): read the uploaded file and process the file, generate the result finally 读取上传的文件并处理，最终生成结果
5. excel2ppt_qt.py: "front end" logic, define the function when click the button and show error message 界面，按钮函数，展示异常

## run运行
dist/excelConverter_v1.2.exe

