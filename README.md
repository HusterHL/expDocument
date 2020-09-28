# expDocument
该项目用于输出检测报表，支持的报表格式有word/pdf/excel等。
注意事项：
1、项目是基于QT的界面设计，QAxObject开发word文档；使用QPainter开发pdf文档。
2、文档中的表格设计：对于word，使用插入表格的函数即可；对于pdf则需要画表格，注意表格的坐标以及长度。
3、输出报表的数据来源于底层的配置文件（输出一些固定的参数或文字）以及qt界面上面的数据或参数（人工添加）。
4、输出的pdf、或者word文档默认保存于doc文件夹下。
5、qss文件主要用于美化界面。