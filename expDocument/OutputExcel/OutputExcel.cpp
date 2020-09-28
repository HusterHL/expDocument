#include "OutputExcel.h"
//#include <QAxWidget>
//
//#include <QAxObject>
//
//#include <QAxBase>
//
//#include <QFile>
//
//#include <QTextStream>
//
//#include <QDir>
//#include <QAxObject>


OutputExcel::OutputExcel(QWidget *parent)
	: QMainWindow(parent)
{
	ui.setupUi(this);
	//open(true);
	//CreatExcel();
	//addText(QStringLiteral("标志及结构检查(%1)检测原始记录").arg(QStringLiteral("低压电缆")), TITLE_ONE, AlignParCenter);
	//QAxObject *selection =addText2(QStringLiteral("标志及结构检查(%1)检测原始记录").arg(QStringLiteral("低压电缆")));
	//selection->querySubObject("Range")->querySubObject("Font")->setProperty("Bold", true);
	connect(this, SIGNAL(ConfigSignal(QString)), this, SLOT(ShowConfig(QString)));
	connect(ui.checkBox, SIGNAL(stateChanged(int)), this, SLOT(hideTabWidget()));
	connect(ui.checkBox_2, SIGNAL(stateChanged(int)), this, SLOT(hideTabWidget()));
	widgetOne = ui.tabWidget->widget(1);
	widgetTwo = ui.tabWidget->widget(2);
	//ui.tabWidget->removeTab(1);
	//ui.tabWidget->removeTab(2);
	hideTabWidget();
	readConfig();
	
}

void OutputExcel::hideTabWidget()
{
	//connect(ui.checkBox, SIGNAL(stateChanged(int)), this, SLOT(hideTabWidget()));
	//ui.tabWidget->tabBar()->show();


	//ui.tabWidget->tabBar()->show();
	if (ui.checkBox->isChecked())
	{
		ui.tabWidget->addTab(widgetOne, QStringLiteral("配置信息"));
		ui.tabWidget->addTab(widgetTwo, QStringLiteral("PDF输出"));
		//ui.tabWidget->tabBar()->show();
		//ui.tabWidget->setCurrentIndex(1);
	}
	if (!ui.checkBox->isChecked())
	{
		int j=ui.tabWidget->indexOf(widgetOne);
		ui.tabWidget->removeTab(j);
		int i = ui.tabWidget->indexOf(widgetTwo);
		ui.tabWidget->removeTab(i);
	}

	
	//if (ui.checkBox_2->isChecked())
	//{
	//	ui.tabWidget->addTab(widgetTwo, QStringLiteral("PDF输出"));
	//	//ui.tabWidget->removeTab(2);
	//}
	//if (!ui.checkBox_2->isChecked())
	//{
	//	int j = ui.tabWidget->indexOf(widgetTwo);
	//	ui.tabWidget->removeTab(j);
	//}
}


void OutputExcel::on_BnCreatWord_clicked()
{
	if (ui.checkBox_7->isChecked())
	{
		if (!laodconfig)
		{
			QMessageBox::information(this, QStringLiteral("提示"), QStringLiteral("没有标准配置文件"));
			return;
		}
		open(false);
		CreatExcel();
		save();
		if (!OutputExcel::close())
		{
			QMessageBox::information(this, QStringLiteral("提示"), QStringLiteral("word版本保存完成"));
		}

	}
	if (ui.checkBox_8->isChecked())
	{
		if (!laodconfig)
		{
			QMessageBox::information(this, QStringLiteral("提示"), QStringLiteral("没有标准配置文件"));
			return;
		}
		OutputPDF();
		QMessageBox::information(this, QStringLiteral("提示"), QStringLiteral("PDF版本保存完成"));
	}
}

void OutputExcel::CreatExcel()
{
	if (!m_bOpened) return;
	QString str= ui.checkBox->text();
	QAxObject *selection = NULL;
	selection = m_wordWidget->querySubObject("Selection");
	if (selection)
	{
	
		selection->querySubObject("Font")->setProperty("Size", 17);
		selection->querySubObject("Font")->setProperty("Bold", true);
		selection->querySubObject("Range")->setProperty("Text", QStringLiteral("标志及结构检查(%1)检测原始记录").arg(str));
		selection->querySubObject("ParagraphFormat")->setProperty("Alignment", AlignParCenter); //文本位置设置

		//背景色
		//selection->querySubObject("Range")
		//	->querySubObject("ParagraphFormat")
		//	->querySubObject("Shading")
		//	->setProperty("BackgroundPatternColor",QColor("blue").value());
		moveToEnd();
		insertEnter();
		selection->querySubObject("Font")->setProperty("Size", 9);
		selection->querySubObject("Font")->setProperty("Bold", false);
		selection->querySubObject("ParagraphFormat")->setProperty("Alignment", AlignParLeft); //文本位置设置
		selection->querySubObject("Range")->setProperty("Text", QStringLiteral("文件代号：%1").arg(configBlack->DocumentCode));
		
		moveToEnd();
		insertEnter();
		selection->querySubObject("Font")->setProperty("Size", 11);
		QAxObject *table = new QAxObject;
		table = createTable(3, 6);
		//table->querySubObject("Borders")->setProperty("Enable", -2);
		setColumnWidth(table, 3, 80);
		setColumnWidth(table, 4, 85);
		setColumnWidth(table, 5, 40);
		MergeCells(table, 0, 0, 0, 2);
		MergeCells(table, 0, 1, 0, 3);
		MergeCells(table, 1, 0, 1, 2);
		MergeCells(table, 1, 1, 1, 3);
		MergeCells(table, 2, 0, 2, 2);
		//QAxObject *cell = table->querySubObject("Cell(int,int)", 2, 2);
		//QAxObject *borders1 = table->querySubObject("Borders(6)");
		//borders1->dynamicCall("SetLineStyle(int)", 0);
		//borders->setProperty("Enable", -2);
		QDate date = QDate::currentDate();
		QString create_time = date.toString(QStringLiteral("yyyy年MM月dd日"));
		QString SampleNumber = ui.lineEdit_2->text();
		//QString SampleNumber = "EETC08-19/05/23-";
		
		QString ManufacturPlant = ui.lineEdit_6->text();;
		//QString DetectLocation = "Detection location";
		QString DetectLocation = configBlack->DetectionLocation;
		//QString Modelspecificate = "Model specifications"; 
		QString Modelspecificate = ui.lineEdit_7->text();
		QString RoomTemperature = ui.lineEdit_8->text();
		QString RelativeHumidity= ui.lineEdit_9->text();
		table->querySubObject("Cell(int,int)", 1, 1)->querySubObject("Range")->dynamicCall("SetText(QString)", 
			QStringLiteral("样品编号：%1").arg(SampleNumber));
		table->querySubObject("Cell(int,int)", 1, 2)->querySubObject("Range")->dynamicCall("SetText(QString)",
			QStringLiteral("制 造 厂：%1").arg(ManufacturPlant));
		table->querySubObject("Cell(int,int)", 2, 2)->querySubObject("Range")->dynamicCall("SetText(QString)",
			QStringLiteral("检测地点：%1").arg(DetectLocation));
		table->querySubObject("Cell(int,int)", 2, 1)->querySubObject("Range")->dynamicCall("SetText(QString)",
			QStringLiteral("型号规格：%1").arg(Modelspecificate));		
		table->querySubObject("Cell(int,int)", 3, 1)->querySubObject("Range")->dynamicCall("SetText(QString)",
			QStringLiteral("检测时间：%1").arg(create_time));
		table->querySubObject("Cell(int,int)", 3, 2)->querySubObject("Range")->dynamicCall("SetText(QString)",
			QStringLiteral("室温：%1℃").arg(RoomTemperature));
		table->querySubObject("Cell(int,int)", 3, 3)->querySubObject("Range")->dynamicCall("SetText(QString)",
			QStringLiteral("相对湿度：%1%").arg(RelativeHumidity));
		//table->querySubObject("Cell(int,int)", 1, 1)->setProperty("Width", 60);

		moveToEnd(table);
		insertEnter();
		selection->querySubObject("Font")->setProperty("Size", 11);
		selection->querySubObject("Font")->setProperty("Bold", true);
		selection->querySubObject("Range")->setProperty("Text", QStringLiteral("一、检测仪器"));
		selection->querySubObject("ParagraphFormat")->setProperty("Alignment", AlignParLeft); //文本位置设置
		moveToEnd();
		insertEnter();
		selection->querySubObject("Font")->setProperty("Size", 10);
		selection->querySubObject("Font")->setProperty("Bold", false);
		selection->querySubObject("Range")->setProperty("Text", 
			QStringLiteral("%1").arg(configBlack->PragraphOne));
		moveToEnd();
		insertEnter();
		selection->querySubObject("Range")->setProperty("Text",
			QStringLiteral("%1").arg(configBlack->PragraphTwo));
		moveToEnd();
		insertEnter();
		selection->querySubObject("Range")->setProperty("Text",
			QStringLiteral("%1").arg(configBlack->PragraphThree));
		moveToEnd();
		insertEnter();
		selection->querySubObject("Range")->setProperty("Text",
			QStringLiteral("%1").arg(configBlack->PragraphFour));
		moveToEnd();
		insertEnter();

		selection->querySubObject("Font")->setProperty("Size", 11);
		selection->querySubObject("Font")->setProperty("Bold", true);
		selection->querySubObject("Range")->setProperty("Text", QStringLiteral("二、检测依据"));
		selection->querySubObject("ParagraphFormat")->setProperty("Alignment", AlignParLeft); //文本位置设置
		moveToEnd();
		insertEnter();
		selection->querySubObject("Font")->setProperty("Size", 8);
		selection->querySubObject("Font")->setProperty("Bold", false);
		selection->querySubObject("Range")->setProperty("Text",QStringLiteral("%1").arg(NomMinConfig->GBTName));
		moveToEnd();
		insertEnter();

		selection->querySubObject("Font")->setProperty("Size", 11);
		selection->querySubObject("Font")->setProperty("Bold", true);
		selection->querySubObject("Range")->setProperty("Text", QStringLiteral("三、检测前对检测仪器设备及被测样品的检查"));
		selection->querySubObject("ParagraphFormat")->setProperty("Alignment", AlignParLeft); //文本位置设置
		moveToEnd();
		insertEnter();
		selection->querySubObject("Font")->setProperty("Size", 10);
		selection->querySubObject("Font")->setProperty("Bold", false);
		selection->querySubObject("Range")->setProperty("Text",
			QStringLiteral("    1. 游标卡尺的零位正确 □；"));
		moveToEnd();
		insertEnter();
		selection->querySubObject("Range")->setProperty("Text",
			QStringLiteral("    2. 数显卡尺的零位正确 □；"));
		moveToEnd();
		insertEnter();
		selection->querySubObject("Range")->setProperty("Text",
			QStringLiteral("    3. 被测样品外观无损 □。"));
		moveToEnd();
		insertEnter();

		selection->querySubObject("Font")->setProperty("Size", 11);
		selection->querySubObject("Font")->setProperty("Bold", true);
		selection->querySubObject("Range")->setProperty("Text", QStringLiteral("四、检测数据及结果"));
		selection->querySubObject("ParagraphFormat")->setProperty("Alignment", AlignParLeft); //文本位置设置
		moveToEnd();
		insertEnter();

		selection->querySubObject("Font")->setProperty("Size", 9);
		selection->querySubObject("Font")->setProperty("Bold", false);
		QAxObject *table2 = new QAxObject;
		table2 = createTable(21, 9);
		setColumnWidth(table2, 0, 70);
		setColumnWidth(table2, 1, 30);
		setColumnWidth(table2, 2, 90);
		setColumnWidth(table2, 3, 37);
		setColumnWidth(table2, 4, 37);
		setColumnWidth(table2, 5, 37);
		setColumnWidth(table2, 6, 37);
		setColumnWidth(table2, 7, 37);
		setColumnWidth(table2, 8, 37);

		MergeCells(table2, 0, 3, 0, 8);
		MergeCells(table2, 1, 3, 1, 8);
		MergeCells(table2, 2, 3, 2, 8);
		MergeCells(table2, 3, 3, 3, 8);
		MergeCells(table2, 4, 3, 4, 8);
		MergeCells(table2, 5, 3, 5, 8);
		MergeCells(table2, 7, 3, 10, 8);
		MergeCells(table2, 11, 3,13, 8);
		MergeCells(table2, 14, 3, 14, 8);
		MergeCells(table2, 16, 3, 16, 8);
		MergeCells(table2, 5, 0, 6, 0);
		MergeCells(table2, 5, 1, 6, 1);
		MergeCells(table2, 5, 2, 6, 2);
		MergeCells(table2, 7, 0, 10, 0);
		MergeCells(table2, 7, 1, 10, 1);
		MergeCells(table2, 7, 2, 10, 2);
		MergeCells(table2, 8, 0, 12, 0);
		MergeCells(table2, 8, 1, 12, 1);
		MergeCells(table2, 8, 2, 10, 2);
		MergeCells(table2, 9, 2, 10, 2);
		MergeCells(table2, 6, 3, 6, 5);
		MergeCells(table2, 6, 4, 6, 6);
		MergeCells(table2, 10, 3, 10, 5);
		MergeCells(table2, 10, 4, 10, 6);
		MergeCells(table2, 12, 3, 12, 4);
		MergeCells(table2, 12, 4, 12, 5);
		MergeCells(table2, 12, 5, 12, 6);
		MergeCells(table2, 13, 3, 13, 4);
		MergeCells(table2, 13, 4, 13, 5);
		MergeCells(table2, 13, 5, 13, 6);
		MergeCells(table2, 14, 3, 14, 4);
		MergeCells(table2, 14, 4, 14, 5);
		MergeCells(table2, 14, 5, 14, 6);
		MergeCells(table2, 15, 3, 15, 4);
		MergeCells(table2, 15, 4, 15, 5);
		MergeCells(table2, 15, 5, 15, 6);

		setAlignment(1);

		QString LogoContent = QStringLiteral("春江潮水连海平，海上明月共潮生。");//标志内容
		QString LogoSharpness = QStringLiteral("滟滟随波千万里，何处春江无月明！");//标志清晰度
		QString LogoSpacing = QStringLiteral("江流宛转绕芳甸，月照花林皆似霰；");//标志间距
		QString CabOuterDiameter = QStringLiteral("空里流霜不觉飞，汀上白沙看不见。");//电缆外径
		QString NomCabOuterThickness = QStringLiteral("%1").arg(NomMinConfig->NomOST);//标称外护层厚度
		QString MinCabOuterThickness = QStringLiteral("%1").arg(NomMinConfig->MinOST);//最小外护层厚度
		QString AvgCabOuterThickness = QStringLiteral("江畔何人");//外护层厚度平 均 厚 度
		QString MinOuterThickness = QStringLiteral("江月何年");//外护层厚度最 小 厚 度
		QString CabOuterThickness1 = QStringLiteral("121");//外护层厚度1
		QString CabOuterThickness2 = QStringLiteral("121");//外护层厚度2
		QString CabOuterThickness3 = QStringLiteral("1234");//外护层厚度3
		QString CabOuterThickness4= QStringLiteral("121");//外护层厚度 4
		QString CabOuterThickness5 = QStringLiteral("121");//外护层厚度5
		QString CabOuterThickness6 = QStringLiteral("1234");//外护层厚度6

		//QString OuterThickness = QStringLiteral("人生代代无穷已，江月年年望相似。");//铠装金属最小厚度
		QString MinThickness = QStringLiteral("%1").arg(NomMinConfig->MinAMSThickness);//铠装金属最小厚度
		QString MinDiamete = QStringLiteral("%1").arg(NomMinConfig->MinAMSDiameter);//铠装金属最小直径
		//QString OuterThickness = QStringLiteral("人生代代无穷已，江月年年望相似。");
		QString MaximumWrapGap = QStringLiteral("%1").arg(NomMinConfig->MaxAMSWrapGap);//铠装金属最大绕包间隙：
		QString ArmorOuterDiameter = QStringLiteral("不应有恨");  //铠装外径
		QString structure = QStringLiteral("何事长向");//结构：
		QString ArmoredMaxWrapGap = QStringLiteral("人有悲欢离合"); //铠装金属带最大绕包间隙： 
		QString ArmoredThickpoint = QStringLiteral("月有阴晴圆缺"); //铠装金属带最薄点厚度
		QString ArmoredMinimumDiameter = QStringLiteral("此事古难全"); //铠装金属丝最小直径

		QString NomLinerOuterDiameter = QStringLiteral("%1").arg(NomMinConfig->NomLOD);//标称内衬层外径
		QString MinWrapTapeODiameter = QStringLiteral("%1").arg(NomMinConfig->MinLOD);//最小内衬层外径
		QString CabOuterThickness7 = QStringLiteral("121");//内衬层外径厚度1
		QString CabOuterThickness8 = QStringLiteral("121");//内衬层外径厚度2
		QString CabOuterThickness9 = QStringLiteral("1234");//内衬层外径厚度3
		QString CabOuterThickness10 = QStringLiteral("121");//内衬层外径厚度 4
		QString CabOuterThickness11 = QStringLiteral("121");//内衬层外径厚度5
		QString CabOuterThickness12 = QStringLiteral("1234");//内衬层外径厚度6
		QString AvgCabOuterThickness1 = QStringLiteral("34234");//内衬层外径度平 均
		QString MinOuterThickness1 = QStringLiteral("12321");//内衬层外径最 小
		QString WrapTapeouterDiameter = QStringLiteral("GB");//绕包带外径
		//QString InsulationSign = QStringLiteral("GBSign");//绝缘标志
		QString ASign = QStringLiteral("A");//绝缘标志
		QString BSign = QStringLiteral("B");//绝缘标志
		QString CSign = QStringLiteral("C");//绝缘标志
		QString InsulationOuterDiameter = QStringLiteral("GBDiameter");//绝缘外径
		QString AOuterDiameter = QStringLiteral("A");//绝缘标志
		QString BOuterDiameter = QStringLiteral("B");//绝缘标志
		QString COuterDiameter = QStringLiteral("C");//绝缘标志Conductor outer diameter
		QString AConductorOuterDiameter = QStringLiteral("AConductor");//导体外径1
		QString BConductorOuterDiameter = QStringLiteral("BConductor");//导体外径2
		QString CConductorOuterDiameter = QStringLiteral("CConductor");//导体外径3
		QString AsingleWiresNumber = QStringLiteral("ANumber");//绝缘标志
		QString BsingleWiresNumber = QStringLiteral("BNumber");//绝缘标志
		QString CsingleWiresNumber = QStringLiteral("CNumber");//绝缘标志Number of singleWiresNumber Insulation thickness
		QString NomInsulationThickness = QStringLiteral("%1").arg(NomMinConfig->NomIT);
		QString MinInsulationThickness = QStringLiteral("%1").arg(NomMinConfig->MinIT);
		//MergeCells(table2, 9, 2, 10, 2);
		//table2->querySubObject("ParagraphFormat")->setProperty("Alignment", AlignParCenter);
		//table2->querySubObject("Cells")->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");//垂直居中
		typeText(QStringLiteral("项   目"));
		moveRight();
		typeText(QStringLiteral("单位"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("标 准 要 求"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("实        测       值"));
		setAlignment(1);
		moveRight();
		moveRight();
		typeText(QStringLiteral("标志内容"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("—"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("厂名、型号、规格"));
		setAlignment(0);
		moveRight();
		typeText(QStringLiteral("%1").arg(LogoContent));
		setAlignment(1);
		moveRight();
		moveRight();
		typeText(QStringLiteral("标志清晰"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("—"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("字迹应清晰，容易辨认，耐擦"));
		setAlignment(0);
		moveRight();
		typeText(QStringLiteral("%1").arg(LogoSharpness));
		setAlignment(1);
		moveRight();
		moveRight();
		typeText(QStringLiteral("标志间距"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("mm"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("≤500"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("%1").arg(LogoSpacing));
		setAlignment(1);
		moveRight();
		moveRight();
		typeText(QStringLiteral("电缆外径"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("mm"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("—"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("%1").arg(CabOuterDiameter));
		setAlignment(1);
		moveRight();
		moveRight();
		typeText(QStringLiteral("外护层厚度"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("mm"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("标称：%1").arg(NomCabOuterThickness));
		setAlignment(0);
		insertEnter();
		typeText(QStringLiteral("最小：%1").arg(MinCabOuterThickness));
		setAlignment(0);
		moveRight();
		typeText(QStringLiteral("1. %1  2. %2  3. %3 4. %4 5. %5  6. %6").arg(CabOuterThickness1).arg(CabOuterThickness2).arg(CabOuterThickness3).arg(CabOuterThickness4).arg(CabOuterThickness5).arg(CabOuterThickness6));
		setAlignment(0);
		moveRight();
		moveRight();
		typeText(QStringLiteral("平 均 厚 度：%1").arg(AvgCabOuterThickness));
		setAlignment(0);
		moveRight();
		typeText(QStringLiteral("最 小 厚 度：%1").arg(MinOuterThickness));
		setAlignment(0);
		moveRight();
		moveRight();
		typeText(QStringLiteral("铠装金属带□"));
		setAlignment(0);
		insertEnter(); insertEnter(); insertEnter();
		typeText(QStringLiteral("铠装金属丝□"));
		moveRight();
		typeText(QStringLiteral("mm"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("最小厚度：%1").arg(MinThickness));
		setAlignment(0);
		insertEnter();
		typeText(QStringLiteral("最大绕包间隙：%1").arg(MaximumWrapGap));
		//setAlignment(0);
		insertEnter();
		typeText(QStringLiteral("最小直径：%1").arg(MinDiamete));
		moveRight();
		typeText(QStringLiteral("铠装外径：%1         结构：%2").arg(ArmorOuterDiameter).arg(structure));
		setAlignment(0);
		insertEnter();
		typeText(QStringLiteral("铠装金属带最大绕包间隙：%1").arg(ArmoredMaxWrapGap));
		insertEnter();
		typeText(QStringLiteral("铠装金属带最薄点厚度：%1").arg(ArmoredThickpoint));
		insertEnter();
		typeText(QStringLiteral("铠装金属丝最小直径：%1").arg(ArmoredMinimumDiameter));
		moveRight();
		moveRight();
		insertEnter();
		typeText(QStringLiteral("内衬层外径"));
		setAlignment(1);
		insertEnter();
		typeText(QStringLiteral("结     构"));
		insertEnter();
		typeText(QStringLiteral("厚     度"));
		insertEnter();
		moveRight();
		typeText(QStringLiteral("mm"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("挤包□"));
		setAlignment(0);
		insertEnter();
		typeText(QStringLiteral("绕包□"));
		insertEnter();
		typeText(QStringLiteral("包带加挤包□"));
		moveRight();
		typeText(QStringLiteral("%1").arg(ArmoredThickpoint));
		setAlignment(1);
        moveRight();
        moveRight();
		//insertEnter();
		typeText(QStringLiteral("标称：%1").arg(NomLinerOuterDiameter));
		insertEnter();
		typeText(QStringLiteral("最小：%1").arg(MinWrapTapeODiameter));
		moveRight();
		typeText(QStringLiteral("1. %1  2. %2  3. %3 4. %4 5. %5  6. %6").arg(CabOuterThickness7).arg(CabOuterThickness8).arg(CabOuterThickness9).arg(CabOuterThickness10).arg(CabOuterThickness11).arg(CabOuterThickness12));
		setAlignment(0);
		moveRight();
		moveRight();
		typeText(QStringLiteral("平 均 厚 度：%1").arg(AvgCabOuterThickness1));
		setAlignment(0);
		moveRight();
		typeText(QStringLiteral("最 小 厚 度：%1").arg(MinOuterThickness1));
		setAlignment(0);
		moveRight();
		moveRight();
		typeText(QStringLiteral("绕包带外径"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("mm"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("—"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("%1").arg(WrapTapeouterDiameter));
		setAlignment(1);
		moveRight();
		moveRight();
		typeText(QStringLiteral("绝缘标志"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("—"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("颜色/数字"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("%1相").arg(ASign));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("%1相").arg(BSign));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("%1相").arg(CSign));
		setAlignment(1);
		moveRight();
		moveRight();
		typeText(QStringLiteral("绝缘外径"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("mm"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("—"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("%1").arg(AOuterDiameter));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("%1").arg(BOuterDiameter));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("%1").arg(COuterDiameter));
		setAlignment(1);
		moveRight();
		moveRight();
		typeText(QStringLiteral("导体外径"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("mm"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("—"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("%1").arg(AConductorOuterDiameter));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("%1").arg(BConductorOuterDiameter));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("%1").arg(CConductorOuterDiameter));
		setAlignment(1);
		moveRight();
		moveRight();
		typeText(QStringLiteral("单线根数"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("根"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("%1").arg(NomMinConfig->NumberSWN));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("%1").arg(AsingleWiresNumber));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("%1").arg(CsingleWiresNumber));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("%1").arg(BsingleWiresNumber));
		setAlignment(1); 
		moveRight();
		moveRight();
		typeText(QStringLiteral("（转背面）"));
		setAlignment(2);
		moveToEnd(table2);
		insertEnter();
		insertEnter();
		typeText(QStringLiteral("文件代号：CEPRI-D-EETC08-JS-701/1 （接正面）"));
		setAlignment(0);
		insertEnter();

		setAlignment(0);
		QAxObject *table3 = new QAxObject;
		table3 = createTable(13, 9);
		setColumnWidth(table3, 0, 70);
		setColumnWidth(table3, 1, 30);
		setColumnWidth(table3, 2, 90);
		setColumnWidth(table3, 3, 37);
		setColumnWidth(table3, 4, 37);
		setColumnWidth(table3, 5, 37);
		setColumnWidth(table3, 6, 37);
		setColumnWidth(table3, 7, 37);
		setColumnWidth(table3, 8, 37);

		MergeCells(table3, 0, 3, 0, 4);
		MergeCells(table3, 0, 4, 0, 5);
		MergeCells(table3, 0, 5, 0, 6);
		MergeCells(table3, 1, 3, 1, 4);
		MergeCells(table3, 1, 4, 1, 5);
		MergeCells(table3, 1, 5, 1, 6);
		MergeCells(table3, 2, 3, 2, 4);
		MergeCells(table3, 2, 4, 2, 5);
		MergeCells(table3, 2, 5, 2, 6);
		MergeCells(table3, 3, 3, 3, 4);
		MergeCells(table3, 3, 4, 3, 5);
		MergeCells(table3, 3, 5, 3, 6);
		MergeCells(table3, 4, 3, 4, 4);
		MergeCells(table3, 4, 4, 4, 5);
		MergeCells(table3, 4, 5, 4, 6);
		MergeCells(table3, 5, 3, 5, 4);
		MergeCells(table3, 5, 4, 5, 5);
		MergeCells(table3, 5, 5, 5, 6);
		MergeCells(table3, 6, 3, 6, 4);
		MergeCells(table3, 6, 4, 6, 5);
		MergeCells(table3, 6, 5, 6, 6);
		MergeCells(table3, 7, 3, 7, 4);
		MergeCells(table3, 7, 4, 7, 5);
		MergeCells(table3, 7, 5, 7, 6);
		MergeCells(table3, 8, 3, 8, 4);
		MergeCells(table3, 8, 4, 8, 5);
		MergeCells(table3, 8, 5, 8, 6);
		MergeCells(table3, 9, 3, 9, 4);
		MergeCells(table3, 9, 4, 9, 5);
		MergeCells(table3, 9, 5, 9, 6);
		MergeCells(table3, 10, 3, 10, 4);
		MergeCells(table3, 10, 4, 10, 5);
		MergeCells(table3, 10, 5, 10, 6);
		MergeCells(table3, 11, 3, 11, 4);
		MergeCells(table3, 11, 4, 11, 5);
		MergeCells(table3, 11, 5, 11, 6);
		MergeCells(table3, 12, 3, 12, 4);
		MergeCells(table3, 12, 4, 12, 5);
		MergeCells(table3, 12, 5, 12, 6);
		setColumnWidth(table3, 3, 20);
		setColumnWidth(table3, 4, 101);
		setColumnWidth(table3, 5, 101);
		MergeCells(table3, 0, 4, 0, 5);
		MergeCells(table3, 1, 4, 1, 5);
		MergeCells(table3, 2, 4, 2, 5);
		MergeCells(table3, 3, 4, 3, 5);
		MergeCells(table3, 5, 4, 5, 5);
		MergeCells(table3, 6, 4, 6, 5);
		MergeCells(table3, 7, 4, 7, 5);
		MergeCells(table3, 9, 4, 9, 5);
		MergeCells(table3, 10, 4, 10, 5);
		MergeCells(table3, 11, 4, 11, 5);
		MergeCells(table3, 1, 3, 4, 3);
		MergeCells(table3, 5, 3, 8, 3);
		MergeCells(table3, 9, 3, 12, 3);
		MergeCells(table3, 1, 0, 12, 0);
		MergeCells(table3, 1, 1, 12, 1);
		MergeCells(table3, 1, 2, 12, 2);
		MergeCells(table3, 0, 3, 0, 4);

		for (int i = 2; i < 5; i++)
		{
			if (i % 2 != 0)
			{
				continue;
			}
			QString m_str = QString("Borders(-%1)").arg(i);
			QAxObject *borders = table->querySubObject(m_str.toLatin1().constData());
			borders->dynamicCall("SetLineStyle(int)", 0);
			QAxObject *borders2 = table2->querySubObject(m_str.toLatin1().constData());
			borders2->dynamicCall("SetLineStyle(int)", 0);
			QAxObject *borders3 = table3->querySubObject(m_str.toLatin1().constData());
			borders3->dynamicCall("SetLineStyle(int)", 0);
		}

		typeText(QStringLiteral("项   目"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("单位"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("标 准 要 求"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("实        测       值"));
		setAlignment(1);
		moveRight();
		moveRight();
		typeText(QStringLiteral("绝缘厚度"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("mm"));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("标称：%1").arg(NomInsulationThickness));
		setAlignment(0);
		insertEnter();
		typeText(QStringLiteral("最小：%1").arg(MinInsulationThickness));
		moveRight();
		insertEnter();
		typeText(QStringLiteral("%1相").arg(ASign));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("1.%1                     2.%2").arg(ASign).arg(ASign));
		moveRight();
		moveRight();
		//typeText(QStringLiteral("%1相").arg(ASign));
		//setAlignment(1);
		//moveRight();
		typeText(QStringLiteral("3.%1                     4.%2").arg(ASign).arg(ASign));
		moveRight();
		moveRight();
		//typeText(QStringLiteral("%1相").arg(ASign));
		//setAlignment(1);
		//moveRight();
		typeText(QStringLiteral("5.%1                     6.%2").arg(ASign).arg(ASign));
		moveRight();
		moveRight();
		typeText(QStringLiteral("平 均 厚 度：%1").arg(AvgCabOuterThickness1));
		setAlignment(0);
		moveRight();
		typeText(QStringLiteral("最 小 厚 度：%1").arg(MinOuterThickness1));
		setAlignment(0);
		moveRight();
		moveRight();
		insertEnter();
		typeText(QStringLiteral("%1相").arg(BSign));
        setAlignment(1);
        moveRight();
		typeText(QStringLiteral("1.%1                     2.%2").arg(BSign).arg(BSign));
		moveRight();
		moveRight();
		typeText(QStringLiteral("3.%1                     4.%2").arg(BSign).arg(BSign));
		moveRight();
		moveRight();
		typeText(QStringLiteral("5.%1                     6.%2").arg(BSign).arg(BSign));
		moveRight();
		moveRight();
		typeText(QStringLiteral("平 均 厚 度：%1").arg(AvgCabOuterThickness1));
		setAlignment(0);
		moveRight();
		typeText(QStringLiteral("最 小 厚 度：%1").arg(MinOuterThickness1));
		moveRight();
		moveRight();
		insertEnter();
		typeText(QStringLiteral("%1相").arg(CSign));
		setAlignment(1);
		moveRight();
		typeText(QStringLiteral("1.%1                     2.%2").arg(CSign).arg(CSign));
		moveRight();
		moveRight();
		typeText(QStringLiteral("3.%1                     4.%2").arg(CSign).arg(CSign));
		moveRight();
		moveRight();
		typeText(QStringLiteral("5.%1                     6.%2").arg(CSign).arg(CSign));
		moveRight();
		moveRight();
		typeText(QStringLiteral("平 均 厚 度：%1").arg(AvgCabOuterThickness1));
		setAlignment(0);
		moveRight();
		typeText(QStringLiteral("最 小 厚 度：%1").arg(MinOuterThickness1));
		moveRight();
		moveRight();

		selection->querySubObject("Font")->setProperty("Size", 11);
		selection->querySubObject("Font")->setProperty("Bold", true);
		selection->querySubObject("Range")->setProperty("Text", QStringLiteral("五、检测后对检测仪器设备及被测样品的检查"));
		selection->querySubObject("ParagraphFormat")->setProperty("Alignment", AlignParLeft); //文本位置设置
		moveToEnd();
		insertEnter();
		selection->querySubObject("Font")->setProperty("Size", 10);
		selection->querySubObject("Font")->setProperty("Bold", false);
		selection->querySubObject("Range")->setProperty("Text",
			QStringLiteral("    1. 游标卡尺的零位正确 □；"));
		moveToEnd();
		insertEnter();
		selection->querySubObject("Range")->setProperty("Text",
			QStringLiteral("    2. 数显卡尺的零位正确 □；"));
		moveToEnd();
		insertEnter();
		//insertEnter();

	}

}

bool OutputExcel::addText(QString titlestr, TITLE_NUMBER number /*= NORMAL*/, WdParagraphAlignment alignment /*= AlignParLeft*/)
{
	if (!m_bOpened) return false;
	QAxObject *selection = NULL;
	selection = m_wordWidget->querySubObject("Selection");
	if (selection)
	{
		selection->querySubObject("Range")->setProperty("Text", titlestr);
		selection->querySubObject("Range")->dynamicCall("SetStyle(QVariant)", getTitleStr(number));
		selection->querySubObject("ParagraphFormat")->setProperty("Alignment",alignment); //文本位置设置

		//背景色
		//selection->querySubObject("Range")
		//	->querySubObject("ParagraphFormat")
		//	->querySubObject("Shading")
		//	->setProperty("BackgroundPatternColor",QColor("blue").value());
		moveToEnd();
		return true;
	}
	return false;
}



bool OutputExcel::addText(QString titlestr, QFont font, QColor fontcolor)
{
	if (!m_bOpened) return false;
	QAxObject *selection = NULL;
	selection = m_wordWidget->querySubObject("Selection");
	if (selection)
	{
		//selection->querySubObject("Range")->querySubObject("Font")->setProperty("Size", fo); //不好用
		//selection->querySubObject("Range")->querySubObject("Font")->dynamicCall("Size", 20);
		selection->querySubObject("Range")->querySubObject("Font")->setProperty("Size", QVariant(font.pointSize()));
		selection->querySubObject("Range")->querySubObject("Font")->setProperty("Color", colorToInt(fontcolor));
		if (font.weight() >= QFont::Bold)
		{
			selection->querySubObject("Range")->querySubObject("Font")->setProperty("Bold", true);
		}
		selection->querySubObject("Range")->setProperty("Text", titlestr); //方式1

		//selection->dynamicCall("TypeText(const QString&)",titlestr);         //方式2 使用方法2不能设置字体
		moveToEnd();
		return true;
	}
	return false;
}



QAxObject* OutputExcel::addText2(QString titlestr)

{
	QAxObject *selection = NULL;
	if (!m_bOpened) return selection;
	selection = m_wordWidget->querySubObject("Selection");
	if (selection)
	{
		selection->querySubObject("Range")->setProperty("Text", titlestr); //方式1
	}
	return selection;
}





bool OutputExcel::insertEnter()
{
	QAxObject *selection = m_wordWidget->querySubObject("Selection");
	if (selection)
	{
		selection->dynamicCall("TypeParagraph(void)");
		return true;
	}
	return false;
}



bool OutputExcel::moveToEnd()
{
	QAxObject *selection = m_wordWidget->querySubObject("Selection");
	QVariantList params;
	params << wdStory << 0;
	if (selection)
	{
		selection->dynamicCall("EndOf(QVariant&, QVariant&)", params);
		return true;
	}
	return false;
}

void OutputExcel::moveRight()
{
	QAxObject* selection = m_wordWidget->querySubObject("Selection");
	if (!selection)
	{
		return;
	}
	selection->dynamicCall("MoveRight(int)", 1);
}


bool OutputExcel::moveToEnd(MOVEEND_INDEX wd)

{
	QAxObject *selection = m_wordWidget->querySubObject("Selection");
	QVariantList params;
	params << wd << 0;
	selection->dynamicCall("EndOf(QVariant&, QVariant&)", params);
	return true;
}



bool OutputExcel::moveToEnd(QAxObject *table)
{
	if (!table) return false;
	moveToEnd(wdTable);
	moveToEnd();
	return true;
}







QAxObject* OutputExcel::createTable(int row, int column)
{
	QAxObject* selection = m_wordWidget->querySubObject("Selection");
	if (!selection) return NULL;
	selection->dynamicCall("InsertAfter(QString&)", "\r\n");
	QAxObject *range = selection->querySubObject("Range");
	QAxObject *tables = m_wordDocuments->querySubObject("Tables");
	QAxObject *table = tables->querySubObject("Add(QVariant,int,int)", range->asVariant(), row, column);
	table->setProperty("Style", "网格型");
	QAxObject* Borders = table->querySubObject("Borders");
	Borders->setProperty("InsideLineStyle", 1);
	Borders->setProperty("OutsideLineStyle", 1);
	//表格自动拉伸列 0固定  1根据内容调整  2 根据窗口调整
	table->dynamicCall("AutoFitBehavior(WdAutoFitBehavior)", 2);
	//for (int i = 0; i < headList.size(); i++)
	//{
	//	table->querySubObject("Cell(int,int)", 1, i + 1)->querySubObject("Range")->dynamicCall("SetText(QString)", headList);
	//	table->querySubObject("Cell(int,int)", 1, i + 1)->querySubObject("Range")->dynamicCall("SetBold(int)", false);
	//}
	return table;
}


void OutputExcel::intsertTable(int row, int column)
{
	QAxObject* tables = m_wordWidget->querySubObject("Tables");
	QAxObject* selection = m_wordWidget->querySubObject("Selection");
	QAxObject* range = selection->querySubObject("Range");
	QVariantList params;
	params.append(range->asVariant());
	params.append(row);
	params.append(column);
	tables->querySubObject("Add(QAxObject*, int, int, QVariant&, QVariant&)", params);
	QAxObject* table = selection->querySubObject("Tables(int)", 1);
	table->setProperty("Style", "网格型");

	QAxObject* Borders = table->querySubObject("Borders");
	Borders->setProperty("InsideLineStyle", 1);
	Borders->setProperty("OutsideLineStyle", 1);

	//QString doc = Borders->generateDocumentation();
	//QFile outFile("D:\\360Downloads\\Picutres\\Borders.html");
	//outFile.open(QIODevice::WriteOnly | QIODevice::Append);
	//QTextStream ts(&outFile);
	//ts << doc << endl;

	/*QString doc = tables->generateDocumentation();
	QFile outFile("D:\\360Downloads\\Picutres\\tables.html");
	outFile.open(QIODevice::WriteOnly|QIODevice::Append);
	QTextStream ts(&outFile);
	ts<<doc<<endl;*/
}



QAxObject* OutputExcel::insertTable(int nStart, int nEnd, int row, int column)
{
	QAxObject* ptst = m_wordDocuments->querySubObject("Range( Long, Long )",
		nStart, nEnd);
	QAxObject* pTables = m_wordDocuments->querySubObject("Tables");
	QVariantList params;
	params.append(ptst->asVariant());
	params.append(row);
	params.append(column);
	if (pTables)
	{
		QAxObject *table = pTables->querySubObject("Add(QAxObject*, Long ,Long )", params);
		table->dynamicCall("AutoFitBehavior(WdAutoFitBehavior)", 2);
		QAxObject* Borders = table->querySubObject("Borders");
		Borders->setProperty("InsideLineStyle", 1);
		Borders->setProperty("OutsideLineStyle", 1);
		return table;
	}
	return NULL;
}



void OutputExcel::setCellString(QAxObject *table, int row, int column, const QString& text)
{
	if (table)
	{
		QAxObject *cell = table->querySubObject("Cell(int, int)", row + 1, column + 1);

		QAxObject *range = table->querySubObject("Range");

		range->dynamicCall("SetText(QString)", text);

		range->dynamicCall("SetBold(int)", false);
	}

}





void OutputExcel::MergeCells(QAxObject *table, int nStartRow, int nStartCol, int nEndRow, int nEndCol)

{

	QAxObject* StartCell = table->querySubObject("Cell(int, int)", nStartRow + 1, nStartCol + 1);

	QAxObject* EndCell = table->querySubObject("Cell(int, int)", nEndRow + 1, nEndCol + 1);

	StartCell->dynamicCall("Merge(LPDISPATCH)", EndCell->asVariant());

}



/******************************************************************************

 * 函数：setColumnWidth

 * 功能：设置表格列宽

 * 参数：table 表格; column 列数; width 宽度

 *****************************************************************************/

void OutputExcel::setColumnWidth(QAxObject *table, int column, int width)

{

	table->querySubObject("Columns(int)", column + 1)->setProperty("Width", width);

}







/******************************************************************************

 * 函数：addTableRow

 * 功能：为表格添加行

 * 参数：table 表格; nRow 插入行; rowCount 插入的行数

 *****************************************************************************/

void OutputExcel::addTableRow(QAxObject *table, int nRow, int rowCount)

{

	QAxObject* rows = table->querySubObject("Rows");

	int Count = rows->dynamicCall("Count").toInt();



	if (nRow > Count)

	{

		nRow = Count;

	}

	QAxObject* row = table->querySubObject("Rows(int)", nRow);

	if (row == NULL)

	{

		row = rows->querySubObject("Last");

	}

	if (0 <= nRow && nRow <= Count)

	{

		for (int i = 0; i < rowCount; ++i)

		{

			rows->dynamicCall("Add(QVariant)", row->asVariant());

		}

	}

}





void OutputExcel::appendTableRow(QAxObject *table, int rowCount)

{

	QAxObject* rows = table->querySubObject("Rows");



	int Count = rows->dynamicCall("Count").toInt();

	QAxObject* row = rows->querySubObject("Last");



	for (int i = 0; i < rowCount; ++i)

	{

		QVariant param = row->asVariant();

		rows->dynamicCall("Add(Variant)", param);

	}

}





/******************************************************************************

 * 函数：setCellFontBold

 * 功能：设置内容粗体  isBold控制是否粗体

 * 参数：table 表格; row 插入行; column 列数; isBold 是否加粗

 *****************************************************************************/

void OutputExcel::setCellFontBold(QAxObject *table, int row, int column, bool isBold)

{

	table->querySubObject("Cell(int, int)", row, column)->querySubObject("Range")

		->dynamicCall("SetBold(int)", isBold);

}



/******************************************************************************

 * 函数：setCellFontSize

 * 功能：设置文字大小

 * 参数：table 表格; row 插入行; column 列数; size 字体大小

 *****************************************************************************/

void OutputExcel::setCellFontSize(QAxObject *table, int row, int column, int size)

{

	table->querySubObject("Cell(int, int)", row, column)->querySubObject("Range")

		->querySubObject("Font")->setProperty("Size", size);

}



/******************************************************************************

 * 函数：insertCellPic

 * 功能：在表格中插入图片

 * 参数：table 表格; row 插入行; column 列数; picPath 图片路径

 *****************************************************************************/

void OutputExcel::insertCellPic(QAxObject *table, int row, int column,

	const QString& picPath)

{

	QAxObject* range = table->querySubObject("Cell(int, int)", row, column)

		->querySubObject("Range");

	range->querySubObject("InlineShapes")

		->dynamicCall("AddPicture(const QString&)", picPath);

}






void OutputExcel::insertPic(QString picPath)

{

	QAxObject *selection = m_wordWidget->querySubObject("Selection");

	selection->querySubObject("ParagraphFormat")->dynamicCall("Alignment", "wdAlignParagraphCenter");

	QVariant tmp = selection->asVariant();

	QList<QVariant>qList;

	qList << QVariant(picPath);

	qList << QVariant(false);

	qList << QVariant(true);

	qList << tmp;

	QAxObject *Inlineshapes = m_wordDocuments->querySubObject("InlineShapes");

	Inlineshapes->dynamicCall("AddPicture(const QString&, QVariant, QVariant ,QVariant)", qList);

}







void OutputExcel::setColor(QColor color)

{

	QAxObject *selection = m_wordWidget->querySubObject("Selection");

	setColor(selection, color);


}



void OutputExcel::setColor(QAxObject *obj, QColor color)

{

	if (!obj) return;

	obj->querySubObject("Range")

		->querySubObject("ParagraphFormat")

		->querySubObject("Shading")

		->setProperty("ForegroundPatternColor", colorToInt(color));

}





void OutputExcel::setBgColor(QAxObject *obj, QColor color)

{

	if (!obj) return;

	obj->querySubObject("Range")

		->querySubObject("ParagraphFormat")

		->querySubObject("Shading")

		->setProperty("BackgroundPatternColor", colorToInt(color));

}







//设置对齐方式

void OutputExcel::setAlignment(int index)

{

	QAxObject *selection = m_wordWidget->querySubObject("Selection");

	if (!selection) return;

	selection->querySubObject("ParagraphFormat")->setProperty("Alignment", index);

}



void OutputExcel::setFontSize(int size)

{

	QAxObject *selection = m_wordWidget->querySubObject("Selection");

	if (!selection) return;

	selection->querySubObject("Font")->setProperty("Size", size);

}



QString OutputExcel::getTitleStr(TITLE_NUMBER number)

{
	QString str;
	switch (number)
	{
	case TITLE_ONE: str = "标题 1"; break;
	case TITLE_TWO: str = "标题 2"; break;
	case TITLE_THREE: str = "标题 3"; break;
	default: str = "正文"; break;
	}
	return str;
}





void OutputExcel::setPropraty(QAxObject *axobj, QString proname, QVariant provalue)
{
	if (!axobj) return;
	axobj->setProperty(proname.toStdString().c_str(), proname);
}



int OutputExcel::colorToInt(QColor color)

{

	int sum = 0;

	int r = color.red() << 16;

	int g = color.green() << 8;

	int b = color.blue();

	int al = color.alpha() << 24;



	sum = al + r + g + b;

	return sum;

}











void OutputExcel::writeFile(QString savestr, QString filename)

{

	QFile savefile(filename);

	savefile.open(QFile::WriteOnly);

	QTextStream saveteam(&savefile);

	saveteam.setCodec("UTF-8");

	saveteam << savestr;

	savefile.close();

}



bool OutputExcel::open(bool bvisable)

{
	m_bOpened = false;
	m_wordWidget = new QAxWidget;
	bool bFlag = m_wordWidget->setControl("Word.Application");

	m_wordWidget->setProperty("Visible", bvisable);

	//获取所有的工作文档
	QAxObject *document = m_wordWidget->querySubObject("Documents");
	if (!document)
	{
		return m_bOpened;
	}
	//新建一个文档页
	document->dynamicCall("Add()");
	//获取当前激活的文档
	m_wordDocuments = m_wordWidget->querySubObject("ActiveDocument");

	if (m_wordDocuments)

		m_bOpened = true;
	else
		m_bOpened = false;
	return m_bOpened;
}



bool OutputExcel::open(const QString& strFile, bool bVisable /*= false*/)

{
	m_filename = strFile;
	close();
	return open(bVisable);
}



bool OutputExcel::close()
{
	if (m_bOpened)
	{
		if (m_wordDocuments)
		{
			m_wordDocuments->dynamicCall("Close (boolean)", true);
		}
		if (m_wordWidget)
		{
			m_wordWidget->dynamicCall("Quit()");//退出word
			m_wordWidget->close();
		}
		if (m_wordDocuments)
			delete m_wordDocuments;
		if (m_wordWidget)
			delete m_wordWidget;
		m_bOpened = false;
	}
	return m_bOpened;
}



bool OutputExcel::isOpen()

{
	return m_bOpened;
}



void OutputExcel::save()
{
	//QDir dir1("./doc");
	//QFileInfoList info_list = dir1.entryInfoList(QDir::Dirs | QDir::Files | QDir::NoDotAndDotDot);
	//QListIterator<QFileInfo> i(info_list);
	//QStringList doc_list, pdf_list;
	//while (i.hasNext())
	//{
	//	QFileInfo info = i.next();
	//	if (info.isFile())
	//	{
	//		if ("docx" == info.suffix())
	//		{
	//			if (ui.checkBox->text() == info.fileName())
	//			{
	//				QMessageBox::StandardButton re = QMessageBox::warning(this, QStringLiteral("提示"), QStringLiteral("兄dei,另存为不？"), QMessageBox::Yes | QMessageBox::No, QMessageBox::No);
	//				if (re== QMessageBox::Yes)
	//				{
	//					QVariant newFileName(m_filename);//保存路径及名称	
	//					QVariant fileFormat(1);//文件格式	
	//					m_wordDocuments->dynamicCall("SaveAs(const QVariant&, const QVariant&)", newFileName, fileFormat);	m_wordDocuments->dynamicCall("Close (boolean)", true);
	//				}
	//				else
	//				{
	//					return;
	//				}
	//			}
	//		}
	//	}
	//}
	QDir dir;
	QString dstPath = dir.currentPath() + "/doc/"+QStringLiteral("%1").arg(ui.checkBox->text())+".docx";
	QVariant newFileName(dstPath);//保存路径及名称
	//QVariant fileFormat(1);//文件格式
	m_wordDocuments->dynamicCall("SaveAs(const QVariant&)", newFileName);
	m_wordDocuments->dynamicCall("Close (boolean)", true);
}



void OutputExcel::saveAs(const QString& strSaveFile)
{
	//return m_wordDocuments->dynamicCall("SaveAs (const QString&)",
	//	strSaveFile).toBool();
	QVariant newFileName(m_filename);//保存路径及名称	
	QVariant fileFormat(1);//文件格式	
	m_wordDocuments->dynamicCall("SaveAs(const QVariant&, const QVariant&)", newFileName, fileFormat);	m_wordDocuments->dynamicCall("Close (boolean)", true);

}


void OutputExcel::typeText(QString text)
{
	QAxObject* selection = m_wordWidget->querySubObject("Selection");
	if (!selection)
	{
		return;
	}
	selection->dynamicCall("TypeText(const QString&)", text);
}

void OutputExcel::on_BnOverLoad_clicked()
{
	ui.GB_comboBox->clear();
	readConfig();
}

void OutputExcel::readConfig()
{
	QSettings set("./Config/LowVoltageCable/TestingEquipment.ini", QSettings::IniFormat);
	set.setIniCodec("UTF8");
	//QSettings set("./Config/GBT1270612008.ini", QSettings::IniFormat);
	set.beginGroup("TestingEquipment");
	//QStringList GBConfigList1 = set.allKeys();
	QString DocumentCode = set.value("DocumentCode").toString();//文件代号
	QString DetectionLocation1 = set.value("DetectionLocation").toString();//检测地点
	QString PragraphOne = set.value("PragraphOne").toString();
	QString PragraphTwo = set.value("PragraphTwo").toString();
	QString PragraphThree = set.value("PragraphThree").toString();
	QString PragraphFour = set.value("PragraphFour").toString();
	ui.lineEdit_10->setText(DocumentCode);
	ui.textEdit->setText(PragraphOne);
	ui.textEdit_2->setText(PragraphTwo);
	ui.textEdit_3->setText(PragraphThree);
	ui.textEdit_4->setText(PragraphFour);
	ui.lineEdit_26->setText(DetectionLocation1);


	configBlack = new ConfigPramerBlack;
	configBlack->DocumentCode = DocumentCode;
	configBlack->DetectionLocation = DetectionLocation1;
	configBlack->PragraphOne = PragraphOne;
	configBlack->PragraphTwo = PragraphTwo;
	configBlack->PragraphThree = PragraphThree;
	configBlack->PragraphFour = PragraphFour;

	set.endGroup();
	QSettings GBConfigSet("./Config/LowVoltageCable/GBTConfig.ini", QSettings::IniFormat);
	GBConfigSet.setIniCodec("UTF8");
	GBConfigSet.beginGroup("GBTConfig");
	QStringList GBConfigList = GBConfigSet.allKeys();
	//ui.GB_comboBox->addItems(GBConfigList);
	for (int i = 0; i < GBConfigList.size(); i++)
	{
		ui.GB_comboBox->addItem(GBConfigSet.value(GBConfigList[i]).toString());
	}
	GBConfigSet.endGroup();
	connect(ui.GB_comboBox, SIGNAL(currentIndexChanged(int)), this, SLOT(LoadConfig()));
	//LoadConfig(ui.GB_comboBox);
}

void OutputExcel::LoadConfig()
{
	laodconfig = false;
	QComboBox *comboxName = (QComboBox*)sender();
	QString CBtext = comboxName->currentText();
	//qDebug() << CBtext;
	CBtext.remove(QChar('/'), Qt::CaseInsensitive);
	CBtext.remove(QChar('.'), Qt::CaseInsensitive);
	//CBtext.remove(QString('—'), Qt::CaseInsensitive);
	QString fliePath = QString("./Config/LowVoltageCable/%1.ini").arg(CBtext);
	QSettings GBT(fliePath, QSettings::IniFormat);
	GBT.setIniCodec("UTF8");
	QStringList GBTcomboxList = GBT.childGroups();
	//clearButton();
	NomMinConfig = new NomMinConfigPramer;

	for (int i = 0; i < GBTcomboxList.size(); i++)
	{
		if ("OuterSheathThickness" == GBTcomboxList[i])
		{
			GBT.beginGroup("OuterSheathThickness");
			NomMinConfig->NomOST = GBT.value("NomOST").toDouble();
            NomMinConfig->MinOST = GBT.value("MinOST").toDouble();
			//QPushButton *pushButtonSix = new QPushButton(QStringLiteral("外护层厚度"), ui.widget);
			//pushButtonSix->move(60, 70);
			//pushButtonSix->show();
			GBT.endGroup();
			//connect(pushButtonSix, SIGNAL(clicked()), this, SLOT(ShowConfig()));
			ui.lineEdit_13->setText(QString::number(NomMinConfig->NomOST));
			ui.lineEdit_17->setText(QString::number(NomMinConfig->MinOST));

		}
		if ("ArmoredMetalStrip" == GBTcomboxList[i])
		{
			GBT.beginGroup("ArmoredMetalStrip");
			NomMinConfig->MinAMSThickness = GBT.value("MinAMSThickness").toDouble();
			NomMinConfig->MaxAMSWrapGap = GBT.value("MaxAMSWrapGap").toDouble();
			NomMinConfig->MinAMSDiameter = GBT.value("MinAMSDiameter").toDouble();
			//QPushButton *pushButtonOne = new QPushButton(QStringLiteral("铠装金属带"), ui.widget);
			//pushButtonOne->move(60, 110);
			//pushButtonOne->show();
			GBT.endGroup();
			//connect(pushButtonOne, SIGNAL(clicked()), this, SLOT(ShowConfig()));
			ui.lineEdit_18->setText(QString::number(NomMinConfig->MinAMSThickness));
			ui.lineEdit_14->setText(QString::number(NomMinConfig->MaxAMSWrapGap));
			ui.lineEdit_19->setText(QString::number(NomMinConfig->MinAMSDiameter));

		}
		if ("ArmoredWire" == GBTcomboxList[i])
		{
			GBT.beginGroup("ArmoredWire");
			NomMinConfig->MinAWThickness = GBT.value("MinAWThickness").toDouble();
			NomMinConfig->MaxAWrapGap = GBT.value("MaxAWrapGap").toDouble();
			NomMinConfig->MinAWDiameter = GBT.value("MinAWDiameter").toDouble();
			//QPushButton *pushButtonTwo = new QPushButton(QStringLiteral("铠装金属丝"), ui.widget);
			//pushButtonTwo->move(60, 150);
			//pushButtonTwo->show();
			GBT.endGroup();
			//connect(pushButtonTwo, SIGNAL(clicked()), this, SLOT(ShowConfig()));

		}
		if ("LinerOuterDiameter" == GBTcomboxList[i])
		{
			GBT.beginGroup("LinerOuterDiameter");
			NomMinConfig->NomLOD = GBT.value("NomLOD").toDouble();
			NomMinConfig->MinLOD = GBT.value("MinLOD").toDouble();
			//QPushButton *pushButtonThree = new QPushButton(QStringLiteral("内衬层外径"), ui.widget);
			//pushButtonThree->move(60, 190);
			//pushButtonThree->show();
			GBT.endGroup();
			//connect(pushButtonThree, SIGNAL(clicked()), this, SLOT(ShowConfig()));
			ui.lineEdit_20->setText(QString::number(NomMinConfig->NomIT));
			ui.lineEdit_21->setText(QString::number(NomMinConfig->MinIT));

		}
		if ("InsulationThickness" == GBTcomboxList[i])
		{
			GBT.beginGroup("InsulationThickness");
			NomMinConfig->NomIT = GBT.value("NomIT").toDouble();
			NomMinConfig->MinIT = GBT.value("MinIT").toDouble();
			//QPushButton *pushButtonFour = new QPushButton(QStringLiteral("绝缘厚度"), ui.widget);
			//pushButtonFour->move(60, 230);
			//pushButtonFour->show();
			GBT.endGroup();
			//connect(pushButtonFour, SIGNAL(clicked()), this, SLOT(ShowConfig()));
			ui.lineEdit_22->setText(QString::number(NomMinConfig->NomIT));
			ui.lineEdit_23->setText(QString::number(NomMinConfig->MinIT));
		}
		if ("SingleWiresNumber" == GBTcomboxList[i])
		{
			GBT.beginGroup("SingleWiresNumber");
			NomMinConfig->NumberSWN = GBT.value("NumberSWN").toDouble();
			//QPushButton *pushButtonFive = new QPushButton(QStringLiteral("单线根数"), ui.widget);
			//pushButtonFive->move(60, 270);
			//pushButtonFive->show();
			GBT.endGroup();
			//connect(pushButtonFive, SIGNAL(clicked()), this, SLOT(ShowConfig()));
			//emit ConfigSignal(GBTcomboxList[i]);
			ui.lineEdit_16->setText(QString::number(NomMinConfig->NumberSWN));
		}
		if ("GBName" == GBTcomboxList[i])
		{
			GBT.beginGroup("GBName");
			NomMinConfig->GBTName = GBT.value("GBnameString").toString();
			//QPushButton *pushButtonFive = new QPushButton(QStringLiteral("单线根数"), ui.widget);
			//pushButtonFive->move(60, 270);
			//pushButtonFive->show();
			GBT.endGroup();
			//connect(pushButtonFive, SIGNAL(clicked()), this, SLOT(ShowConfig()));
			//emit ConfigSignal(GBTcomboxList[i]);
			//ui.lineEdit_16->setText(QString::number(NomMinConfig->NumberSWN));
		}
	}
	laodconfig = true;
}

void OutputExcel::clearConfig()
{
	ui.lineEdit_6->clear();
	ui.lineEdit_2->clear();
	ui.lineEdit_7->clear();
	ui.lineEdit_8->clear();
	ui.lineEdit_10->clear();
	ui.lineEdit_9->clear();
}

//void OutputExcel::clearButton()
//{
//	QList<QPushButton*> btns = ui.widget->findChildren<QPushButton*>();
//	foreach(QPushButton* btn, btns) 
//	{ 
//		delete btn; 
//	}
//}

//
void OutputExcel::ShowConfig()
{
	QPushButton *pushButtonName = (QPushButton*)sender();
	if (pushButtonName->text() == QStringLiteral("单线根数"))
	{
		clearConfig();
		ui.lineEdit_9->setText(QString::number(NomMinConfig->NumberSWN));
	}
	if (pushButtonName->text() == QStringLiteral("绝缘厚度"))
	{
		clearConfig();
		ui.lineEdit_2->setText(QString::number(NomMinConfig->NomIT));
		ui.lineEdit_6->setText(QString::number(NomMinConfig->MinIT));
	}
	if (pushButtonName->text() == QStringLiteral("内衬层外径"))
	{
		clearConfig();
		ui.lineEdit_2->setText(QString::number(NomMinConfig->NomLOD));
		ui.lineEdit_6->setText(QString::number(NomMinConfig->MinLOD));
	}
	if (pushButtonName->text() == QStringLiteral("铠装金属丝"))
	{
		clearConfig();
		ui.lineEdit_8->setText(QString::number(NomMinConfig->MinAMSThickness));
		ui.lineEdit_10->setText(QString::number(NomMinConfig->MaxAMSWrapGap));
		ui.lineEdit_7->setText(QString::number(NomMinConfig->MinAMSDiameter));
	}
	if (pushButtonName->text() == QStringLiteral("铠装金属带"))
	{
		clearConfig();
		ui.lineEdit_8->setText(QString::number(NomMinConfig->MinAMSThickness));
		ui.lineEdit_10->setText(QString::number(NomMinConfig->MaxAMSWrapGap));
		ui.lineEdit_7->setText(QString::number(NomMinConfig->MinAMSDiameter));
	}
	if (pushButtonName->text() == QStringLiteral("外护层厚度"))
	{
		clearConfig();
		ui.lineEdit_2->setText(QString::number(NomMinConfig->NomOST));
		ui.lineEdit_6->setText(QString::number(NomMinConfig->MinOST));
	}
}


void OutputExcel::on_pushButton_clicked()
{
	QString GBTJoinName = ui.lineEdit_11->text();
	GBTJoinName.remove(QChar('/'), Qt::CaseInsensitive);
	GBTJoinName.remove(QChar('.'), Qt::CaseInsensitive);
	QDir dir("./Config/LowVoltageCable");
	QFileInfoList info_list = dir.entryInfoList(QDir::Dirs | QDir::Files | QDir::NoDotAndDotDot);
	QListIterator<QFileInfo> i(info_list);
	QStringList dir_list, file_list;
	while (i.hasNext())
	{
		QFileInfo info = i.next();
		if (info.isFile())
		{
			if (GBTJoinName == info.fileName())
			{
				QMessageBox::information(this, QStringLiteral("提示"), QStringLiteral("已经存在了！换名字吧"));
				return;
			}
		}
	}
	QString fliePath = QString("./Config/LowVoltageCable/GBTConfig.ini");
	QSettings GBTadd("./Config/LowVoltageCable/GBTConfig.ini", QSettings::IniFormat);
	GBTadd.setIniCodec("UTF8");
	QStringList GBTcomboxList = GBTadd.allKeys();
	GBTadd.beginGroup("GBTConfig");
	GBTadd.setValue(QString("GT%1").arg(GBTcomboxList.size()+1), GBTJoinName);
	GBTadd.endGroup();

	//QString pathNewConfig = QString("./Config/%1.ini").arg(GBTJoinName);
	QDir configdir;
	QString pathNewConfig = configdir.currentPath() + "/Config/" + QString("%1").arg(GBTJoinName) + ".ini";
	QFile NewConfig(pathNewConfig);
	NewConfig.open(QIODevice::WriteOnly);

	//添加内容
	NewConfig.close();
}

void OutputExcel::on_BnConfigJion_clicked()
{
	bool docStaes = false;
	bool nomStaes = false;
	QMessageBox::StandardButton rb=QMessageBox::warning(this, QStringLiteral("提示"), QStringLiteral("小老弟想清楚了！确定要重新配置参数？"), QMessageBox::Yes | QMessageBox::No, QMessageBox::No);
	if (rb == QMessageBox::Yes)
	{
		QMessageBox::StandardButton ra = QMessageBox::warning(this, QStringLiteral("提示"), QStringLiteral("兄dei,要配置文件信息参数嘛？"), QMessageBox::Yes | QMessageBox::No, QMessageBox::No);
		if (ra == QMessageBox::Yes)
		{
			docStaes = true;
		}
		QMessageBox::StandardButton rc = QMessageBox::warning(this, QStringLiteral("提示"), QStringLiteral("兄dei,要配置标准信息参数嘛？"), QMessageBox::Yes | QMessageBox::No, QMessageBox::No);
		if (rc == QMessageBox::Yes)
		{
			nomStaes = true;
		}

	}
	else
	{
		return;
	}

	if (docStaes)
	{
		QString strOne = ui.textEdit->toPlainText();
		QString strTwo = ui.textEdit_2->toPlainText();
		QString strThree = ui.textEdit_3->toPlainText();
		QString strFour = ui.textEdit_4->toPlainText();


		QSettings set("./Config/LowVoltageCable/TestingEquipment.ini", QSettings::IniFormat);
		set.setIniCodec("UTF8");
		//QSettings set("./Config/GBT1270612008.ini", QSettings::IniFormat);
		set.beginGroup("TestingEquipment");
		set.setValue("DocumentCode", ui.lineEdit_10->text());
		set.setValue("DetectionLocation", ui.lineEdit_26->text());
		set.setValue("PragraphOne", strOne);
		set.setValue("PragraphTwo", strTwo);
		set.setValue("PragraphThree", strThree);
		set.setValue("PragraphFour", strFour);
		set.endGroup();
	}
	if (nomStaes)
	{
		QString CBtext = ui.GB_comboBox->currentText();
		CBtext.remove(QChar('/'), Qt::CaseInsensitive);
		CBtext.remove(QChar('.'), Qt::CaseInsensitive);
		QString fliePath = QString("./Config/LowVoltageCable/%1.ini").arg(CBtext);
		QSettings GBT(fliePath, QSettings::IniFormat);
		GBT.setIniCodec("UTF8");
		QStringList GBTcomboxList = GBT.childGroups();
		NomMinConfig = new NomMinConfigPramer;

		for (int i = 0; i < GBTcomboxList.size(); i++)
		{
			if ("OuterSheathThickness" == GBTcomboxList[i])
			{
				GBT.beginGroup("OuterSheathThickness");
				GBT.setValue("NomOST", ui.lineEdit_13->text().toDouble());
				GBT.setValue("MinOST", ui.lineEdit_17->text().toDouble());
				GBT.endGroup();

			}
			if ("ArmoredMetalStrip" == GBTcomboxList[i])
			{
				GBT.beginGroup("ArmoredMetalStrip");
				GBT.setValue("MinAMSThickness", ui.lineEdit_18->text().toDouble());
				GBT.setValue("MaxAMSWrapGap", ui.lineEdit_14->text().toDouble());
				GBT.setValue("MinAMSDiameter", ui.lineEdit_19->text().toDouble());
				GBT.endGroup();

			}
			if ("ArmoredWire" == GBTcomboxList[i])
			{
				GBT.beginGroup("ArmoredWire");
				GBT.setValue("MinAWThickness", ui.lineEdit_18->text().toDouble());
				GBT.setValue("MaxAWrapGap", ui.lineEdit_14->text().toDouble());
				GBT.setValue("MinAWDiameter", ui.lineEdit_19->text().toDouble());
				GBT.endGroup();

			}
			if ("LinerOuterDiameter" == GBTcomboxList[i])
			{
				GBT.beginGroup("LinerOuterDiameter");
				GBT.setValue("NomLOD", ui.lineEdit_20->text().toDouble());
				GBT.setValue("MinLOD", ui.lineEdit_21->text().toDouble());
				GBT.endGroup();

			}
			if ("InsulationThickness" == GBTcomboxList[i])
			{
				GBT.beginGroup("InsulationThickness");
				GBT.setValue("NomIT", ui.lineEdit_22->text().toDouble());
				GBT.setValue("MinIT", ui.lineEdit_23->text().toDouble());
				GBT.endGroup();
			}
			if ("SingleWiresNumber" == GBTcomboxList[i])
			{
				GBT.beginGroup("SingleWiresNumber");
				GBT.setValue("NumberSWN", ui.lineEdit_16->text().toDouble());
				GBT.endGroup();
			}
		}
	}	
}




void OutputExcel::OutputPDF()
{
	QString comBoxName = ui.checkBox->text();
	QString SampleNumber = ui.lineEdit_2->text();
	QString Manufactur = ui.lineEdit_6->text();
	QString ModelSpecifications = ui.lineEdit_7->text();
	QString RoomTemperature = ui.lineEdit_8->text();
	QString RelativeHumidity = ui.lineEdit_9->text();

	QString straname = comBoxName;

	QString fileName = straname + ".pdf";        //qDebug()<<str.at(i);        
	QFile pdfFile(QString("./doc/%1").arg(fileName));        //判断文件是否存在      
// 打开要写入的pdf文件        
	pdfFile.open(QIODevice::WriteOnly);
	QPdfWriter* pPdfWriter = new QPdfWriter(&pdfFile);  // 创建pdf写入器     
	pPdfWriter->setPageSize(QPagedPaintDevice::A4);     // 设置纸张为A4           
	pPdfWriter->setResolution(300);                     // 设置纸张的分辨率为300,因此其像素为3508X2479      
	int iMargin = 60;                   // 页边距        
	pPdfWriter->setPageMargins(QMarginsF(iMargin, iMargin, iMargin, iMargin));
	QPainter* pPdfPainter = new QPainter(pPdfWriter);   // qt绘制工具            // 标题,居中        
	QTextOption option(Qt::AlignHCenter | Qt::AlignVCenter);
	option.setWrapMode(QTextOption::WordWrap);             //二维码       
	//pPdfPainter->drawPixmap(1600,70,qrimage);            //字体       
	QFont font;
	font.setFamily("宋体");            //标题，字号       
	int fontSize = 18;
	font.setPointSize(fontSize);
	pPdfPainter->setFont(font);                    // 为绘制工具设置字体      
	pPdfPainter->drawText(QRect(0, 0, 1980, 100), Qt::AlignHCenter | Qt::AlignBottom,
		QStringLiteral("标志及结构检查(%1)检测原始记录").arg(comBoxName));            //option.setWrapMode(QTextOption::WordWrap);            //画表格       
	pPdfPainter->setFont(QFont("宋体", 10));
	pPdfPainter->drawText(0, 180, QStringLiteral("文件代号：%1").arg(configBlack->DocumentCode));
	//pPdfPainter->drawText(1000, 180, QStringLiteral("%1").arg(straname));
	pPdfPainter->drawText(1750, 180, QStringLiteral("第1页 共2页"));
	//pPdfPainter->drawText(0,250, QStringLiteral("班级："));
	QDate date = QDate::currentDate();
	QString create_time = date.toString(QStringLiteral("yyyy年MM月dd日"));
	//pPdfPainter->drawText(500,250, QStringLiteral("体检日期：%1").arg(create_time));
	//pPdfPainter->drawText(20,320, QStringLiteral("编号：%1").arg(str.at(i)));             // 标题下边留白           
	int iTop = 200;            // 左侧缩进       
	int iLeft = 0;            // 设置画笔颜色、宽度           
							   //pPdfPainter.setPen(QPen(QColor(0, 160, 230), 2));         
	pPdfPainter->setPen(2);            // 设置画刷颜色            //pPdfPainter->setBrush(QColor(255, 160, 90));      
	//pPdfPainter->drawRect(iLeft, iTop, 1980, 3750);//表格矩形方框    
	pPdfPainter->drawLine(iLeft, iTop, iLeft + 1980, iTop);
	pPdfPainter->drawLine(iLeft, iTop + 100, iLeft + 1980, iTop + 100);
	pPdfPainter->drawLine(iLeft, iTop + 200, iLeft + 1980, iTop + 200);
	pPdfPainter->drawLine(iLeft, iTop + 300, iLeft + 1980, iTop + 300);
	//pPdfPainter->drawLine(iLeft,iTop+1000,iLeft+1980,iTop+1000);         
	pPdfPainter->drawLine(990, iTop, 990, iTop + 300);
	pPdfPainter->drawLine(990 + 400, iTop + 200, 990 + 400, iTop + 300);
	pPdfPainter->drawLine(990 + 800, iTop + 200, 990 + 800, iTop + 300);

	pPdfPainter->setFont(QFont("宋体", 12));
	pPdfPainter->drawText(iLeft + 50, iTop + 65, QStringLiteral("样品编号：%1").arg(SampleNumber));
	pPdfPainter->drawText(990 + 50, iTop + 65, QStringLiteral("制 造 厂：%1").arg(Manufactur));
	pPdfPainter->drawText(iLeft + 50, iTop + 165, QStringLiteral("型号规格：%1").arg(ModelSpecifications));
	pPdfPainter->drawText(990 + 50, iTop + 165, QStringLiteral("检测地点：%1").arg(configBlack->DetectionLocation));
	pPdfPainter->drawText(iLeft + 50, iTop + 265, QStringLiteral("检测时间：%1").arg(create_time));
	pPdfPainter->drawText(990 + 50, iTop + 265, QStringLiteral("室温：%1℃").arg(RoomTemperature));
	pPdfPainter->drawText(990 + 50 + 400, iTop + 265, QStringLiteral("相对湿度：%1%").arg(RelativeHumidity));

	pPdfPainter->setFont(QFont("宋体", 12, QFont::Bold));
	pPdfPainter->drawText(iLeft, iTop + 380, QStringLiteral("一、检测仪器"));
	pPdfPainter->setFont(QFont("宋体", 12));
	pPdfPainter->drawText(iLeft + 100, iTop + 450, QStringLiteral("%1").arg(configBlack->PragraphOne));
	//pPdfPainter->drawText(iLeft + 100, iTop + 450, QStringLiteral("%1").arg(QStringLiteral("560Ⅱ立式投影仪")));
	//pPdfPainter->drawText(iLeft + 650, iTop + 450, QStringLiteral("编号：%1").arg(QStringLiteral("05 □")));
	//pPdfPainter->drawText(iLeft + 950, iTop + 450, QStringLiteral("名称：%1").arg(QStringLiteral("电缆自动投影仪")));
	//pPdfPainter->drawText(iLeft + 1500, iTop + 450, QStringLiteral("编号：%1").arg(QStringLiteral("0C08251130 □")));
	pPdfPainter->drawText(iLeft + 100, iTop + 530, QStringLiteral("%1").arg(configBlack->PragraphTwo));
	//pPdfPainter->drawText(iLeft + 100, iTop + 530, QStringLiteral("名称：%1").arg(QStringLiteral("CPJ-3015数字式测量投影仪")));
	//pPdfPainter->drawText(iLeft + 900, iTop + 530, QStringLiteral("编号：%1").arg(QStringLiteral("JGG10017 □")));
	pPdfPainter->drawText(iLeft + 100, iTop + 610, QStringLiteral("%1").arg(configBlack->PragraphThree));
	//pPdfPainter->drawText(iLeft + 100, iTop + 610, QStringLiteral("名称：%1").arg(QStringLiteral("外径千分尺")));
	//pPdfPainter->drawText(iLeft + 600, iTop + 610, QStringLiteral("编号：%1").arg(QStringLiteral("2435")));
	//pPdfPainter->drawText(iLeft + 900, iTop + 610, QStringLiteral("名称：%1").arg(QStringLiteral("游标卡尺")));
	//pPdfPainter->drawText(iLeft + 1400, iTop + 610, QStringLiteral("编号：%1").arg(QStringLiteral("050712319")));
	pPdfPainter->drawText(iLeft + 100, iTop + 690, QStringLiteral("%1").arg(configBlack->PragraphFour));
	//pPdfPainter->drawText(iLeft + 100, iTop + 690, QStringLiteral("名称：%1").arg(QStringLiteral("钢卷尺")));
	//pPdfPainter->drawText(iLeft + 600, iTop + 690, QStringLiteral("编号：%1").arg(QStringLiteral("3.5-1")));
	//pPdfPainter->drawText(iLeft + 900, iTop + 690, QStringLiteral("名称：%1").arg(QStringLiteral("外沟槽数显卡尺")));
	//pPdfPainter->drawText(iLeft + 1500, iTop + 690, QStringLiteral("编号：%1").arg(QStringLiteral("0704766")));

	pPdfPainter->setFont(QFont("宋体", 12, QFont::Bold));
	pPdfPainter->drawText(iLeft, iTop + 800, QStringLiteral("二、检测依据"));
	pPdfPainter->setFont(QFont("宋体", 10));
	pPdfPainter->drawText(iLeft + 100, iTop + 880, QStringLiteral("%1").arg(NomMinConfig->GBTName));
	//pPdfPainter->drawText(iLeft + 100, iTop + 880, QStringLiteral("GB/T12706.1—2008、国家电网公司总部配网标准化物资固化技术规范书 低压电力电缆"));

	pPdfPainter->setFont(QFont("宋体", 12, QFont::Bold));
	pPdfPainter->drawText(iLeft, iTop + 980, QStringLiteral("三、检测前对检测仪器设备及被测样品的检查"));
	pPdfPainter->setFont(QFont("宋体", 12));
	pPdfPainter->drawText(iLeft + 100, iTop + 1060, QStringLiteral("1. 游标卡尺的零位正确 □"));
	pPdfPainter->drawText(iLeft + 100, iTop + 1140, QStringLiteral("2. 数显卡尺的零位正确 □"));
	pPdfPainter->drawText(iLeft + 100, iTop + 1220, QStringLiteral("3. 被测样品外观无损 □"));

	pPdfPainter->setFont(QFont("宋体", 12, QFont::Bold));
	pPdfPainter->drawText(iLeft, iTop + 1320, QStringLiteral("四、检测数据及结果"));
	pPdfPainter->setFont(QFont("宋体", 12));

	pPdfPainter->drawLine(iLeft, iTop + 1350, iLeft + 1980, iTop + 1350);
	pPdfPainter->drawLine(iLeft + 300, iTop + 1350, iLeft + 300, iTop + 2870);
	pPdfPainter->drawLine(iLeft + 450, iTop + 1350, iLeft + 450, iTop + 2870);
	pPdfPainter->drawLine(iLeft + 800, iTop + 1350, iLeft + 800, iTop + 2870);
	pPdfPainter->drawLine(iLeft, iTop + 1430, iLeft + 1980, iTop + 1430);
	pPdfPainter->drawLine(iLeft, iTop + 1630, iLeft + 1980, iTop + 1630);
	pPdfPainter->drawLine(iLeft, iTop + 1830, iLeft + 1980, iTop + 1830);
	pPdfPainter->drawLine(iLeft, iTop + 1910, iLeft + 1980, iTop + 1910);
	pPdfPainter->drawLine(iLeft, iTop + 1990, iLeft + 1980, iTop + 1990);
	pPdfPainter->drawLine(iLeft, iTop + 2150, iLeft + 1980, iTop + 2150);
	pPdfPainter->drawLine(iLeft + 800, iTop + 2070, iLeft + 1980, iTop + 2070);
	pPdfPainter->drawLine(iLeft + 800, iTop + 2070, iLeft + 1980, iTop + 2070);
	pPdfPainter->drawLine(iLeft + 1390, iTop + 2070, iLeft + 1390, iTop + 2150);
	pPdfPainter->drawLine(iLeft, iTop + 2470, iLeft + 1980, iTop + 2470);
	pPdfPainter->drawLine(iLeft, iTop + 2870, iLeft + 1980, iTop + 2870);
	pPdfPainter->drawLine(iLeft + 800, iTop + 2710, iLeft + 1980, iTop + 2710);
	pPdfPainter->drawLine(iLeft + 800, iTop + 2790, iLeft + 1980, iTop + 2790);
	pPdfPainter->drawLine(iLeft + 1390, iTop + 2790, iLeft + 1390, iTop + 2870);
	//pPdfPainter->drawLine(iLeft, iTop + 2950, iLeft + 1980, iTop + 2950);

	pPdfPainter->setFont(QFont("宋体", 10));
	pPdfPainter->drawText(QRect(iLeft, iTop + 1350, 300, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("项  目"));
	pPdfPainter->drawText(QRect(iLeft + 300, iTop + 1350, 150, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("单 位"));
	pPdfPainter->drawText(QRect(iLeft + 450, iTop + 1350, 350, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("标 准 要 求"));
	pPdfPainter->drawText(QRect(iLeft + 800, iTop + 1350, 1180, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("实     测     值"));

	pPdfPainter->drawText(QRect(iLeft, iTop + 1430, 300, 200), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("标志内容"));
	pPdfPainter->drawText(QRect(iLeft + 300, iTop + 1430, 150, 200), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("—"));
	pPdfPainter->drawText(QRect(iLeft + 450, iTop + 1430, 350, 200), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("厂名、型号、规格"));
	QTextOption FactoryNameoption(Qt::AlignLeft | Qt::AlignVCenter);
	FactoryNameoption.setWrapMode(QTextOption::WordWrap);
	pPdfPainter->drawText(QRect(iLeft + 830, iTop + 1430, 1150, 200),
		QStringLiteral("%1").arg(QStringLiteral("FactoryName、传入FactoryName的参数值可以实现本部分的更改，字体可以自动换行")), FactoryNameoption);
	pPdfPainter->drawText(QRect(iLeft, iTop + 1630, 300, 200), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("标志清晰度"));
	pPdfPainter->drawText(QRect(iLeft + 300, iTop + 1630, 150, 200), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("—"));
	QTextOption LogoSharpnessoption(Qt::AlignLeft | Qt::AlignVCenter);
	LogoSharpnessoption.setWrapMode(QTextOption::WordWrap);
	pPdfPainter->drawText(QRect(iLeft + 480, iTop + 1630, 320, 200),
		QStringLiteral("字迹应清晰，容易辨认，耐擦"), LogoSharpnessoption);
	pPdfPainter->drawText(QRect(iLeft + 830, iTop + 1630, 1150, 200),
		QStringLiteral("%1").arg(QStringLiteral("LogoSharpness、传入LogoSharpness的参数值可以实现本部分的更改，字体可以自动换行")), LogoSharpnessoption);
	pPdfPainter->drawText(QRect(iLeft, iTop + 1830, 300, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("标志间距"));
	pPdfPainter->drawText(QRect(iLeft + 300, iTop + 1830, 150, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("mm"));
	pPdfPainter->drawText(QRect(iLeft + 450, iTop + 1830, 350, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("≤500"));
	QTextOption LogoSpacingoption(Qt::AlignLeft | Qt::AlignVCenter);
	LogoSpacingoption.setWrapMode(QTextOption::WordWrap);
	pPdfPainter->drawText(QRect(iLeft + 830, iTop + 1830, 1150, 80),
		QStringLiteral("%1").arg(QStringLiteral("LogoSpacing、传入LogoSpacing的参数值")), LogoSpacingoption);
	pPdfPainter->drawText(QRect(iLeft, iTop + 1910, 300, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("电缆外径"));
	pPdfPainter->drawText(QRect(iLeft + 300, iTop + 1910, 150, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("mm"));
	pPdfPainter->drawText(QRect(iLeft + 450, iTop + 1910, 350, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("—"));
	QTextOption CableOuteroption(Qt::AlignLeft | Qt::AlignVCenter);
	CableOuteroption.setWrapMode(QTextOption::WordWrap);
	pPdfPainter->drawText(QRect(iLeft + 830, iTop + 1910, 1150, 80),
		QStringLiteral("%1").arg(QStringLiteral("CableOuter、传入CableOuter的参数值")), CableOuteroption);
	pPdfPainter->drawText(QRect(iLeft, iTop + 1990, 300, 160), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("外护层厚度"));
	pPdfPainter->drawText(QRect(iLeft + 300, iTop + 1990, 150, 160), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("mm"));
	pPdfPainter->drawText(QRect(iLeft + 480, iTop + 1990, 320, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("标称：%1").arg(NomMinConfig->NomOST));
	pPdfPainter->drawText(QRect(iLeft + 480, iTop + 2070, 320, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("最小：%1").arg(NomMinConfig->MinOST));
	pPdfPainter->drawText(QRect(iLeft + 830, iTop + 1990, 188, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("1.%1").arg("Outer1"));
	pPdfPainter->drawText(QRect(iLeft + 1018, iTop + 1990, 188, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("2.%1").arg("Outer2"));
	pPdfPainter->drawText(QRect(iLeft + 1206, iTop + 1990, 188, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("3.%1").arg("Outer3"));
	pPdfPainter->drawText(QRect(iLeft + 1394, iTop + 1990, 188, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("4.%1").arg("Outer4"));
	pPdfPainter->drawText(QRect(iLeft + 1582, iTop + 1990, 188, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("5.%1").arg("Outer5"));
	pPdfPainter->drawText(QRect(iLeft + 1770, iTop + 1990, 188, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("6.%1").arg("Outer6"));
	pPdfPainter->drawText(QRect(iLeft + 830, iTop + 2070, 560, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("平均厚度：%1").arg("AvgThickness"));
	pPdfPainter->drawText(QRect(iLeft + 1420, iTop + 2070, 560, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("最小厚度：%1").arg("MinThickness"));
	pPdfPainter->drawText(QRect(iLeft, iTop + 2150, 300, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("铠装金属带□"));
	pPdfPainter->drawText(QRect(iLeft, iTop + 2390, 300, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("铠装金属丝□"));
	pPdfPainter->drawText(QRect(iLeft + 300, iTop + 2150, 150, 320), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("mm"));
	//pPdfPainter->drawText(QRect(iLeft + 460, iTop + 2150, 340, 80), Qt::AlignLeft | Qt::AlignVCenter,
	//	QStringLiteral("最小厚度：%1").arg("kaiMinThick"));
	pPdfPainter->drawText(QRect(iLeft + 460, iTop + 2150, 340, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("最小厚度：%1").arg(NomMinConfig->MinAMSThickness));
	pPdfPainter->drawText(QRect(iLeft + 830, iTop + 2150, 560, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("铠装外径：%1").arg("kaiOuterDia"));
	pPdfPainter->drawText(QRect(iLeft + 1420, iTop + 2150, 560, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("结构：%1").arg("KaiStructure"));
	QTextOption MaxWrapoption(Qt::AlignLeft | Qt::AlignVCenter);
	MaxWrapoption.setWrapMode(QTextOption::WordWrap);
	//pPdfPainter->drawText(QRect(iLeft + 460, iTop + 2230, 340, 160),
	//	QStringLiteral("最大绕包间隙：%1").arg("kaiMinThi123444"), MaxWrapoption);
	pPdfPainter->drawText(QRect(iLeft + 460, iTop + 2230, 340, 160),
		QStringLiteral("最大绕包间隙：%1").arg(NomMinConfig->MaxAMSWrapGap), MaxWrapoption);
	pPdfPainter->drawText(QRect(iLeft + 830, iTop + 2230, 1150, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("铠装最大绕包间隙：%1").arg("kaiOuterDia"));
	pPdfPainter->drawText(QRect(iLeft + 830, iTop + 2310, 1150, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("铠装金属带最薄点厚度：%1").arg("minThickness"));
	//pPdfPainter->drawText(QRect(iLeft + 460, iTop + 2390, 340, 80), Qt::AlignLeft | Qt::AlignVCenter,
	//	QStringLiteral("最小直径：%1").arg("minDia"));
	pPdfPainter->drawText(QRect(iLeft + 460, iTop + 2390, 340, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("最小直径：%1").arg(NomMinConfig->MinAMSDiameter));
	pPdfPainter->drawText(QRect(iLeft + 830, iTop + 2390, 1150, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("铠装金属丝最小直径：%1").arg("minSiDia"));
	pPdfPainter->drawText(QRect(iLeft, iTop + 2470, 300, 133), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("内衬层外径"));
	pPdfPainter->drawText(QRect(iLeft, iTop + 2603, 300, 133), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("结     构"));
	pPdfPainter->drawText(QRect(iLeft, iTop + 2736, 300, 133), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("厚     度"));
	pPdfPainter->drawText(QRect(iLeft + 300, iTop + 2470, 150, 400), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("mm"));

	//pPdfPainter->drawText(QRect(iLeft + 480, iTop + 2710, 320, 80), Qt::AlignLeft | Qt::AlignVCenter,
	//	QStringLiteral("标称：%1").arg("NomOuter"));
	//pPdfPainter->drawText(QRect(iLeft + 480, iTop + 2790, 320, 80), Qt::AlignLeft | Qt::AlignVCenter,
	//	QStringLiteral("最小：%1").arg("MinOuter"));
	pPdfPainter->drawText(QRect(iLeft + 480, iTop + 2710, 320, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("标称：%1").arg(NomMinConfig->NomLOD));
	pPdfPainter->drawText(QRect(iLeft + 480, iTop + 2790, 320, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("最小：%1").arg(NomMinConfig->MinLOD));
	pPdfPainter->drawText(QRect(iLeft + 830, iTop + 2710, 188, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("1.%1").arg("Outer1"));
	pPdfPainter->drawText(QRect(iLeft + 1018, iTop + 2710, 188, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("2.%1").arg("Outer2"));
	pPdfPainter->drawText(QRect(iLeft + 1206, iTop + 2710, 188, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("3.%1").arg("Outer3"));
	pPdfPainter->drawText(QRect(iLeft + 1394, iTop + 2710, 188, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("4.%1").arg("Outer4"));
	pPdfPainter->drawText(QRect(iLeft + 1582, iTop + 2710, 188, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("5.%1").arg("Outer5"));
	pPdfPainter->drawText(QRect(iLeft + 1770, iTop + 2710, 188, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("6.%1").arg("Outer6"));
	pPdfPainter->drawText(QRect(iLeft + 830, iTop + 2790, 560, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("平均厚度：%1").arg("AvgThickness"));
	pPdfPainter->drawText(QRect(iLeft + 1420, iTop + 2790, 560, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("最小厚度：%1").arg("MinThickness"));
	pPdfPainter->drawText(QRect(iLeft + 480, iTop + 2470, 350, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("挤包□"));
	pPdfPainter->drawText(QRect(iLeft + 480, iTop + 2550, 350, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("绕包□"));
	pPdfPainter->drawText(QRect(iLeft + 480, iTop + 2630, 350, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("包带加挤包□"));
	pPdfPainter->drawText(iLeft + 1750, iTop + 2920, QStringLiteral("（转背面）"));


	iLeft = 0;
	iTop = 50;
	pPdfWriter->newPage();
	pPdfPainter->drawText(0, 0, QStringLiteral("文件代号：%1").arg(QStringLiteral("CEPRI-D-EETC08-JS-701/1 （接正面）")));
	pPdfPainter->drawLine(iLeft, iTop, iLeft + 1980, iTop);
	pPdfPainter->drawLine(iLeft, iTop + 80, iLeft + 1980, iTop + 80);
	pPdfPainter->drawLine(iLeft, iTop + 160, iLeft + 1980, iTop + 160);
	pPdfPainter->drawLine(iLeft, iTop + 240, iLeft + 1980, iTop + 240);
	pPdfPainter->drawLine(iLeft, iTop + 320, iLeft + 1980, iTop + 320);
	pPdfPainter->drawLine(iLeft + 300, iTop, iLeft + 300, iTop + 1440);
	pPdfPainter->drawLine(iLeft + 450, iTop, iLeft + 450, iTop + 1440);
	pPdfPainter->drawLine(iLeft + 800, iTop, iLeft + 800, iTop + 1440);
	pPdfPainter->drawLine(iLeft + 1193, iTop + 160, iLeft + 1193, iTop + 320);
	pPdfPainter->drawLine(iLeft + 1586, iTop + 160, iLeft + 1586, iTop + 320);
	pPdfPainter->drawLine(iLeft, iTop + 1280, iLeft + 1980, iTop + 1280);
	pPdfPainter->drawLine(iLeft, iTop + 1360, iLeft + 1980, iTop + 1360);
	pPdfPainter->drawLine(iLeft, iTop + 1440, iLeft + 1980, iTop + 1440);
	pPdfPainter->drawLine(iLeft + 800, iTop + 640, iLeft + 1980, iTop + 640);
	pPdfPainter->drawLine(iLeft + 800, iTop + 960, iLeft + 1980, iTop + 960);
	pPdfPainter->drawLine(iLeft + 800, iTop + 1280, iLeft + 1980, iTop + 1280);
	pPdfPainter->drawLine(iLeft + 900, iTop + 400, iLeft + 1980, iTop + 400);
	pPdfPainter->drawLine(iLeft + 900, iTop + 480, iLeft + 1980, iTop + 480);
	pPdfPainter->drawLine(iLeft + 900, iTop + 560, iLeft + 1980, iTop + 560);
	pPdfPainter->drawLine(iLeft + 900, iTop + 720, iLeft + 1980, iTop + 720);
	pPdfPainter->drawLine(iLeft + 900, iTop + 800, iLeft + 1980, iTop + 800);
	pPdfPainter->drawLine(iLeft + 900, iTop + 880, iLeft + 1980, iTop + 880);
	pPdfPainter->drawLine(iLeft + 900, iTop + 960, iLeft + 1980, iTop + 960);
	pPdfPainter->drawLine(iLeft + 900, iTop + 1040, iLeft + 1980, iTop + 1040);
	pPdfPainter->drawLine(iLeft + 900, iTop + 1120, iLeft + 1980, iTop + 1120);
	pPdfPainter->drawLine(iLeft + 900, iTop + 1200, iLeft + 1980, iTop + 1200);
	pPdfPainter->drawLine(iLeft + 900, iTop + 320, iLeft + 900, iTop + 1280);
	pPdfPainter->drawLine(iLeft + 1193, iTop + 1280, iLeft + 1193, iTop + 1440);
	pPdfPainter->drawLine(iLeft + 1586, iTop + 1280, iLeft + 1586, iTop + 1440);




	pPdfPainter->setFont(QFont("宋体", 10));
	pPdfPainter->drawText(QRect(iLeft, iTop, 300, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("项  目"));
	pPdfPainter->drawText(QRect(iLeft + 300, iTop, 150, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("单 位"));
	pPdfPainter->drawText(QRect(iLeft + 450, iTop, 350, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("标 准 要 求"));
	pPdfPainter->drawText(QRect(iLeft + 800, iTop, 1180, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("实     测     值"));
	pPdfPainter->drawText(QRect(iLeft, iTop + 80, 300, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("绕包带外径"));
	pPdfPainter->drawText(QRect(iLeft + 300, iTop + 80, 150, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("mm"));
	pPdfPainter->drawText(QRect(iLeft + 450, iTop + 80, 350, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("—"));
	pPdfPainter->drawText(QRect(iLeft + 800, iTop + 80, 1180, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("%1").arg(QStringLiteral("WrapTapeOuter、传入WrapTapeOuter的参数值")));
	pPdfPainter->drawText(QRect(iLeft, iTop + 160, 300, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("绝缘标志"));
	pPdfPainter->drawText(QRect(iLeft + 300, iTop + 160, 150, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("—"));
	pPdfPainter->drawText(QRect(iLeft + 450, iTop + 160, 350, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("颜色/数字"));
	pPdfPainter->drawText(QRect(iLeft + 800, iTop + 160, 393, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("%1相").arg("A"));
	pPdfPainter->drawText(QRect(iLeft + 1193, iTop + 160, 393, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("%1相").arg("B"));
	pPdfPainter->drawText(QRect(iLeft + 1586, iTop + 160, 393, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("%1相").arg("C"));
	pPdfPainter->drawText(QRect(iLeft, iTop + 240, 300, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("绝缘外径"));
	pPdfPainter->drawText(QRect(iLeft + 300, iTop + 240, 150, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("mm"));
	pPdfPainter->drawText(QRect(iLeft + 450, iTop + 240, 350, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("—"));
	pPdfPainter->drawText(QRect(iLeft + 800, iTop + 240, 393, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("%1").arg("AR"));
	pPdfPainter->drawText(QRect(iLeft + 1193, iTop + 240, 393, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("%1").arg("BR"));
	pPdfPainter->drawText(QRect(iLeft + 1586, iTop + 240, 393, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("%1").arg("CR"));
	pPdfPainter->drawText(QRect(iLeft, iTop + 320, 300, 960), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("绝缘厚度"));
	pPdfPainter->drawText(QRect(iLeft + 300, iTop + 320, 150, 960), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("mm"));
	pPdfPainter->drawText(QRect(iLeft + 480, iTop + 720, 320, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("标称：%1").arg(NomMinConfig->NomIT));
	pPdfPainter->drawText(QRect(iLeft + 480, iTop + 800, 320, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("最小：%1").arg(NomMinConfig->MinIT));
	QTextOption Phaseoption(Qt::AlignHCenter | Qt::AlignVCenter);
	Phaseoption.setWrapMode(QTextOption::WordWrap);
	pPdfPainter->drawText(QRect(iLeft + 800, iTop + 320, 100, 320), QStringLiteral("%1相").arg("A001"), Phaseoption);
	pPdfPainter->drawText(QRect(iLeft + 800, iTop + 640, 100, 320), QStringLiteral("%1相").arg("B001"), Phaseoption);
	pPdfPainter->drawText(QRect(iLeft + 800, iTop + 960, 100, 320), QStringLiteral("%1相").arg("C001"), Phaseoption);
	pPdfPainter->drawText(QRect(iLeft + 930, iTop + 320, 510, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("1.%1").arg("AR111"));
	pPdfPainter->drawText(QRect(iLeft + 1470, iTop + 320, 510, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("2.%1").arg("AR222"));
	pPdfPainter->drawText(QRect(iLeft + 930, iTop + 400, 510, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("3.%1").arg("AR333"));
	pPdfPainter->drawText(QRect(iLeft + 1470, iTop + 400, 510, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("4.%1").arg("AR444"));
	pPdfPainter->drawText(QRect(iLeft + 930, iTop + 480, 510, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("5.%1").arg("AR555"));
	pPdfPainter->drawText(QRect(iLeft + 1470, iTop + 480, 510, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("6.%1").arg("AR666"));
	pPdfPainter->drawText(QRect(iLeft + 930, iTop + 560, 510, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("平均厚度：%1").arg("AvTH555"));
	pPdfPainter->drawText(QRect(iLeft + 1470, iTop + 560, 510, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("最小厚度：%1").arg("MinTH666"));
	pPdfPainter->drawText(QRect(iLeft + 930, iTop + 640, 510, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("1.%1").arg("AR111"));
	pPdfPainter->drawText(QRect(iLeft + 1470, iTop + 640, 510, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("2.%1").arg("AR222"));
	pPdfPainter->drawText(QRect(iLeft + 930, iTop + 720, 510, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("3.%1").arg("AR333"));
	pPdfPainter->drawText(QRect(iLeft + 1470, iTop + 720, 510, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("4.%1").arg("AR444"));
	pPdfPainter->drawText(QRect(iLeft + 930, iTop + 800, 510, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("5.%1").arg("AR555"));
	pPdfPainter->drawText(QRect(iLeft + 1470, iTop + 800, 510, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("6.%1").arg("AR666"));
	pPdfPainter->drawText(QRect(iLeft + 930, iTop + 880, 510, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("平均厚度：%1").arg("AvTH555"));
	pPdfPainter->drawText(QRect(iLeft + 1470, iTop + 880, 510, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("最小厚度：%1").arg("MinTH666"));
	pPdfPainter->drawText(QRect(iLeft + 930, iTop + 960, 510, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("1.%1").arg("AR111"));
	pPdfPainter->drawText(QRect(iLeft + 1470, iTop + 960, 510, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("2.%1").arg("AR222"));
	pPdfPainter->drawText(QRect(iLeft + 930, iTop + 1040, 510, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("3.%1").arg("AR333"));
	pPdfPainter->drawText(QRect(iLeft + 1470, iTop + 1040, 510, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("4.%1").arg("AR444"));
	pPdfPainter->drawText(QRect(iLeft + 930, iTop + 1120, 510, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("5.%1").arg("AR555"));
	pPdfPainter->drawText(QRect(iLeft + 1470, iTop + 1120, 510, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("6.%1").arg("AR666"));
	pPdfPainter->drawText(QRect(iLeft + 930, iTop + 1200, 510, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("平均厚度：%1").arg("AvTH555"));
	pPdfPainter->drawText(QRect(iLeft + 1470, iTop + 1200, 510, 80), Qt::AlignLeft | Qt::AlignVCenter,
		QStringLiteral("最小厚度：%1").arg("MinTH666"));

	pPdfPainter->drawText(QRect(iLeft, iTop + 1280, 300, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("导体外径"));
	pPdfPainter->drawText(QRect(iLeft + 300, iTop + 1280, 150, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("mm"));
	pPdfPainter->drawText(QRect(iLeft + 450, iTop + 1280, 350, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("—"));
	pPdfPainter->drawText(QRect(iLeft + 800, iTop + 1280, 393, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("%1").arg("ConductorOuter1"));
	pPdfPainter->drawText(QRect(iLeft + 1193, iTop + 1280, 393, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("%1").arg("ConductorOuter2"));
	pPdfPainter->drawText(QRect(iLeft + 1586, iTop + 1280, 393, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("%1").arg("ConductorOuter3"));
	pPdfPainter->drawText(QRect(iLeft, iTop + 1360, 300, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("单线根数"));
	pPdfPainter->drawText(QRect(iLeft + 300, iTop + 1360, 150, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("根"));
	pPdfPainter->drawText(QRect(iLeft + 450, iTop + 1360, 350, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("%1").arg(NomMinConfig->NumberSWN));
	pPdfPainter->drawText(QRect(iLeft + 800, iTop + 1360, 393, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("%1").arg("TheNumber1"));
	pPdfPainter->drawText(QRect(iLeft + 1193, iTop + 1360, 393, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("%1").arg("TheNumber2"));
	pPdfPainter->drawText(QRect(iLeft + 1586, iTop + 1360, 393, 80), Qt::AlignHCenter | Qt::AlignVCenter,
		QStringLiteral("%1").arg("TheNumber3"));

	pPdfPainter->setFont(QFont("宋体", 12, QFont::Bold));
	pPdfPainter->drawText(iLeft, iTop + 1550, QStringLiteral("五、检测后对检测仪器设备及被测样品的检查"));
	pPdfPainter->setFont(QFont("宋体", 12));
	pPdfPainter->drawText(iLeft + 100, iTop + 1630, QStringLiteral("1. 游标卡尺的零位正确 □"));
	pPdfPainter->drawText(iLeft + 100, iTop + 1680, QStringLiteral("2. 数显卡尺的零位正确 □"));

	pPdfPainter->setFont(QFont("宋体", 12, QFont::Bold));
	pPdfPainter->drawText(iLeft, iTop + 3000, QStringLiteral("检测："));
	pPdfPainter->drawText(iLeft + 800, iTop + 3000, QStringLiteral("记录："));
	pPdfPainter->drawText(iLeft + 1600, iTop + 3000, QStringLiteral("校核："));


	delete pPdfPainter;
	delete pPdfWriter;
	pdfFile.close();            //将程序当前路径设置为原来的路径    
}



