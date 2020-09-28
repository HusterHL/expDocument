#pragma once

#include <QtWidgets/QMainWindow>
#include "ui_OutputExcel.h"
#include <QObject>
#include<ActiveQt\QAxWidget> 
#include <ActiveQt\QAxObject>
#include <ActiveQt\QAxBase>
#include <QFile>
#include <QTextStream>
#include <QDir>
#include <QDate>
#include <QSettings>
#include <QTextCodec>
#include <QMessageBox>
#include <QPdfWriter>
#include <QPainter>
#include <QFileDialog>


/*

* ΢���word�Ĳ鿴������վ

* https://docs.microsoft.com/zh-cn/dotnet/api/microsoft.office.interop.word?view=word-pia

*

* �ο�https://blog.csdn.net/u010304326/article/details/82292195#comments

* �ο�https://blog.csdn.net/qq_35192280/article/details/83021975

* https://blog.csdn.net/zy47675676/article/details/86251991 ���ֱ���С�ˮƽ����

*/



enum TITLE_NUMBER
{
	TITLE_ONE = 0,
	TITLE_TWO,
	TITLE_THREE,
	NORMAL
};

//MOVEEND_INDEX������https://docs.microsoft.com/zh-cn/dotnet/api/microsoft.office.interop.word.wdunits?view=word-pia
enum MOVEEND_INDEX
{
	wdParagraph = 4, //���䡣
	wdStory = 6, 	//���֡�
	wdRow = 10, //��
	wdParagraphFormatting = 14, //�����ʽ��
	wdTable = 15 //���
};

//�ı����뷽ʽ
enum WdParagraphAlignment
{
	AlignParLeft = 0, //�����
	AlignParCenter = 1, //���ж��롣	AlignParRight = 2, //�Ҷ��롣
	AlignParJustify = 3, //��ȫ���˶��롣
};


struct ConfigPramerBlack
{
	QString DocumentCode;//�ļ���
	QString DetectionLocation;//���ص�
	QString PragraphOne;//��������ĵ�һ�仰
	QString PragraphTwo;//��������ĵڶ��仰
	QString PragraphThree;//��������ĵ����仰
	QString PragraphFour;//��������ĵ��ľ仰
};

struct NomMinConfigPramer
{
	QString	GBTName;
	double NomOST;//�⻤���ƺ��
	double MinOST;//�⻤����С���
	double MinAMSThickness;//��װ��������С���
	double MaxAMSWrapGap;//��װ����������ư���϶
	double MinAMSDiameter;//��װ��������Сֱ��
	double MinAWThickness;//��װ����˿��С���
	double MaxAWrapGap;//��װ����˿����ư���϶
	double MinAWDiameter;//��װ����˿��Сֱ��
	double NomLOD;//�ڳĲ����⾶
	double MinLOD;//�ڳĲ���С�⾶
	double NomIT;//��Ե��ƺ��
	double MinIT;//��Ե��С���
	double NumberSWN;//���߸���
};

//struct boolStates
//{
//	bool OuterSheathThicknessStates;
//	bool ArmoredMetalStripStates;
//	bool ArmoredWireStates;
//	bool LinerOuterDiameterStates;
//	bool InsulationThicknessStates;
//	bool SingleWiresNumberStates;
//};



class OutputExcel : public QMainWindow
{
	Q_OBJECT

public:
	OutputExcel(QWidget *parent = Q_NULLPTR);
	void CreatExcel();

private:
	Ui::OutputExcelClass ui;
public:
	void intsertTable(int row, int column);
	//���ļ� bVisable �Ƿ���ʾ����
	bool open(bool bvisable = false);
	bool open(const QString& strFile, bool bVisable = false);
	///////////////////////////////////////////////////////////////////////////
	//�ر��ļ�
	bool close();
	bool isOpen();
	//����
	void save();
	void saveAs(const QString& strSaveFile);
	//////////////////////////////////////////////////////////////////////////
	//����ı� titlestr ��ӵ��ı� number ���⻹�����ģ�Ĭ��������
	bool addText(QString titlestr, TITLE_NUMBER number = NORMAL, WdParagraphAlignment alignment = AlignParLeft);
	//Ĭ���Ǻ�ɫ,����:û��ʵ����һ��ʵ�ֲ�ͬ��ɫ
	bool addText(QString titlestr, QFont font, QColor fontcolor = Qt::black);
	//����QAxObject����������ɫ�Լ�������ʽ
	QAxObject* addText2(QString titlestr);
	//
	//////////////////////////////////////////////////////////////////////////
	//���ܣ�����س�
	bool insertEnter();
	//////////////////////////////////////////////////////////////////////////
	void moveRight();
	//����ƶ������
	bool moveToEnd();
	bool moveToEnd(MOVEEND_INDEX wd);
	//�ƶ�������������һ��
	bool moveToEnd(QAxObject *table);

	//////////////////////////////////////////////////////////////////////////
//������====================================================================

	/******************************************************************************
	* ������insertTable
	* ���ܣ��������
	* ������nStart ��ʼλ��; nEnd ����λ��; row ��; column ��
	* ����ֵ�� QAxObject*
	*****************************************************************************/
	QAxObject* insertTable(int nStart, int nEnd, int row, int column);

	/******************************************************************************
	*�������
	*QStringList headList ��ӱ�ͷ
	******************************************************************************/
	QAxObject* createTable(int row, int column);

	//�����п�
	void setColumnWidth(QAxObject *table, int column, int width);
	// Ϊ��������
	void addTableRow(QAxObject *table, int nRow, int rowCount);
	void appendTableRow(QAxObject *table, int rowCount);
	/******************************************************************************
	* ������setCellString
	* ���ܣ����ñ������
	* ������table ���; row ����; column ����; text �����ı�   row �� column��0��ʼ
	*****************************************************************************/
	void setCellString(QAxObject *table, int row, int column, const QString& text);
	// �������ݴ���  isBold�����Ƿ����
	void setCellFontBold(QAxObject *table, int row, int column, bool isBold);
	// �������ִ�С
	void setCellFontSize(QAxObject *table, int row, int column, int size);
	// �ڱ���в���ͼƬ
	void insertCellPic(QAxObject *table, int row, int column, const QString& picPath);

	/******************************************************************************
	* ������MergeCells
	* ���ܣ��ϲ���Ԫ��
	* ������table ���; nStartRow ��ʼ��Ԫ������; nStartCol ; nEndRow ; nEndCol
	*****************************************************************************/
	void MergeCells(QAxObject *table, int nStartRow, int nStartCol, int nEndRow, int nEndCol);
	//===============================================================================
		//����ͼƬ picPath ͼƬ·��
	void insertPic(QString picPath);
	void typeText(QString text);
	//���ֶ��뷽ʽ
	void setAlignment(int index);
	//������ɫ ����ֱ������QColor��Ҫ����ɫת��intֵ
	void setColor(QColor color);
	void setColor(QAxObject *obj, QColor color);
	void setBgColor(QAxObject *obj, QColor color);
	//�����ֺ�
	void setFontSize(int size);
	void readConfig();
	void clearConfig();
	void OutputPDF();
private:
	void writeFile(QString savestr, QString filename);
	QString getTitleStr(TITLE_NUMBER number); //���ر����ַ���
	void setPropraty(QAxObject *axobj, QString proname, QVariant provalue); //����ĳ�������ĳ������ֵ
	int colorToInt(QColor color); //����ɫת������������ΪQColor("blue").value()��255��������Ҫ�Ľ��
private:
	QString m_filename;
	bool m_bOpened;
	bool laodconfig;
	QAxObject *m_wordDocuments;
	QAxWidget *m_wordWidget;
	WdParagraphAlignment m_paralignment; //�ı����뷽ʽ

	ConfigPramerBlack *configBlack;
	NomMinConfigPramer *NomMinConfig;
	QWidget *widgetOne;
	QWidget *widgetTwo;
	//QPushButton *pushButtonOne;
	//QPushButton *pushButtonTwo;
	//QPushButton *pushButtonThree;
	//QPushButton *pushButtonFour;
	//QPushButton *pushButtonFive;
	//QPushButton *pushButtonSix;
	//boolStates *configStates;
private slots:
	void on_BnCreatWord_clicked();
	void on_pushButton_clicked();
	void on_BnConfigJion_clicked();
	void on_BnOverLoad_clicked();
	void LoadConfig();
	void ShowConfig();
	void hideTabWidget(); 
signals:
	void ConfigSignal(QString);

};
