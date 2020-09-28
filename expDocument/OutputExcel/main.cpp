#include "OutputExcel.h"
#include <QtWidgets/QApplication>

int main(int argc, char *argv[])
{
	QApplication a(argc, argv);
	OutputExcel w;
	w.show();
	QFile styleFile("style.qss");
	if (styleFile.open(QIODevice::ReadOnly))
	{
		qDebug("open success");
		QString setStyleSheet(styleFile.readAll());
		a.setStyleSheet(setStyleSheet);
		styleFile.close();
	}
	else
	{
		qDebug("Open failed");
	}
	return a.exec();
}
