#include "mainwindow.h"
#include <QApplication>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    QString strPath = QApplication::applicationDirPath();
    strPath += "/img/timg.jpg";
    a.setWindowIcon(QIcon(strPath));
    MainWindow w;
    w.show();
    return a.exec();
}
