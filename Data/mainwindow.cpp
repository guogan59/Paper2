#include "mainwindow.h"
#include "ui_mainwindow.h"
#include "readexcel.h"
#include <QSqlQuery>
#include <QDebug>
#include <QElapsedTimer>
#include <QDateTime>
#include <QSqlError>
# pragma execution_character_set("utf-8")

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    int                          i = 0;
    CReadWrite                   ReadWrite;
    CUpdateMessage               UpdateMessage;
    CContrastMessage             ContrastMessage;
    ui->setupUi(this);
    connect(ui->lineEdit, SIGNAL(returnPressed()), ui->pushButton_Search, SIGNAL(clicked()), Qt::UniqueConnection);

    ///////搜索界面更新///////
    QStringList header;
    ReadWrite.XML_Undate_Read(UpdateMessage);                  //读取配置文件信息，需要更新哪几列数据
    for(i=0;i<UpdateMessage.m_Column.size();i++)
    {
        header<<QString::fromStdString(UpdateMessage.m_Column[i].m_ColumnName);
    }
    ui->tableWidget->setRowCount(0);                                          //设置行数
    ui->tableWidget->setColumnCount(UpdateMessage.m_Column.size());         //设置列数,不设置不显示表头
    ui->tableWidget->setHorizontalHeaderLabels(header);
    ui->tableWidget->horizontalHeader()->setStyleSheet("QHeaderView::section{background:skyblue;}"); //设置表头背景色
    ui->tableWidget->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);
    ui->tableWidget->setEditTriggers(QAbstractItemView::NoEditTriggers);      //表格只读模式
    ui->tableWidget->setSelectionBehavior ( QAbstractItemView::SelectRows);

    ///////比对界面更新///////
    QStringList header2;
    ReadWrite.XML_Contrast_Read(ContrastMessage);                  //读取配置文件信息，需要比对哪几列数据
    for(i=0;i<ContrastMessage.m_Column.size();i++)
    {
        header2<<QString::fromStdString(ContrastMessage.m_Column[i].m_ColumnName);
    }
    for(i=0;i<ContrastMessage.m_Contrast_Column.size();i++)
    {
        header2<<("数据库-"+QString::fromStdString(ContrastMessage.m_Contrast_Column[i]));
    }
    header2<<"所在Sheet";
    ui->tableWidget_2->setRowCount(0);                                            //设置行数
    ui->tableWidget_2->setColumnCount(ContrastMessage.m_Column.size()
                                      +ContrastMessage.m_Contrast_Column.size()+1);         //设置列数,不设置不显示表头
    ui->tableWidget_2->setHorizontalHeaderLabels(header2);
    ui->tableWidget_2->horizontalHeader()->setStyleSheet("QHeaderView::section{background:skyblue;}"); //设置表头背景色
    ui->tableWidget_2->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);
    ui->tableWidget_2->setEditTriggers(QAbstractItemView::NoEditTriggers);      //表格只读模式
    ui->tableWidget_2->setSelectionBehavior ( QAbstractItemView::SelectRows);
}

MainWindow::~MainWindow()
{
    delete ui;
    resize( QSize( 800, 600 ));
}

//********************************************************************
/// 读取excel更新数据库
/// @Author	Guo Shuai
/// @Date	2019-01-4
/// @Input
/// @Return
//********************************************************************
void MainWindow::on_pushButton_excelname_clicked()
{
    int                          i = 0;
    int                          j = 0;
    int                          SheetNum = 1;
    CReadWrite                   ReadWrite;
    QList< QList<QVariant> >     Datas;
    CUpdateMessage               UpdateMessage;

    //
    ///////读取excel并更新数据库///////
    //
    ///////读取Excel位置///////
    QString ExcelFile = QFileDialog::getOpenFileName(this,("打开"),"","",
          &QString::fromLocal8Bit("excel(*.xls *.xlsx)"),0);  //读取excel位置
    SheetNum = ui->spinBox->value();                          //读取Sheet序号
    if(ExcelFile.isEmpty())                                   //判断是否为空
    {
        QMessageBox::warning(NULL, "警告", "Excel路径为空!");
    }
    else
    {
        ///////新建数据库///////
        QDateTime time = QDateTime::currentDateTime();                  //获取系统现在的时间
        QString DataBaseName = QApplication::applicationDirPath();
        DataBaseName += "/DataBase/"  + time.toString("yyyyMMddhhmmss") + ".db"; //设置显示格式
        QSqlDatabase Database = QSqlDatabase::addDatabase("QSQLITE");
        Database.setDatabaseName(DataBaseName);
        if (!Database.open())
        {
           QMessageBox::warning(NULL, "警告", "更新数据库出错!");
        }
        else
        {
            ///////读取数据库以及配置信息///////
            ReadWrite.XML_Undate_Read(UpdateMessage);                  //读取配置文件信息，需要更新哪几列数据
            ReadWrite.Excel_Read(ExcelFile, SheetNum, Datas);     //读取文件
            ReadWrite.XML_Undate_Modify(DataBaseName);                   //修改配置文件中更新数据库名称
            QSqlQuery Query(DataBaseName);                        //数据库连接
            ///////数据库构建///////
            if(UpdateMessage.m_Column.size() == 0)
            {
                QMessageBox::warning(NULL, "警告", "Excel配置信息不全!");    //配置信息错误警告
            }
            else if (Datas.size() == 0)
            {
                QMessageBox::warning(NULL, "警告", "Excel文件有问题!");    //配置信息错误警告
            }
            else
            {
                QString ColumnName = QString::fromStdString(UpdateMessage.m_Column[0].m_ColumnName);          //数据库各列信息名称
                Query.exec(QString("create table SearchData (%1 varchar(20) primary key)").arg(ColumnName));    //数据库各列信息名称

                for(i=1;i<UpdateMessage.m_Column.size();i++)
                {
                    ColumnName = QString::fromStdString(UpdateMessage.m_Column[i].m_ColumnName);
                    Query.exec(QString("alter table SearchData add %1 varchar(20)").arg(ColumnName));
                }
                ///////数据录入///////
                QSqlQuery SqlQuery;
                QString ValuesNum = "?";                    //数据数量
                for(i=1;i<UpdateMessage.m_Column.size();i++)
                {
                    ValuesNum = ValuesNum + ", " +"?";
                }
                SqlQuery.prepare(QString("insert into SearchData values (%1)").arg(ValuesNum));

                ///////数据录入///////
                bool OutRange=true;
                int  Startline = ui->lineEdit_3->text().toInt();
                for(j=0;j<UpdateMessage.m_Column.size();j++)
                {
                    QVariantList values;
                    for(i = Startline -1;i<Datas.size();i++)
                    {
                        if(UpdateMessage.m_Column[j].m_ColumnNum <= Datas[i].size())
                        {
                            OutRange = true;
                            values<<Datas[i][UpdateMessage.m_Column[j].m_ColumnNum-1];
                        }
                        else
                        {
                            QMessageBox::warning(NULL, "警告", "Excel文件中找不到需更新数据!");
                            OutRange = false;
                            break;
                        }
                    }
                    if(OutRange == true && values.size() != 0)
                    {
                        SqlQuery.addBindValue(values);
                    }
                }
                if (!SqlQuery.execBatch())    //进行批处理，如果出错就输出错误
                {
                    QMessageBox::warning(NULL, "警告", "数据库更新失败!");    //配置信息错误警告
                }
                else
                {
                    QString strPath = QApplication::applicationDirPath();
                    strPath += "/img/chai.jpg";
                    QMessageBox message(QMessageBox::NoIcon, "提示", "数据库更新成功！");
                    message.setIconPixmap(QPixmap(strPath));
                    message.exec();
                }
            }
        }
    }
}

//********************************************************************
/// 数据搜索
/// @Author	Guo Shuai
/// @Date	2019-01-4
/// @Input
/// @Return
//********************************************************************
void MainWindow::on_pushButton_Search_clicked()
{
    int                          i=0;
    int                          k=0;
    int                          j=0;
    CReadWrite                   ReadWrite;
    CUpdateMessage               UpdateMessage;
    QSqlDatabase                 Database = QSqlDatabase::addDatabase("QSQLITE");

    //
    ///////从数据库中搜索信息///////
    //
    ///////数据库链接///////
    ReadWrite.XML_Undate_Read(UpdateMessage);                  //读取配置文件信息，需要更新哪几列数据
    Database.setDatabaseName(QString::fromStdString(UpdateMessage.m_DataBaseName));
    if (!Database.open())
    {
       QMessageBox::warning(NULL, "警告", "更新数据库出错!");
    }

    ///////查找数据///////
    std::string searchColumn;
    if(UpdateMessage.m_Column.size() != 0)
    {
        searchColumn = UpdateMessage.m_Column[0].m_ColumnName;
    }
    for(i=1;i<UpdateMessage.m_Column.size();i++)
    {
        searchColumn = searchColumn +","+ UpdateMessage.m_Column[i].m_ColumnName;
    }
    QString QSearchColumn = QString::fromStdString(searchColumn);      //SQL语言，显示列名称

    ///////更新界面///////
    ui->tableWidget->clear();
    QStringList header;
    for(i=0;i<UpdateMessage.m_Column.size();i++)
    {
        header<<QString::fromStdString(UpdateMessage.m_Column[i].m_ColumnName);
    }
    ui->tableWidget->setColumnCount(UpdateMessage.m_Column.size());         //设置列数,不设置不显示表头
    ui->tableWidget->setHorizontalHeaderLabels(header);
    int iLen = ui->tableWidget->rowCount();
    for(int i=0;i<iLen;i++)
    {
        ui->tableWidget->removeRow(i-j);
        j++;
    }

    ///////搜索关键词///////
    QSqlQuery query;
    QString SearchValue= "'%" + ui->lineEdit->text() + "%'";           //搜索关键词
    QList<QVariant>    OldData;                                        //用以存储重要关键字，删除重复数据
    bool               DeleteJudge;                                    //删除判断
    if(ui->lineEdit->text() != NULL)
    {
        for(i=0;i<UpdateMessage.m_Column.size();i++)                 //开始搜索
        {
            query.exec(QString("select %1 from SearchData where %2 like %3 ").arg(QSearchColumn)
                       .arg(QString::fromStdString(UpdateMessage.m_Column[i].m_ColumnName)).arg(SearchValue));
            while(query.next())
            {
                ///////重复数据判断///////
                DeleteJudge = true;
                for(j=0;j<OldData.size();j++)
                {
                    if (OldData[j] == query.value(0))
                    {
                        DeleteJudge = false;
                    }
                }
                ///////将搜索数据更新至界面上///////
                if(DeleteJudge == true)
                {
                    OldData.push_back(query.value(0));
                    int row = ui->tableWidget->rowCount();
                    ui->tableWidget->insertRow(row);
                    //qDebug()<<query.value(0).toString()<<query.value(1).toString();
                    for(j=0;j<UpdateMessage.m_Column.size();j++)
                    {
                        ui->tableWidget->setItem(k,j,new QTableWidgetItem(query.value(j).toString()));
                        if( j%2 == 0)
                        {
                            ui->tableWidget->horizontalHeader()->setSectionResizeMode(j, QHeaderView::ResizeToContents);
                            //ui->tableWidget->horizontalHeader()->setSectionResizeMode(j, QHeaderView::Stretch);
                        }
                        else
                        {
                            ui->tableWidget->horizontalHeader()->setSectionResizeMode(j, QHeaderView::Stretch);
                        }
                    }
                    k++;
                }
            }
        }
        ui->tableWidget->resizeRowsToContents();//所有行高度自适应
    }
    else
    {
        query.exec("select * from SearchData");
        while(query.next())
        {
            ///////重复数据判断///////
            DeleteJudge = true;
            for(j=0;j<OldData.size();j++)
            {
                if (OldData[j] == query.value(0))
                {
                    DeleteJudge = false;
                }
            }
            ///////将搜索数据更新至界面上///////
            if(DeleteJudge == true)
            {
                OldData.push_back(query.value(0));
                int row = ui->tableWidget->rowCount();
                ui->tableWidget->insertRow(row);
                //qDebug()<<query.value(0).toString()<<query.value(1).toString();
                for(j=0;j<UpdateMessage.m_Column.size();j++)
                {
                    ui->tableWidget->setItem(k,j,new QTableWidgetItem(query.value(j).toString()));
                    if( j%2 == 0)
                    {
                        ui->tableWidget->horizontalHeader()->setSectionResizeMode(j, QHeaderView::ResizeToContents);
                    }
                    else
                    {
                        ui->tableWidget->horizontalHeader()->setSectionResizeMode(j, QHeaderView::Stretch);
                    }
                }
                k++;
            }
        }

        ui->tableWidget->resizeRowsToContents();//所有行高度自适应
    }
}

//********************************************************************
/// Excel对比数据
/// @Author	Guo Shuai
/// @Date	2019-01-8
/// @Input
/// @Return
//********************************************************************
void MainWindow::on_pushButton_clicked()
{
    int                            i = 0;
    int                            j = 0;
    int                            k = 0;
    int                            n = 0;
    int                            g = 0;
    int                            SheetNum;
    int                            SingleSheetNum = 1;
    CReadWrite                     ReadWrite;
    QList< QList<QVariant> >       SingleSheet;
    QList<QList< QList<QVariant>>> SeveralSheets;
    CContrastMessage               ContrastMessage;

    ///////数据库链接///////
    CUpdateMessage                 UpdateMessage;
    QSqlDatabase                   Database = QSqlDatabase::addDatabase("QSQLITE");
    ReadWrite.XML_Undate_Read(UpdateMessage);                  //读取配置文件信息，需要更新哪几列数据
    Database.setDatabaseName(QString::fromStdString(UpdateMessage.m_DataBaseName));
    if (!Database.open())
    {
       QMessageBox::warning(NULL, "警告", "更新数据库出错!");
    }

    ///////更新界面///////
    ReadWrite.XML_Contrast_Read(ContrastMessage);                  //读取配置文件信息，需要比对哪几列数据
    ui->tableWidget_2->clear();                                    //将界面表格内容清空
    QStringList header;                                            //更新表头
    for(i=0;i<ContrastMessage.m_Column.size();i++)
    {
        header<<QString::fromStdString(ContrastMessage.m_Column[i].m_ColumnName);
    }
    for(i=0;i<ContrastMessage.m_Contrast_Column.size();i++)
    {
        header<<("数据库-"+QString::fromStdString(ContrastMessage.m_Contrast_Column[i]));
    }
    header<<"所在Sheet";
    ui->tableWidget_2->setColumnCount(ContrastMessage.m_Column.size() + 1
                                      +ContrastMessage.m_Contrast_Column.size());         //设置列数,不设置不显示表头
    ui->tableWidget_2->setHorizontalHeaderLabels(header);
    int iLen = ui->tableWidget_2->rowCount();                         //删除各行
    for(int i=0;i<iLen;i++)
    {
        ui->tableWidget_2->removeRow(i-j);
        j++;
    }

    ///////读取比对Excel位置///////
    QString ExcelFile = QFileDialog::getOpenFileName(this,("打开"),"",
                                                     "",&QString::fromLocal8Bit("excel(*.xls *.xlsx)"),0);  //读取excel位置
    SingleSheetNum = ui->spinBox_2->value();                          //读取Sheet序号
    int comboBoxindex = ui->comboBox->currentIndex();
    if(ExcelFile.isEmpty())                                           //判断是否为空
    {
        QMessageBox::warning(NULL, "警告", "Excel路径为空!");
    }
    else
    {
        ///////单Sheet比对模式///////
        if(comboBoxindex == 0)
        {
            ReadWrite.Excel_Read(ExcelFile, SingleSheetNum, SingleSheet);          //读取对比excel文件
            if(SingleSheet.size() == 0)
            {
                QMessageBox::warning(NULL, "警告", "Excel读取失败!");    //配置信息错误警告
            }
            else
            {
                ///////根据配置文件，从数据库提取数据///////
                std::string searchColumn;
                if(ContrastMessage.m_Contrast_Column.size() != 0)
                {
                    searchColumn = ContrastMessage.m_Contrast_Column[0];
                }
                for(i=1;i<ContrastMessage.m_Contrast_Column.size();i++)
                {
                    searchColumn = searchColumn +","+ ContrastMessage.m_Contrast_Column[i];
                }
                QString SearchColumn = QString::fromStdString(searchColumn);      //SQL语言，显示列名称
                ///////判断关键词列数///////
                int KeyColumn;
                for(i=0;i<ContrastMessage.m_Column.size();i++)
                {
                    if(ContrastMessage.m_Column[i].m_ColumnName == ContrastMessage.m_Contrast_KeyColumn)
                    {
                        KeyColumn = ContrastMessage.m_Column[i].m_ColumnNum;
                        break;
                    }
                }
                ///////根据关键词，判断每一行数据///////
                int       m =0;
                int       ContrastColumn = 0;
                bool      DifferentJudge = true;
                QSqlQuery query;
                if(SingleSheet.size() >= ui->lineEdit_2->text().toInt())    //判断行数
                {
                    for(i=ui->lineEdit_2->text().toInt()-1;i<SingleSheet.size();i++)
                    {
                        QString SearchValue= "'%" + SingleSheet[i][KeyColumn-1].toString() + "%'";
                        query.exec(QString("select %1 from SearchData where %2 like %3 ").arg(SearchColumn)
                                   .arg(QString::fromStdString(ContrastMessage.m_Contrast_KeyColumn)).arg(SearchValue));
                        while(query.next())
                        {
                            for(j=0;j<ContrastMessage.m_Contrast_Column.size();j++)
                            {
                                for(k=0;k<ContrastMessage.m_Column.size();k++)
                                {
                                    if(ContrastMessage.m_Column[k].m_ColumnName == ContrastMessage.m_Contrast_Column[j])
                                    {
                                        ContrastColumn = ContrastMessage.m_Column[k].m_ColumnNum;
                                        break;
                                    }
                                }

                                if((query.value(j).toString()) != (SingleSheet[i][ContrastColumn-1]).toString())
                                {
                                    DifferentJudge = false;
                                }
                            }
                            if(DifferentJudge == false)
                            {
                                int row = ui->tableWidget_2->rowCount();
                                ui->tableWidget_2->insertRow(row);
                                for(k=0;k<ContrastMessage.m_Column.size();k++)
                                {
                                    ui->tableWidget_2->setItem(m,k,new QTableWidgetItem(SingleSheet[i][ContrastMessage.m_Column[k].m_ColumnNum-1].toString()));
                                }
                                for(n=0;n<ContrastMessage.m_Contrast_Column.size();n++)
                                {
                                    ui->tableWidget_2->setItem(m,k,new QTableWidgetItem(query.value(n).toString()));
                                    k++;
                                }
                                ui->tableWidget_2->setItem(m,k,new QTableWidgetItem(QString::number(ui->spinBox_2->value())));
                                m++;
                            }
                            DifferentJudge = true;
                        }
                    }
                    ui->tableWidget_2->resizeRowsToContents();//所有行高度自适应
                }
           }
            QString strPath = QApplication::applicationDirPath();
            strPath += "/img/chai2.jpg";
            QMessageBox message(QMessageBox::NoIcon, "提示", "比对完成！");
            message.setIconPixmap(QPixmap(strPath));
            message.exec();
        }
        else if(comboBoxindex == 1)
        {
            qsrand(QTime(0,0,0).secsTo(QTime::currentTime()));
            ReadWrite.Excel_SeveralSheets_Read(ExcelFile, SingleSheetNum, SeveralSheets);          //读取对比excel文件
            int       m =0;
            for(g=0;g<SingleSheetNum;g++)
            {
                int color1 = qrand()%255;
                int color2 = qrand()%255;
                int color3 = qrand()%255;
                if(SeveralSheets[g].size() == 0)
                {
                    QMessageBox::warning(NULL, "警告", "Excel读取失败!");    //配置信息错误警告
                }
                else
                {
                    ///////根据配置文件，从数据库提取数据///////
                    std::string searchColumn;
                    if(ContrastMessage.m_Contrast_Column.size() != 0)
                    {
                        searchColumn = ContrastMessage.m_Contrast_Column[0];
                    }
                    for(i=1;i<ContrastMessage.m_Contrast_Column.size();i++)
                    {
                        searchColumn = searchColumn +","+ ContrastMessage.m_Contrast_Column[i];
                    }
                    QString SearchColumn = QString::fromStdString(searchColumn);      //SQL语言，显示列名称
                    ///////判断关键词列数///////
                    int KeyColumn;
                    for(i=0;i<ContrastMessage.m_Column.size();i++)
                    {
                        if(ContrastMessage.m_Column[i].m_ColumnName == ContrastMessage.m_Contrast_KeyColumn)
                        {
                            KeyColumn = ContrastMessage.m_Column[i].m_ColumnNum;
                            break;
                        }
                    }
                    ///////根据关键词，判断每一行数据///////
                    int       ContrastColumn = 0;
                    bool      DifferentJudge = true;
                    QSqlQuery query;
                    if(SeveralSheets[g].size() >= ui->lineEdit_2->text().toInt())    //判断行数
                    {
                        for(i=ui->lineEdit_2->text().toInt()-1;i<SeveralSheets[g].size();i++)
                        {
                            QString SearchValue= "'%" + SeveralSheets[g][i][KeyColumn-1].toString() + "%'";
                            query.exec(QString("select %1 from SearchData where %2 like %3 ").arg(SearchColumn)
                                       .arg(QString::fromStdString(ContrastMessage.m_Contrast_KeyColumn)).arg(SearchValue));
                            while(query.next())
                            {
                                for(j=0;j<ContrastMessage.m_Contrast_Column.size();j++)
                                {
                                    for(k=0;k<ContrastMessage.m_Column.size();k++)
                                    {
                                        if(ContrastMessage.m_Column[k].m_ColumnName == ContrastMessage.m_Contrast_Column[j])
                                        {
                                            ContrastColumn = ContrastMessage.m_Column[k].m_ColumnNum;
                                            break;
                                        }
                                    }

                                    if((query.value(j).toString()) != (SeveralSheets[g][i][ContrastColumn-1]).toString())
                                    {
                                        DifferentJudge = false;
                                    }
                                }
                                if(DifferentJudge == false)
                                {
                                    int row = ui->tableWidget_2->rowCount();
                                    ui->tableWidget_2->insertRow(row);
                                    for(k=0;k<ContrastMessage.m_Column.size();k++)
                                    {
                                        ui->tableWidget_2->setItem(m,k,new QTableWidgetItem(SeveralSheets[g][i][ContrastMessage.m_Column[k].m_ColumnNum-1].toString()));
                                        ui->tableWidget_2->item(m,k)->setForeground(QBrush(QColor(color1,color2,color3)));
                                        ui->tableWidget_2->item(m,k)->setFont( QFont( "Times", 10, QFont::Black ) );
                                        ui->tableWidget_2->item(m,k)->setBackground(QBrush(QColor(245,255,250)));
                                    }
                                    for(n=0;n<ContrastMessage.m_Contrast_Column.size();n++)
                                    {
                                        ui->tableWidget_2->setItem(m,k,new QTableWidgetItem(query.value(n).toString()));
                                        ui->tableWidget_2->item(m,k)->setForeground(QBrush(QColor(color1,color2,color3)));
                                        ui->tableWidget_2->item(m,k)->setFont( QFont( "Times", 10, QFont::Black ) );
                                        ui->tableWidget_2->item(m,k)->setBackground(QBrush(QColor(245,255,250)));
                                        k++;
                                    }
                                    ui->tableWidget_2->setItem(m,k,new QTableWidgetItem(QString::number(g+1)));
                                    ui->tableWidget_2->item(m,k)->setForeground(QBrush(QColor(color1,color2,color3)));
                                    ui->tableWidget_2->item(m,k)->setFont( QFont( "Times", 10, QFont::Black ) );
                                    ui->tableWidget_2->item(m,k)->setBackground(QBrush(QColor(245,255,250)));
                                    m++;
                                }
                                DifferentJudge = true;
                            }
                        }
                        ui->tableWidget_2->resizeRowsToContents();//所有行高度自适应
                    }
               }
            }            
            QString strPath = QApplication::applicationDirPath();
            strPath += "/img/chai2.jpg";
            QMessageBox message(QMessageBox::NoIcon, "提示", "比对完成！");
            message.setIconPixmap(QPixmap(strPath));
            message.exec();
        }
    }
}
