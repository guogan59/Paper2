#include "readexcel.h"
#include <QAxObject>
#include <QDebug>
#include <QMessageBox>
#include <QFile>
#include <QXmlStreamReader>
#include <fstream>
#include <iostream>
#include <tchar.h>
#include "Markup.h"
#include <QApplication>
# pragma execution_character_set("utf-8")
using namespace std;

//********************************************************************
/// 读取excel文件函数
/// @Author	Guo Shuai
/// @Date	2019-1-4
/// @Input
/// @Param  xlsFile     excel文件路径
/// @Param  sheetNum    需读取的sheet号码
/// @Param  res         待转换QList
/// @Output
//********************************************************************
QVariant CReadWrite::Excel_Read(QString ExcelFile, int ReadSheet, QList<QList<QVariant> > &Datas)
{
    QAxObject excel("Excel.Application");
    excel.setControl("Excel.Application");                                              //连接Excel控件
    excel.dynamicCall("SetVisible (bool Visible)","false");                             //设置为不显示窗体
    excel.setProperty("DisplayAlerts", false);                                          //不显示任何警告信息，如关闭时的是否保存提示
    //excel操作
    QAxObject *workbooks = excel.querySubObject("WorkBooks");                           //获取工作簿(excel文件)集合
    workbooks->dynamicCall("Open(const QString&)", ExcelFile);

    QAxObject *workbook = excel.querySubObject("ActiveWorkBook");
    QAxObject *worksheets = workbook->querySubObject("Sheets");                         //获取所有工作表

    int sheet_count = worksheets->property("Count").toInt();                            //获取工作表数目
    if (ReadSheet > sheet_count)
    {
        QMessageBox::warning(NULL, "警告", "Sheet序号不存在!");
    }
    else if (sheet_count > 0)
    {
        QAxObject *work_sheet = worksheets->querySubObject("Item(int)",ReadSheet);
        //读取sheet的数据
        QVariant var;
        if (work_sheet != NULL && ! work_sheet->isNull())
        {
           QAxObject *usedRange = work_sheet->querySubObject("UsedRange");
           if(NULL == usedRange || usedRange->isNull())
           {
               return var;
           }
           var = usedRange->dynamicCall("Value");
           delete usedRange;
        }
        //将QVariant转为Qlist
        QVariantList varRows = var.toList();
        if(varRows.isEmpty())
        {
           return 0;
        }
        const int rowCount = varRows.size();
        QVariantList rowData;
        for(int i=0;i<rowCount;++i)
        {
           rowData = varRows[i].toList();
           Datas.push_back(rowData);
        }
    }

    workbook->dynamicCall("Close (Boolean)", false); //关闭文件
    excel.dynamicCall("Quit(void)");//关闭excel
    delete worksheets;
    delete workbook;
    delete workbooks;
    return true;
}
//********************************************************************
/// 读取xml文件,明确需读取列数
/// @Author	Guo Shuai
/// @Date	2019-1-4
/// @Input
/// @Param  ColumnNum     读取excel列数
/// @Output
//********************************************************************
void CReadWrite::XML_Undate_Read(CUpdateMessage &DataBaseMessage)
{
    CMarkup                InputFile;                    //xml文件
    CcolumnMessage         ColumnMessage;                //读取列数信息   
    QString                XmlName = QApplication::applicationDirPath();
    XmlName += "/InitialFile/Update.xml"; //设置显示格式
    InputFile.Load(_TEXT(XmlName.toStdString()));
    if (!InputFile.IsWellFormed())
    {
        QMessageBox::warning(NULL, "警告", "该文件不是一个有效的XML格式文件!");
        return;		//该文件不是一个有效的InputFile格式文件！
    }
    if (!InputFile.FindElem(_TEXT("DataBaseMessage")))
    {
        QMessageBox::warning(NULL, "警告", "XML格式出错!");
        return;
    }
    InputFile.IntoElem();
    {
        InputFile.FindElem(_T("DataBaseName"));
        DataBaseMessage.m_DataBaseName = InputFile.GetData();
        while(InputFile.FindElem(_T("column")))
        {
            InputFile.IntoElem();
            {
                InputFile.FindElem(_T("columnNum"));
                ColumnMessage.m_ColumnNum = QString::fromStdString(InputFile.GetData()).toInt();

                InputFile.FindElem(_T("columnName"));
                ColumnMessage.m_ColumnName = InputFile.GetData();

                DataBaseMessage.m_Column.push_back(ColumnMessage);
            }
            InputFile.OutOfElem();
        }
    }
    InputFile.OutOfElem();
}
//********************************************************************
/// 读取xml文件,明确需读取列数
/// @Author	Guo Shuai
/// @Date	2019-1-4
/// @Input
/// @Param  ColumnNum     读取excel列数
/// @Output
//********************************************************************
void CReadWrite::XML_Undate_Modify(QString DataBaseName)
{
    CMarkup                InputFile;                    //xml文件
    QString                XmlName = QApplication::applicationDirPath();
    XmlName += "/InitialFile/Update.xml"; //设置显示格式
    InputFile.Load(_TEXT(XmlName.toStdString()));
    if (!InputFile.IsWellFormed())
    {
        QMessageBox::warning(NULL, "警告", "该文件不是一个有效的XML格式文件!");
        return;		//该文件不是一个有效的InputFile格式文件！
    }
    if (!InputFile.FindElem(_TEXT("DataBaseMessage")))
    {
        QMessageBox::warning(NULL, "警告", "XML格式出错!");
        return;
    }
    InputFile.IntoElem();
    {
        InputFile.FindElem(_T("DataBaseName"));
        InputFile.SetData(DataBaseName.toStdString());
    }
    InputFile.OutOfElem();
    InputFile.Save(_TEXT(XmlName.toStdString()));//默认创建路径为该工程的根目录
}
//********************************************************************
/// 读取比对xml文件,明确需读取列数
/// @Author	Guo Shuai
/// @Date	2019-1-8
/// @Input
/// @Param  DataBaseMessage     比对excel信息
/// @Output
//********************************************************************
void CReadWrite::XML_Contrast_Read(CContrastMessage &ContrastMessage)
{
    CMarkup                InputFile;                    //xml文件
    CcolumnMessage         ColumnMessage;                //读取列数信息

    QString                XmlName = QApplication::applicationDirPath();
    XmlName += "/InitialFile/Contrast.xml"; //设置显示格式
    InputFile.Load(_TEXT(XmlName.toStdString()));
    if (!InputFile.IsWellFormed())
    {
        QMessageBox::warning(NULL, "警告", "该文件不是一个有效的XML格式文件!");
        return;		//该文件不是一个有效的InputFile格式文件！
    }
    if (!InputFile.FindElem(_TEXT("DataBaseMessage")))
    {
        QMessageBox::warning(NULL, "警告", "XML格式出错!");
        return;
    }
    InputFile.IntoElem();
    {
        InputFile.FindElem(_T("keycolumn"));
        ContrastMessage.m_Contrast_KeyColumn = InputFile.GetData();
        while(InputFile.FindElem(_T("contrastcolumn")))
        {
            ContrastMessage.m_Contrast_Column.push_back(InputFile.GetData());
        }

        while(InputFile.FindElem(_T("column")))
        {
            InputFile.IntoElem();
            {
                InputFile.FindElem(_T("columnNum"));
                ColumnMessage.m_ColumnNum = QString::fromStdString(InputFile.GetData()).toInt();

                InputFile.FindElem(_T("columnName"));
                ColumnMessage.m_ColumnName = InputFile.GetData();

                ContrastMessage.m_Column.push_back(ColumnMessage);
            }
            InputFile.OutOfElem();
        }
    }
    InputFile.OutOfElem();
}
//********************************************************************
/// 读取excel文件函数
/// @Author	Guo Shuai
/// @Date	2019-1-4
/// @Input
/// @Param  xlsFile     excel文件路径
/// @Param  sheetNum    需读取的sheet号码
/// @Param  res         待转换QList
/// @Output
//********************************************************************
QVariant CReadWrite::Excel_SeveralSheets_Read(QString ExcelFile, int ReadSheet, QList<QList< QList<QVariant>>> &Datas)
{
    int      i=0;
    QAxObject excel("Excel.Application");
    excel.setControl("Excel.Application");                                              //连接Excel控件
    excel.dynamicCall("SetVisible (bool Visible)","false");                             //设置为不显示窗体
    excel.setProperty("DisplayAlerts", false);                                          //不显示任何警告信息，如关闭时的是否保存提示
    //excel操作
    QAxObject *workbooks = excel.querySubObject("WorkBooks");                           //获取工作簿(excel文件)集合
    workbooks->dynamicCall("Open(const QString&)", ExcelFile);

    QAxObject *workbook = excel.querySubObject("ActiveWorkBook");
    QAxObject *worksheets = workbook->querySubObject("Sheets");                         //获取所有工作表

    int sheet_count = worksheets->property("Count").toInt();                            //获取工作表数目
    if (ReadSheet > sheet_count)
    {
        QMessageBox::warning(NULL, "警告", "Sheet序号不存在!");
    }
    else if (sheet_count > 0)
    {
        for(i=0;i<ReadSheet;i++)
        {
            QList< QList<QVariant>> SingleDatas;
            QAxObject *work_sheet = worksheets->querySubObject("Item(int)",i+1);
            //读取sheet的数据
            QVariant var;
            if (work_sheet != NULL && ! work_sheet->isNull())
            {
               QAxObject *usedRange = work_sheet->querySubObject("UsedRange");
               if(NULL == usedRange || usedRange->isNull())
               {
                   return var;
               }
               var = usedRange->dynamicCall("Value");
               delete usedRange;
            }
            //将QVariant转为Qlist
            QVariantList varRows = var.toList();
            if(varRows.isEmpty())
            {
               return 0;
            }
            const int rowCount = varRows.size();
            QVariantList rowData;
            for(int i=0;i<rowCount;++i)
            {
               rowData = varRows[i].toList();
               SingleDatas.push_back(rowData);
            }
            Datas.push_back(SingleDatas);
        }
    }

    workbook->dynamicCall("Close (Boolean)", false); //关闭文件
    excel.dynamicCall("Quit(void)");//关闭excel
    delete worksheets;
    delete workbook;
    delete workbooks;
    return true;
}
