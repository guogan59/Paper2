////////////////////////////////////////////////////////////////////////////
//                                                                        //
//                      Copyright (c) 2019-2020                           //
//             College of Aerospace Science and Engineering               //
//               National University of Defense Technology                //
//                                                                        //
////////////////////////////////////////////////////////////////////////////

// readexcel.h
//
//////////////////////////////////////////////////////////////////////
/// @file
/// @brief	重要功能子函数
#ifndef READEXCEL_H
#define READEXCEL_H
#pragma once

#include <QVariant>

//********************************************************************
/// 数据库信息结构体
/// @Author	Guo Shuai
/// @Date	2019-1-4
//********************************************************************
typedef struct
{
    int                                    m_ColumnNum;           //需读取Excel列数
    std::string                            m_ColumnName;          //上次更新数据库名称
}CcolumnMessage;

typedef struct
{
    std::string                            m_DataBaseName;        //上次更新数据库名称
    std::vector<CcolumnMessage>            m_Column;              //需读取Excel列数信息
}CUpdateMessage;

typedef struct
{
    std::string                            m_Contrast_KeyColumn;  //对比查找关键
    std::vector<std::string>               m_Contrast_Column;     //需对比的列数
    std::vector<CcolumnMessage>            m_Column;              //需读取Excel列数信息
}CContrastMessage;

//********************************************************************
/// Mode1 读取Excel文件类
/// @Author	Guo Shuai
/// @Date	2019-1-4
//********************************************************************
class CReadWrite
{
public:
    CReadWrite(){};
    ~CReadWrite(){};

public:
    //
    //操作函数
    //
    QVariant   Excel_Read(QString ExcelFile, int ReadSheet, QList<QList<QVariant> > &Datas);      //读取excel文件
    QVariant   Excel_SeveralSheets_Read(QString ExcelFile, int ReadSheet, QList<QList< QList<QVariant>>> &Datas);
    void       XML_Undate_Read(CUpdateMessage &UpdateMessage);
    void       XML_Undate_Modify(QString DataBaseName);
    void       XML_Contrast_Read(CContrastMessage &ContrastMessage);
};

#endif // READEXCEL_H
