#-------------------------------------------------
#
# Project created by QtCreator 2019-01-04T10:43:39
#
#-------------------------------------------------

QT       += core gui
QT       += core sql
QT       += axcontainer
QT       += core  xml
RC_FILE  = logo.rc
greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

TARGET = Data
TEMPLATE = app

# The following define makes your compiler emit warnings if you use
# any feature of Qt which as been marked as deprecated (the exact warnings
# depend on your compiler). Please consult the documentation of the
# deprecated API in order to know how to port your code away from it.
DEFINES  -= UNICODE
DEFINES += QT_DEPRECATED_WARNINGS


# You can also make your code fail to compile if you use deprecated APIs.
# In order to do so, uncomment the following line.
# You can also select to disable deprecated APIs only up to a certain version of Qt.
#DEFINES += QT_DISABLE_DEPRECATED_BEFORE=0x060000    # disables all the APIs deprecated before Qt 6.0.0


SOURCES += main.cpp\
        mainwindow.cpp \
    readexcel.cpp \
    Markup.cpp

HEADERS  += mainwindow.h \
    readexcel.h \
    Markup.h

FORMS    += mainwindow.ui

DISTFILES +=
