#-------------------------------------------------
#
# Project created by QtCreator 2018-01-22T10:23:08
#
#-------------------------------------------------

QT       += core gui
QT       += core gui axcontainer
greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

TARGET = ExcelTools_518
TEMPLATE = app


SOURCES += main.cpp\
        mainwindow.cpp \
    excel.cpp \
    pushbtn.cpp \
    workthread.cpp \
    sortutils.cpp \
    mywindow.cpp

HEADERS  += \
    workthread.h \
    excel.h \
    pushBtn.h \
    mainwindow.h \
    sortutils.h \
    mywindow.h

FORMS    += mainwindow.ui
