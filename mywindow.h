#ifndef MYWINDOW_H
#define MYWINDOW_H
#include <QDialog>
#include <QDebug>

#include "excel.h"


class MyWindow : public QDialog
{
    Q_OBJECT

public:
    explicit MyWindow(QWidget *parent = 0);
    ~MyWindow(){

    }

    excel *e;

protected:
    //这是一个虚函数，继承自QEvent.只要重写了这个虚函数，当你按下窗口右上角的"×"时，就会调用你所重写的此函数.
    void closeEvent(QCloseEvent*event);
};

#endif // MYWINDOW_H
