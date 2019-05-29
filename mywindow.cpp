#include "mywindow.h"

MyWindow::MyWindow(QWidget *parent)
{

}

void MyWindow::closeEvent(QCloseEvent*event)
{
    qDebug() << "my windows close!!";
//    WorkThread * wt = e->wt;
//    wt->exit(0);
//    wt->close();
}
