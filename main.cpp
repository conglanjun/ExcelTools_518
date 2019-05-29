#include "mainwindow.h"
#include <QApplication>
#include <QPushButton>
#include <QPushButton>
#include "pushBtn.h"
#include "excel.h"
#include "sortutils.h"
#include "mywindow.h"

#include <QDialog>
#include <QRect>
#include <QFont>
#include <QLineEdit>
#include <QGridLayout>
#include <QProgressBar>


void testFunction(){
//    QList<int> a;
//    a.append(49);
//    a.append(38);
//    a.append(65);
//    a.append(97);
//    a.append(76);
//    a.append(13);
//    a.append(27);
//    int temp[a.size()];
//    SortUtils util;
//    util.mergeSort(a, 0, a.size() - 1, temp);
//    for(int i = 0; i < a.size(); ++ i){
//        qDebug() << i << ":" << a.at(i) << "\n";
//    }
    QList<User_info> users;
    users.append(User_info(0, 49));
    users.append(User_info(1, 38));
    users.append(User_info(2, 65));
    users.append(User_info(3, 97));
    users.append(User_info(4, 76));
    users.append(User_info(5, 13));
    users.append(User_info(6, 27));
    QList<User_info> temp;
    temp.append(User_info(0, 0));
    temp.append(User_info(0, 0));
    temp.append(User_info(0, 0));
    temp.append(User_info(0, 0));
    temp.append(User_info(0, 0));
    temp.append(User_info(0, 0));
    temp.append(User_info(0, 0));
    SortUtils util;
    util.mergeSortStruct(users, 0, users.size() - 1, temp);
    for(int i = 0; i < users.size(); ++ i){
        qDebug() << i << ":" << users.at(i).row_num << "," << users.at(i).month_over_code << "\n";
    }
}

int main(int argc, char *argv[])
{
//    testFunction();
    QApplication a(argc, argv);
//    MainWindow mainWindow;

//    QDialog mainWindow;
    MyWindow mainWindow;
    mainWindow.resize(300, 150);
    QSizePolicy sizePolicy(QSizePolicy::Fixed, QSizePolicy::Fixed);
    sizePolicy.setHorizontalStretch(0);
    sizePolicy.setVerticalStretch(0);
    sizePolicy.setHeightForWidth(mainWindow.sizePolicy().hasHeightForWidth());
    mainWindow.setSizePolicy(sizePolicy);
    mainWindow.setMinimumSize(QSize(500, 300));
    mainWindow.setMaximumSize(QSize(500, 300));
    mainWindow.setSizeGripEnabled(false);

    QGridLayout mainLayout(&mainWindow);

    QProgressBar progressBar;
    QLabel label;

    excel e(&progressBar, &label);
    mainWindow.e = &e;
    pushbtn btn(&e, &mainWindow);
    btn.setText("选择输入文件");
    btn.setGeometry(QRect(20, 40, 80, 40));

    mainLayout.addWidget(&label, 0, 1);
    mainLayout.addWidget(&btn, 1, 0);
    mainLayout.addWidget(&progressBar, 1, 1);

    mainWindow.setWindowTitle("月结工具");
    mainWindow.show();

    return a.exec();
}
