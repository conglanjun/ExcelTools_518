#ifndef SORTUTILS_H
#define SORTUTILS_H
#include "qlist.h"
#include <ActiveQt/QAxObject>
#include <QDateTime>
#include <QDebug>

struct User_info{
    int row_num;
    int month_over_code;
    QString username;

    User_info(int row_num = 0, int month_over_code = 0, QString username = "")
        : row_num(row_num), month_over_code(month_over_code), username(username){}
};

class SortUtils
{
public:
    SortUtils();

    void merge(QList<int> &resultList, int low, int mid, int high, int* tempList);

    void mergeSort(QList<int> &resultList, int low, int high, int* tempList);

    void mergeStruct(QList<User_info> &resultList, int low, int mid, int high, QList<User_info> &tempList);

    void mergeSortStruct(QList<User_info> &resultList, int low, int high, QList<User_info> &tempList);

    int binary_search(QList<User_info> userList, int key);

    int binary_search2(QVariantList DM_rows, int column, int data_row, int count, qint64 key);

    QList<int> rowList;
};

#endif // SORTUTILS_H
