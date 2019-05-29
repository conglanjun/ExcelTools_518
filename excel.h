#ifndef EXCEL_H
#define EXCEL_H
#include <ActiveQt/QAxObject>
#include <QFileDialog>
#include <QMessageBox>
#include <QVector>
#include <QDebug>
#include <QProgressDialog>
#include <QProgressBar>
#include <QDateTime>
#include <QLabel>
#include "workthread.h"
//#include "pushBtn.h"
class WorkThread;

class excel : public QWidget
{
public:
    excel(){}
    excel(QProgressBar *bar):progressBar(bar)
    {

    }
    excel(QProgressBar *bar, QLabel *label):progressBar(bar), label(label)
    {

    }

public:
    void excelExport();
public:
    void excelImport();
public:
    void excelImportDemo();
    void setCellValue(QAxObject *work_sheet, int row, QAxObject *data_sheet, int data_row, bool isDouble, int index);
public slots:
    void send_over_cmd(QString path);
    void send_cmd(QString path);
    void receive_row_count(int row_count, QString detail);
    void receive_row_detail_count(int row_count);
    void receive_row_done();
public slots:
    void ClickButton();
private:
    QProgressDialog * qpDialog;
public:
    int row_count;
    WorkThread *wt;
    QProgressBar *progressBar;
    QLabel *label;
};

void importDemo(excel e);

#endif // EXCEL_H

