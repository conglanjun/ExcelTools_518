#ifndef WORKTHREAD_H
#define WORKTHREAD_H
#include <QThread>
#include "excel.h"
#include "sortutils.h"

struct Point_part{
    QString net_code;
    QString branch_part;
    QString finance_person;
    QString receive_person;
    QList<int> start;
    QList<int> end;
    int bill_no_use_count;
    int bill_use_count;
    int bill_sum;
    float bill_percent;
    int receipt_no_use_count;
    int receipt_use_count;
    int receipt_sum;
    float receipt_percent;

    Point_part(QString code, QString part, QString finance, QString receive){
        net_code = code;
        branch_part = part;
        finance_person = finance;
        receive_person = receive;
    }
    Point_part(){}
};

struct Key_follow{
    QList<int> start;
    QList<int> end;
    QString username;
    QString branch_part;
    QString type;

    Key_follow(QList<int> s, QList<int> e, QString u){
        start = s;
        end = e;
        username = u;
    }
    Key_follow(){}
};

struct Finance_receive{
    QList<int> start;
    QList<int> end;
    QString fr;

    Finance_receive(QList<int> s, QList<int> e, QString v){
        start = s;
        end = e;
        fr = v;
    }
    Finance_receive(){}
};


class WorkThread : public QThread
{
    Q_OBJECT
public:
    WorkThread(const QStringList & paths) : m_paths(paths){}

    QStringList getPath()
    {
        return m_paths;
    }
    ~WorkThread(){
        close();
        qDebug() << "work thread deconstruct!!";
    }
    void close();
    void clearPP(Point_part &pp){
        pp.net_code = "";
        pp.branch_part = "";
        pp.finance_person = "";
        pp.receive_person = "";
        pp.bill_no_use_count = 0;
        pp.bill_use_count = 0;
        pp.bill_sum = 0;
        pp.bill_percent = 0.0;
        pp.receipt_no_use_count = 0;
        pp.receipt_use_count = 0;
        pp.receipt_sum = 0;
        pp.receipt_percent = 0.0;
    }
    void castListListVariant2Variant(QVariant &var, const QList<QList<QVariant>> &res)
    {
        QVariant temp = QVariant(QVariantList());
        QVariantList record;

        int listSize = res.size();
        for (int i = 0; i < listSize;++i)
        {
            temp = res.at(i);
            record << temp;
        }
        temp = record;
        var = temp;
    }

    QAxObject *getExcel1(){
        return excel1;
    }
    QAxObject *getWorkBook1(){
        return work_book1;
    }
    QAxObject *getExcel2(){
        return excel2;
    }
    QAxObject *getWorkBook2(){
        return work_book2;
    }
    QAxObject *getExcel3(){
        return excel3;
    }
    QAxObject *getWorkBook3(){
        return work_book3;
    }
    QAxObject *getExcel4(){
        return excel4;
    }
    QAxObject *getWorkBook4(){
        return work_book4;
    }
    QAxObject *getExcel5(){
        return excel5;
    }
    QAxObject *getWorkBook5(){
        return work_book5;
    }
    QAxObject *getExcel6(){
        return excel6;
    }
    QAxObject *getWorkBook6(){
        return work_book6;
    }

//自定义信号
signals:
    void send_export_over_signal(QString path);
    void send_export_signal(QString path);
    void send_excel_row_done();
    void send_excel_row_count(int row_count, QString detail);
    void send_excel_row_detail_count(int row_count);
    void send_btn_enable(bool flag);

protected:
    void run();
private:
    QStringList m_paths;
    QAxObject *excel1;
    QAxObject *work_book1;
    QAxObject *excel2;
    QAxObject *work_book2;
    QAxObject *excel3;
    QAxObject *work_book3;
    QAxObject *excel4;
    QAxObject *work_book4;
    QAxObject *excel5;
    QAxObject *work_book5;
    QAxObject *excel6;
    QAxObject *work_book6;
};

const QString NO_USE = "未使用";
const QString HAVE_USE = "已使用";
const QString YES = "是";
const QString NO = "否";

#endif // WORKTHREAD_H
//class Newspaper : public QThread
//{
//    Q_OBJECT
//public :
//    Newspaper(const QString & name) : m_name(name){}
//    void send()
//    {
//        emit newPaper(m_name);
//    }
//signals:
//    void newPaper(const QString &name);

//private :
//    QString m_name;
//};
