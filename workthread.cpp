#include "workthread.h"
#include "qt_windows.h"
#include <QDebug>

void setCellValue(QAxObject *work_sheet, int row, int column, QString data)
{
    QAxObject *cell = work_sheet->querySubObject("Cells(int,int)", row, column);
    if(row == 12){
        if(column == 5 || column == 6){
            QAxObject* interior = cell->querySubObject("Interior");
            interior->setProperty("Color", QColor(255, 255, 0));   //设置单元格背景色（黄色）
            QAxObject *font = cell->querySubObject("Font");
            font->setProperty("Color", QColor(0, 0, 0));
        }else{
            QAxObject* interior = cell->querySubObject("Interior");
            interior->setProperty("Color", QColor(0, 0, 255));   //设置单元格背景色（绿色）
            QAxObject *font = cell->querySubObject("Font");
            font->setProperty("Color", QColor(255, 255, 255));
        }
    }
    // 寄件公司地址 收件公司地址 两个成本中心
    cell->setProperty("Value", data);  //设置单元格值
    if(row == 12){
        if(column == 1 || column == 3){
            cell->setProperty("ColumnWidth", 60);
        }else if (column == 2 || column == 4 || column == 5 || column == 6){
            cell->setProperty("ColumnWidth", 16);
        }
    }
}


void WorkThread::run ()
{
    CoInitializeEx(NULL, COINIT_MULTITHREADED);

    QStringList paths = getPath();//得到用户选择的文件列表
    QString path1 = "";
    QString path2 = "";
    QString path3 = "";
    QString path4 = "";
    QString path5 = "";
    QString path6 = "";
    /**
     * 1.月结管家数据分析5.16.xlsx等相关文件
     * 2.DM_SHARE_获取账单明细表 (账单).xlsx
     * 3.SHARE_portal获取账单明细表_PC (账单).xlsx
     * 4.SHARE_申请发票明细表 (发票).xlsx
     * 5.2019-05-17-09-32-10.1768.xls
     * 6.对账人及客户信息报表_522143 (发票订阅).xlsx
     **/
    for(int i = 0; i < paths.length(); ++i){
        QString str_path = paths[i];
        //单个文件路径
        QFileInfo file = QFileInfo(str_path);
        //获得文件名
        QString file_name = file.fileName();
        if(file_name.startsWith("1.")){
            path1 = str_path;
        }else if(file_name.startsWith("2.")){
            path2 = str_path;
        }else if(file_name.startsWith("3.")){
            path3 = str_path;
        }else if(file_name.startsWith("4.")){
            path4 = str_path;
        }else if(file_name.startsWith("5.")){
            path5 = str_path;
        }else if(file_name.startsWith("6.")){
            path6 = str_path;
        }
    }
    excel1 = new QAxObject("Excel.Application");
    QAxObject *work_books1 = excel1->querySubObject("WorkBooks");
    work_books1->dynamicCall("Open(const QString&)", path1);
    excel1->setProperty("Caption", "Qt Excel");
    work_book1 = excel1->querySubObject("ActiveWorkBook");
    QAxObject *work_sheets1 = work_book1->querySubObject("Sheets");  //Sheets也可换用WorkSheets

    // 客户明细
    QAxObject *user_info_sheet = work_sheets1->querySubObject("Item(int)", 8);
    QAxObject *user_usedRange = user_info_sheet->querySubObject("UsedRange");
    QAxObject *user_rows = user_usedRange->querySubObject("Rows");
    int user_row_count = user_rows->property("Count").toInt();

    int code_row = 4;// 记录数据行数，为了显示进度条 要减去3
    emit send_excel_row_count(user_row_count - code_row + 1, "处理客户明细 2.DM_SHARE,3.SHARE_portal PC,4.发票,5.订阅,6.对账人");

    excel2 = new QAxObject("Excel.Application");
    QAxObject *work_books2 = excel2->querySubObject("WorkBooks");
    work_books2->dynamicCall("Open(const QString&)", path2);
    excel2->setProperty("Caption", "Qt Excel");
    work_book2 = excel2->querySubObject("ActiveWorkBook");
    QAxObject *work_sheets2 = work_book2->querySubObject("Sheets");  //Sheets也可换用WorkSheets

    // DM
    QAxObject *DM_SHARE_sheet = work_sheets2->querySubObject("Item(int)", 1);
    QAxObject *DM_usedRange = DM_SHARE_sheet->querySubObject("UsedRange");
    QAxObject *rows = DM_usedRange->querySubObject("Rows");
//    QAxObject *columns = DM_usedRange->querySubObject("Columns");
    int intRows = rows->property("Count").toInt();
//    int intColumns = columns->property("Count").toInt();
//    qDebug() << "int rows:" << intRows << ",columns:" << intColumns << "\n";

    excel3 = new QAxObject("Excel.Application");
    QAxObject *work_books3 = excel3->querySubObject("WorkBooks");
    work_books3->dynamicCall("Open(const QString&)", path3);
    excel3->setProperty("Caption", "Qt Excel");
    work_book3 = excel3->querySubObject("ActiveWorkBook");
    QAxObject *work_sheets3 = work_book3->querySubObject("Sheets");  //Sheets也可换用WorkSheets

    // SHARE PC
    QAxObject *PC_SHARE_sheet = work_sheets3->querySubObject("Item(int)", 1);
    QAxObject *PC_usedRange = PC_SHARE_sheet->querySubObject("UsedRange");
    QAxObject *PC_rows = PC_usedRange->querySubObject("Rows");
//    QAxObject *PC_columns = PC_usedRange->querySubObject("Columns");
    int PC_row_count = PC_rows->property("Count").toInt();
//    int PC_columns_count = PC_columns->property("Count").toInt();
//    qDebug() << "pc rows:" << PC_row_count << ",pc columns:" << PC_columns_count;

    excel4 = new QAxObject("Excel.Application");
    QAxObject *work_books4 = excel4->querySubObject("WorkBooks");
    work_books4->dynamicCall("Open(const QString&)", path4);
    excel4->setProperty("Caption", "Qt Excel");
    work_book4 = excel4->querySubObject("ActiveWorkBook");
    QAxObject *work_sheets4 = work_book4->querySubObject("Sheets");  //Sheets也可换用WorkSheets

    // SHARE RE
    QAxObject *RE_SHARE_sheet = work_sheets4->querySubObject("Item(int)", 1);
    QAxObject *RE_usedRange = RE_SHARE_sheet->querySubObject("UsedRange");
    QAxObject *RE_rows = RE_usedRange->querySubObject("Rows");
//    QAxObject *RE_columns = RE_usedRange->querySubObject("Columns");
    int RE_row_count = RE_rows->property("Count").toInt();
//    int RE_columns_count = RE_columns->property("Count").toInt();
//    qDebug() << "re rows:" << RE_row_count << ",re columns:" << RE_columns_count;

    excel5 = new QAxObject("Excel.Application");
    QAxObject *work_books5 = excel5->querySubObject("WorkBooks");
    work_books5->dynamicCall("Open(const QString&)", path5);
    excel5->setProperty("Caption", "Qt Excel");
    work_book5 = excel5->querySubObject("ActiveWorkBook");
    QAxObject *work_sheets5 = work_book5->querySubObject("Sheets");  //Sheets也可换用WorkSheets

    // SHARE SU
    QAxObject *SU_SHARE_sheet = work_sheets5->querySubObject("Item(int)", 1);
    QAxObject *SU_usedRange = SU_SHARE_sheet->querySubObject("UsedRange");
    QAxObject *SU_rows = SU_usedRange->querySubObject("Rows");
//    QAxObject *SU_columns = SU_usedRange->querySubObject("Columns");
    int SU_row_count = SU_rows->property("Count").toInt();
//    int SU_columns_count = SU_columns->property("Count").toInt();
//    qDebug() << "su rows:" << SU_row_count << ",su columns:" << SU_columns_count;

    excel6 = new QAxObject("Excel.Application");
    QAxObject *work_books6 = excel6->querySubObject("WorkBooks");
    work_books6->dynamicCall("Open(const QString&)", path6);
    excel6->setProperty("Caption", "Qt Excel");
    work_book6 = excel6->querySubObject("ActiveWorkBook");
    QAxObject *work_sheets6 = work_book6->querySubObject("Sheets");  //Sheets也可换用WorkSheets

    // SHARE AC
    QAxObject *AC_SHARE_sheet = work_sheets6->querySubObject("Item(int)", 1);
    QAxObject *AC_usedRange = AC_SHARE_sheet->querySubObject("UsedRange");
    QAxObject *AC_rows = AC_usedRange->querySubObject("Rows");
//    QAxObject *AC_columns = AC_usedRange->querySubObject("Columns");
    int AC_row_count = AC_rows->property("Count").toInt();
//    int AC_columns_count = AC_columns->property("Count").toInt();
//    qDebug() << "ac rows:" << AC_row_count << ",ac columns:" << AC_columns_count;

    SortUtils util;

    int DM_PC_RE_column = 0;
    int DM_PC_RE_row = 0;

    QVariantList params;
    QString end("G");
    QString end_number = QString::number(user_row_count);
    params << "G2" << QVariant::fromValue(end + end_number);// 月结号

    QAxObject *used_range = user_usedRange->querySubObject("Range(QVariant, QVariant)", params);
    QVariant used_data = used_range->dynamicCall("Value");
    QVariantList tempList_row = used_data.value<QVariantList>();

    // DM 客户卡号
    QVariantList DM_params;
    QString DM_end("D");
    QString DM_end_number = QString::number(intRows);
    DM_params << "D4" << QVariant::fromValue(DM_end + DM_end_number);// 客户卡号

    QAxObject *DM_range = DM_usedRange->querySubObject("Range(QVariant, QVariant)", DM_params);
    QVariant DM_data = DM_range->dynamicCall("Value");
    QVariantList DM_rows = DM_data.value<QVariantList>();

    // PC 客户卡号
    QVariantList PC_params;
    QString PC_end("D");
    QString PC_end_number = QString::number(PC_row_count);
    PC_params << "D4" << QVariant::fromValue(PC_end + PC_end_number);// 客户卡号

    QAxObject *PC_range = PC_usedRange->querySubObject("Range(QVariant, QVariant)", PC_params);
    QVariant PC_data = PC_range->dynamicCall("Value");
    QVariantList PC_data_rows = PC_data.value<QVariantList>();

    // 发票 月结账号
    QVariantList RE_params;
    QString RE_end("F");
    QString RE_end_number = QString::number(RE_row_count);
    RE_params << "F4" << QVariant::fromValue(RE_end + RE_end_number);// 客户卡号

    QAxObject *RE_range = RE_usedRange->querySubObject("Range(QVariant, QVariant)", RE_params);
    QVariant RE_data = RE_range->dynamicCall("Value");
    QVariantList RE_data_rows = RE_data.value<QVariantList>();

    // 账单订阅 月结卡号
    QVariantList SU_params;
    QString SU_end("C");
    QString SU_end_number = QString::number(SU_row_count);
    SU_params << "B2" << QVariant::fromValue(SU_end + SU_end_number);// 客户卡号

    QAxObject *SU_range = SU_usedRange->querySubObject("Range(QVariant, QVariant)", SU_params);
    QVariant SU_data = SU_range->dynamicCall("Value");
    QVariantList SU_data_rows = SU_data.value<QVariantList>();


    // 对账人客户 客户卡号
    QVariantList AC_params;
    QString AC_end("D");
    QString AC_end_number = QString::number(AC_row_count);
    AC_params << "D4" << QVariant::fromValue(AC_end + AC_end_number);// 客户卡号

    QAxObject *AC_range = AC_usedRange->querySubObject("Range(QVariant, QVariant)", AC_params);
    QVariant AC_data = AC_range->dynamicCall("Value");
    QVariantList AC_data_rows = AC_data.value<QVariantList>();

    QVariantList AC_data_params;
    QString AC_data_end("P");
    QString AC_data_end_number = QString::number(AC_row_count);
    AC_data_params << "M4" << QVariant::fromValue(AC_data_end + AC_data_end_number);

    QAxObject *AC_data_range = AC_usedRange->querySubObject("Range(QVariant, QVariant)", AC_data_params);
    QVariant AC_data_data = AC_data_range->dynamicCall("Value");
    QVariantList AC_data_data_rows = AC_data_data.value<QVariantList>();

    QList<QList<QVariant>> bill_segments;
    bill_segments.reserve(user_row_count - 1);
    qint64 start_stamp = QDateTime::currentDateTime().toMSecsSinceEpoch();
    for(int i=0; i < tempList_row.size(); ++i)
    {
        QList<QVariant> billList;
        QVariantList tempList_col = tempList_row[i].value<QVariantList>();
        for(int j=0; j < tempList_col.size(); ++j){
            QString month_over_code_str = tempList_col[j].toString();
            int index = util.binary_search2(DM_rows, DM_PC_RE_column, DM_PC_RE_row, intRows - 4, month_over_code_str.toLongLong());
//            qDebug() << "row:" << i << ",index:" << index << ",code:" << month_over_code_str;
            if(index >= 0){
                billList.append(QVariant::fromValue(HAVE_USE));
            }else{// 没有就要去PC查询
                index = util.binary_search2(PC_data_rows, DM_PC_RE_column, DM_PC_RE_row, PC_row_count - 4, month_over_code_str.toLongLong());
//                qDebug() << "PC row:" << i << ",index:" << index << ",code:" << month_over_code_str;
                if(index >= 0){
                    billList.append(QVariant::fromValue(HAVE_USE));
                }else{
                    billList.append(QVariant::fromValue(NO_USE));
                }
            }
            // 处理4.发票
            index = util.binary_search2(RE_data_rows, DM_PC_RE_column, DM_PC_RE_row, RE_row_count - 4, month_over_code_str.toLongLong());
//            qDebug() << "RE row:" << i << ",index:" << index << ",code:" << month_over_code_str;
            if(index >= 0){
                billList.append(QVariant::fromValue(HAVE_USE));
            }else{
                billList.append(QVariant::fromValue(NO_USE));
            }
            // 处理5.账单自动订阅
            index = util.binary_search2(SU_data_rows, DM_PC_RE_column, DM_PC_RE_row, SU_row_count - 2, month_over_code_str.toLongLong());
//            qDebug() << "SU row:" << i << ",index:" << index << ",code:" << month_over_code_str;
            if(index > 0){
                QVariantList tempList_col = SU_data_rows[index].value<QVariantList>();
                QString cell_key = tempList_col[DM_PC_RE_column + 1].toString();
                if(cell_key == YES){
                    billList.append(QVariant::fromValue(HAVE_USE));
                }else{
                    billList.append(QVariant::fromValue(NO_USE));
                }
            }else{
                billList.append(QVariant::fromValue(NO_USE));
            }
            // 处理6.对账人 客户
            index = util.binary_search2(AC_data_rows, DM_PC_RE_column, DM_PC_RE_row, AC_row_count - 4, month_over_code_str.toLongLong());
            qDebug() << "row:" << i << ",code:" << month_over_code_str;
            if(index >= 0){
                QString cell_key = AC_data_data_rows[index].value<QVariantList>()[0].toString();// M列 是否自动开票
                QString date = AC_data_data_rows[index].value<QVariantList>()[1].toString();// N列 开票日期
                QString open = AC_data_data_rows[index].value<QVariantList>()[3].toString();// P列 开通日期
                if(cell_key == YES){
                    billList.append(QVariant::fromValue(YES));
                    billList.append(QVariant::fromValue(date));
                    billList.append(QVariant::fromValue(open));
                }else{
                    billList.append(QVariant::fromValue(NO));
                }
            }else{
                billList.append(QVariant::fromValue(NO));
            }
            emit send_excel_row_done();
        }
        bill_segments.append(billList);
    }
    qint64 end_stamp = QDateTime::currentDateTime().toMSecsSinceEpoch();
    qDebug() << "----stamp:" << (end_stamp - start_stamp);

    QVariantList bill_segment_params;
    QString bill_segment_end("L");
    QString bill_segment_end_number = QString::number(user_row_count);
    bill_segment_params << "Q2" << QVariant::fromValue(bill_segment_end + bill_segment_end_number);// 账单环节 发票环节 账单自动订阅 发票订阅 自动开票日 开通日期

    QVariant bill_var;
    castListListVariant2Variant(bill_var, bill_segments);

    QAxObject *bill_segment_range = user_usedRange->querySubObject("Range(QVariant, QVariant)", bill_segment_params);
    bill_segment_range->setProperty("Value", bill_var);

    work_book2->dynamicCall("Save()");
    work_book2->dynamicCall("Close(Boolean)", false);
    excel2->dynamicCall("Quit(void)");
    delete excel2;

    work_book3->dynamicCall("Save()");
    work_book3->dynamicCall("Close(Boolean)", false);
    excel3->dynamicCall("Quit(void)");
    delete excel3;

    work_book4->dynamicCall("Save()");
    work_book4->dynamicCall("Close(Boolean)", false);
    excel4->dynamicCall("Quit(void)");
    delete excel4;

    work_book5->dynamicCall("Save()");
    work_book5->dynamicCall("Close(Boolean)", false);
    excel5->dynamicCall("Quit(void)");
    delete excel5;

    work_book6->dynamicCall("Save()");
    work_book6->dynamicCall("Close(Boolean)", false);
    excel6->dynamicCall("Quit(void)");
    delete excel6;

    emit send_excel_row_count(user_row_count - code_row + 1, "处理跟进 重点跟进订阅");

    // 重点跟进订阅
    QAxObject *kf_sheet = work_sheets1->querySubObject("Item(int)", 6);
    QAxObject *kf_usedRange = kf_sheet->querySubObject("UsedRange");
    // 重点跟进 账单环节 发票环节
    QAxObject *kf_br_sheet = work_sheets1->querySubObject("Item(int)", 5);
    QAxObject *kf_br_usedRange = kf_br_sheet->querySubObject("UsedRange");
    // 点部网点代码 账单环节 发票环节
    QAxObject *nc_sheet = work_sheets1->querySubObject("Item(int)", 4);
    QAxObject *nc_usedRange = nc_sheet->querySubObject("UsedRange");
    // 分部
    QAxObject *bp_sheet = work_sheets1->querySubObject("Item(int)", 3);
    QAxObject *bp_usedRange = bp_sheet->querySubObject("UsedRange");
    // 账务中心
    QAxObject *fr_sheet = work_sheets1->querySubObject("Item(int)", 2);
    QAxObject *fr_usedRange = fr_sheet->querySubObject("UsedRange");
    // 应收
    QAxObject *re_sheet = work_sheets1->querySubObject("Item(int)", 1);
    QAxObject *re_usedRange = re_sheet->querySubObject("UsedRange");

    QVariantList kf_params;
    QString kf_end("A");
    QString kf_end_number = QString::number(user_row_count);
    kf_params << "F2" << QVariant::fromValue(kf_end + kf_end_number);// 重点跟进 A-F分类 重点跟进 财务人员 应收 分部 网点代码

    QAxObject *kf_used_range = user_usedRange->querySubObject("Range(QVariant, QVariant)", kf_params);
    QVariant kf_used_data = kf_used_range->dynamicCall("Value");
    QVariantList kf_used_rows = kf_used_data.value<QVariantList>();

    QVariantList kf_data_params;
    QString kf_data_end("L");
    QString kf_data_end_number = QString::number(user_row_count);
    kf_data_params << "Q2" << QVariant::fromValue(kf_data_end + kf_data_end_number);// 重点跟进 L-Q账单环节 发票环节 账单自动订阅 发票订阅 自动开票日期 开通日期

    QAxObject *kf_data_used_range = user_usedRange->querySubObject("Range(QVariant, QVariant)", kf_data_params);
    QVariant kf_data_used_data = kf_data_used_range->dynamicCall("Value");
    QVariantList kf_data_used_rows = kf_data_used_data.value<QVariantList>();

    int i;
    QMap<QString, Key_follow> kf_map;
    QString net_code_history = "";
    QMap<QString, Point_part> pp_map;
    QMap<QString, Point_part> bp_map;
    QMap<QString, Finance_receive> fr_map;
    QMap<QString, Finance_receive> re_map;
    for(i=0; i < kf_used_rows.size(); ++i)
    {
        QVariantList kf_col = kf_used_rows[i].value<QVariantList>();
        QString type = kf_col[0].toString(); // 分类
        QString username = kf_col[1].toString(); // 重点跟进人员名
        QString finance = kf_col[2].toString(); // 财务人员
        QString receive = kf_col[3].toString(); // 应收人员
        QString branch_part = kf_col[4].toString(); // 分部
        QString net_code = kf_col[5].toString(); // 网点代码
        if(!username.isEmpty()){
            // 找到一样的，更新end
            QMap<QString, Key_follow>::iterator it = kf_map.find(username);
            if(it != kf_map.end()){
                Key_follow fk_t = kf_map[username];
                int end = fk_t.end.last();
                if(end + 1 == i){// 挨着的，不用放到另一段中
                    fk_t.end.replace(fk_t.end.size() - 1, i);
                }else{// 不挨着，新段start，end
                    fk_t.start.append(i);
                    fk_t.end.append(i);
                }
                kf_map.insert(username, fk_t);
            } else {// 没找到一样的，创建key_follow start和end
                QList<int> start;
                start.append(i);
                QList<int> end;
                end.append(i);
                Key_follow kf(start, end, username);
                kf.type = type;
                kf.branch_part  = branch_part;
                kf_map.insert(username, kf);
            }
        }
        if(!net_code.isEmpty()){
            // 网点代码
            if(QString::compare(net_code_history, net_code) != 0){
                net_code_history = net_code;
                Point_part pp(net_code, branch_part, finance, receive);
                QList<int> start;
                start.append(i);
                QList<int> end;
                end.append(i);
                pp.start = start;
                pp.end = end;
                pp_map.insert(net_code, pp);
            }else{
                Point_part p = pp_map[net_code];
                p.end.replace(p.end.size() - 1, i);
                pp_map.insert(net_code, p);
            }
        }
        if(!branch_part.isEmpty()){
            // 找到一样的，更新end
            QMap<QString, Point_part>::iterator it = bp_map.find(branch_part);
            if(it != bp_map.end()){
                Point_part bp_t = bp_map[branch_part];
                int end = bp_t.end.last();
                if(end + 1 == i){// 挨着的，不用放到另一段中
                    bp_t.end.replace(bp_t.end.size() - 1, i);
                }else{// 不挨着，新段start，end
                    bp_t.start.append(i);
                    bp_t.end.append(i);
                }
                bp_map.insert(branch_part, bp_t);
            } else {// 没找到一样的，创建Point_part start和end
                QList<int> start;
                start.append(i);
                QList<int> end;
                end.append(i);
                Point_part bp;
                bp.branch_part  = branch_part;
                bp.finance_person = finance;
                bp.receive_person = receive;
                bp.end = end;
                bp.start = start;
                bp_map.insert(branch_part, bp);
            }
        }

        if(!finance.isEmpty()){
            // 找到一样的，更新end
            QMap<QString, Finance_receive>::iterator it = fr_map.find(finance);
            if(it != fr_map.end()){
                Finance_receive fr_t = fr_map[finance];
                int end = fr_t.end.last();
                if(end + 1 == i){// 挨着的，不用放到另一段中
                    fr_t.end.replace(fr_t.end.size() - 1, i);
                }else{// 不挨着，新段start，end
                    fr_t.start.append(i);
                    fr_t.end.append(i);
                }
                fr_map.insert(finance, fr_t);
            } else {// 没找到一样的，创建Finance_receive start和end
                QList<int> start;
                start.append(i);
                QList<int> end;
                end.append(i);
                Finance_receive fr(start, end, finance);
                fr_map.insert(finance, fr);
            }
        }

        if(!receive.isEmpty()){
            // 找到一样的，更新end
            QMap<QString, Finance_receive>::iterator it = re_map.find(receive);
            if(it != re_map.end()){
                Finance_receive re_t = re_map[receive];
                int end = re_t.end.last();
                if(end + 1 == i){// 挨着的，不用放到另一段中
                    re_t.end.replace(re_t.end.size() - 1, i);
                }else{// 不挨着，新段start，end
                    re_t.start.append(i);
                    re_t.end.append(i);
                }
                re_map.insert(receive, re_t);
            } else {// 没找到一样的，创建Finance_receive start和end
                QList<int> start;
                start.append(i);
                QList<int> end;
                end.append(i);
                Finance_receive re(start, end, receive);
                re_map.insert(receive, re);
            }
        }

    }


    QList<QList<QVariant>> kf_segments;
    QList<QList<QVariant>> kf_br_segments;
    QMap<QString, Key_follow>::const_iterator kf_i;
    for(kf_i = kf_map.begin(); kf_i != kf_map.end(); ++kf_i){
        QList<QVariant> kfList;
        QList<QVariant> kf_brList;
        Key_follow element = kf_i.value();
        int bs_no = 0;
        int bs_yes = 0;
        int rs_no = 0;
        int rs_yes = 0;
        int su_no = 0;
        int su_yes = 0;
        int re_no = 0;
        int re_yes = 0;
        for(int k = 0; k < element.start.size(); ++ k){
            int start = element.start[k];
            int end = element.end[k];
            while(start <= end){
                QString L = kf_data_used_rows[start].value<QVariantList>()[0].toString();// L列账单环节
                if(QString::compare(L, NO_USE) == 0){
                    ++ bs_no;
                }else if(QString::compare(L, HAVE_USE) == 0){
                    ++ bs_yes;
                }
                QString M = kf_data_used_rows[start].value<QVariantList>()[1].toString();// M列发票环节
                if(QString::compare(M, NO_USE) == 0){
                    ++ rs_no;
                }else if(QString::compare(M, HAVE_USE) == 0){
                    ++ rs_yes;
                }
                QString N = kf_data_used_rows[start].value<QVariantList>()[2].toString();// N列账单自动订阅
                if(QString::compare(N, NO_USE) == 0){
                    ++ su_no;
                }else if(QString::compare(N, HAVE_USE) == 0){
                    ++ su_yes;
                }
                QString O = kf_data_used_rows[start].value<QVariantList>()[3].toString();// O列发票订阅
                if(QString::compare(O, NO) == 0){
                    ++ re_no;
                }else if(QString::compare(O, YES) == 0){
                    ++ re_yes;
                }
                ++ start;
            }
        }

        kfList.append(element.username);
        kfList.append(element.branch_part);
        kfList.append(element.type);
        kfList.append(su_no);
        kfList.append(su_yes);
        int su_sum = su_no + su_yes;
        kfList.append(su_sum);
        kfList.append("");
        if(su_sum == 0){
            kfList.append(0);
        }else{
            kfList.append((float) su_yes / su_sum);
        }
        kfList.append(re_no);
        kfList.append(re_yes);
        int re_sum = re_no + re_yes;
        kfList.append(re_sum);
        kfList.append("");
        if(re_sum == 0){
            kfList.append(0);
        }else{
            kfList.append((float) re_yes / re_sum);
        }
        kf_segments.append(kfList);

        kf_brList.append(element.username);
        kf_brList.append(element.branch_part);
        kf_brList.append(element.type);
        kf_brList.append(bs_no);
        kf_brList.append(bs_yes);
        int bs_sum = bs_no + bs_yes;
        kf_brList.append(bs_sum);
        if(bs_sum == 0){
            kf_brList.append(0);
        }else{
            kf_brList.append((float) bs_yes / bs_sum);
        }
        kf_brList.append(rs_no);
        kf_brList.append(rs_yes);
        int rs_sum = rs_no + rs_yes;
        kf_brList.append(rs_sum);
        if(rs_sum == 0){
            kf_brList.append(0);
        }else{
            kf_brList.append((float) rs_yes / rs_sum);
        }
        kf_br_segments.append(kf_brList);
    }

    QList<QList<QVariant>> nc_segments; // net code
    QMap<QString, Point_part>::const_iterator nc_i;
    for(nc_i = pp_map.begin(); nc_i != pp_map.end(); ++nc_i){
        QList<QVariant> ncList;
        Point_part element = nc_i.value();
        int bs_no = 0;
        int bs_yes = 0;
        int rs_no = 0;
        int rs_yes = 0;
        for(int k = 0; k < element.start.size(); ++ k){
            int start = element.start[k];
            int end = element.end[k];
            while(start <= end){
                QString L = kf_data_used_rows[start].value<QVariantList>()[0].toString();// L列账单环节
                if(QString::compare(L, NO_USE) == 0){
                    ++ bs_no;
                }else if(QString::compare(L, HAVE_USE) == 0){
                    ++ bs_yes;
                }
                QString M = kf_data_used_rows[start].value<QVariantList>()[1].toString();// M列发票环节
                if(QString::compare(M, NO_USE) == 0){
                    ++ rs_no;
                }else if(QString::compare(M, HAVE_USE) == 0){
                    ++ rs_yes;
                }
                ++ start;
            }
        }

        ncList.append(element.net_code);
        ncList.append(element.branch_part);
        ncList.append(element.finance_person);
        ncList.append(element.receive_person);
        ncList.append(bs_no);
        ncList.append(bs_yes);
        int bs_sum = bs_no + bs_yes;
        ncList.append(bs_sum);
        if(bs_sum == 0){
            ncList.append(0);
        }else{
            ncList.append((float) bs_yes / bs_sum);
        }
        ncList.append(rs_no);
        ncList.append(rs_yes);
        int rs_sum = rs_no + rs_yes;
        ncList.append(rs_sum);
        if(rs_sum == 0){
            ncList.append(0);
        }else{
            ncList.append((float) rs_yes / rs_sum);
        }
        nc_segments.append(ncList);
    }

    QList<QList<QVariant>> bp_segments; // net code
    QMap<QString, Point_part>::const_iterator bp_i;
    for(bp_i = bp_map.begin(); bp_i != bp_map.end(); ++bp_i){
        QList<QVariant> bpList;
        Point_part element = bp_i.value();
        int bs_no = 0;
        int bs_yes = 0;
        int rs_no = 0;
        int rs_yes = 0;
        for(int k = 0; k < element.start.size(); ++ k){
            int start = element.start[k];
            int end = element.end[k];
            while(start <= end){
                QString L = kf_data_used_rows[start].value<QVariantList>()[0].toString();// L列账单环节
                if(QString::compare(L, NO_USE) == 0){
                    ++ bs_no;
                }else if(QString::compare(L, HAVE_USE) == 0){
                    ++ bs_yes;
                }
                QString M = kf_data_used_rows[start].value<QVariantList>()[1].toString();// M列发票环节
                if(QString::compare(M, NO_USE) == 0){
                    ++ rs_no;
                }else if(QString::compare(M, HAVE_USE) == 0){
                    ++ rs_yes;
                }
                ++ start;
            }
        }

        bpList.append(element.branch_part);
        bpList.append(element.finance_person);
        bpList.append(element.receive_person);
        bpList.append(bs_no);
        bpList.append(bs_yes);
        int bs_sum = bs_no + bs_yes;
        bpList.append(bs_sum);
        if(bs_sum == 0){
            bpList.append(0);
        }else{
            bpList.append((float) bs_yes / bs_sum);
        }
        bpList.append(rs_no);
        bpList.append(rs_yes);
        int rs_sum = rs_no + rs_yes;
        bpList.append(rs_sum);
        if(rs_sum == 0){
            bpList.append(0);
        }else{
            bpList.append((float) rs_yes / rs_sum);
        }
        bp_segments.append(bpList);
    }

    QList<QList<QVariant>> fr_segments; // finace
    QMap<QString, Finance_receive>::const_iterator fr_i;
    for(fr_i = fr_map.begin(); fr_i != fr_map.end(); ++fr_i){
        QList<QVariant> frList;
        Finance_receive element = fr_i.value();
        int bs_no = 0;
        int bs_yes = 0;
        int rs_no = 0;
        int rs_yes = 0;
        for(int k = 0; k < element.start.size(); ++ k){
            int start = element.start[k];
            int end = element.end[k];
            while(start <= end){
                QString L = kf_data_used_rows[start].value<QVariantList>()[0].toString();// L列账单环节
                if(QString::compare(L, NO_USE) == 0){
                    ++ bs_no;
                }else if(QString::compare(L, HAVE_USE) == 0){
                    ++ bs_yes;
                }
                QString M = kf_data_used_rows[start].value<QVariantList>()[1].toString();// M列发票环节
                if(QString::compare(M, NO_USE) == 0){
                    ++ rs_no;
                }else if(QString::compare(M, HAVE_USE) == 0){
                    ++ rs_yes;
                }
                ++ start;
            }
        }

        frList.append(element.fr);
        frList.append(bs_no);
        frList.append(bs_yes);
        int bs_sum = bs_no + bs_yes;
        frList.append(bs_sum);
        if(bs_sum == 0){
            frList.append(0);
        }else{
            frList.append((float) bs_yes / bs_sum);
        }
        frList.append(0.8); // 目标值
        frList.append(0.8 - frList.at(4).toFloat());
        frList.append(rs_no);
        frList.append(rs_yes);
        int rs_sum = rs_no + rs_yes;
        frList.append(rs_sum);
        if(rs_sum == 0){
            frList.append(0);
        }else{
            frList.append((float) rs_yes / rs_sum);
        }
        frList.append(0.9); // 目标值
        frList.append(0.9 - frList.at(10).toFloat());
        fr_segments.append(frList);
    }

    QList<QList<QVariant>> re_segments; // receive
    QMap<QString, Finance_receive>::const_iterator re_i;
    int bs_no_sum = 0;
    int bs_yes_sum = 0;
    int bs_sum_sum = 0;
    float bs_percent_sum = 0;
    int rs_no_sum = 0;
    int rs_yes_sum = 0;
    int rs_sum_sum = 0;
    float rs_percent_sum = 0;
    int count = 0;
    for(re_i = re_map.begin(); re_i != re_map.end(); ++re_i){
        ++ count;
        QList<QVariant> reList;
        Finance_receive element = re_i.value();
        int bs_no = 0;
        int bs_yes = 0;
        int rs_no = 0;
        int rs_yes = 0;
        for(int k = 0; k < element.start.size(); ++ k){
            int start = element.start[k];
            int end = element.end[k];
            while(start <= end){
                QString L = kf_data_used_rows[start].value<QVariantList>()[0].toString();// L列账单环节
                if(QString::compare(L, NO_USE) == 0){
                    ++ bs_no;
                }else if(QString::compare(L, HAVE_USE) == 0){
                    ++ bs_yes;
                }
                QString M = kf_data_used_rows[start].value<QVariantList>()[1].toString();// M列发票环节
                if(QString::compare(M, NO_USE) == 0){
                    ++ rs_no;
                }else if(QString::compare(M, HAVE_USE) == 0){
                    ++ rs_yes;
                }
                ++ start;
            }
        }

        reList.append(element.fr);
        reList.append(bs_no);
        reList.append(bs_yes);
        int bs_sum = bs_no + bs_yes;
        reList.append(bs_sum);
        if(bs_sum == 0){
            reList.append(0);
        }else{
            reList.append((float) bs_yes / bs_sum);
        }
        reList.append(0.8); // 目标值
        reList.append(0.8 - reList.at(4).toFloat());
        reList.append(rs_no);
        reList.append(rs_yes);
        int rs_sum = rs_no + rs_yes;
        reList.append(rs_sum);
        if(rs_sum == 0){
            reList.append(0);
        }else{
            reList.append((float) rs_yes / rs_sum);
        }
        reList.append(0.9); // 目标值
        reList.append(0.9 - reList.at(10).toFloat());
        re_segments.append(reList);

        bs_no_sum += bs_no;
        bs_yes_sum += bs_yes;
        bs_sum_sum += bs_sum;
        bs_percent_sum += reList.at(4).toFloat();

        rs_no_sum += rs_no;
        rs_yes_sum += rs_yes;
        rs_sum_sum += rs_sum;
        rs_percent_sum += reList.at(10).toFloat();

    }
    QList<QVariant> reLastList;
    reLastList.append("合计");
    reLastList.append(bs_no_sum);
    reLastList.append(bs_yes_sum);
    reLastList.append(bs_sum_sum);
    reLastList.append(bs_percent_sum / count);
    reLastList.append("");
    reLastList.append("");
    reLastList.append(rs_no_sum);
    reLastList.append(rs_yes_sum);
    reLastList.append(rs_sum_sum);
    reLastList.append(rs_percent_sum / count);
    reLastList.append("");
    reLastList.append("");
    re_segments.append(reLastList);

    QVariantList kf_segment_params;
    QString kf_segment_end("M");
    QString kf_segment_end_number = QString::number(kf_segments.size() + 2);
    kf_segment_params << "A3" << QVariant::fromValue(kf_segment_end + kf_segment_end_number);// 重点跟进 分部 分类 账单订阅是否总计订阅率 发票订阅是否订阅率

    QVariant kf_var;
    castListListVariant2Variant(kf_var, kf_segments);

    QAxObject *kf_segment_range = kf_usedRange->querySubObject("Range(QVariant, QVariant)", kf_segment_params);
    kf_segment_range->setProperty("Value", kf_var);

    QVariantList kf_br_segment_params;
    QString kf_br_segment_end("K");
    QString kf_br_segment_end_number = QString::number(kf_br_segments.size() + 2);
    kf_br_segment_params << "A3" << QVariant::fromValue(kf_br_segment_end + kf_br_segment_end_number);// 重点跟进 分类 分部 账单环节未使用已使用合计使用率 发票环节未使用已使用合计使用率

    QVariant kf_br_var;
    castListListVariant2Variant(kf_br_var, kf_br_segments);

    QAxObject *kf_br_segment_range = kf_br_usedRange->querySubObject("Range(QVariant, QVariant)", kf_br_segment_params);
    kf_br_segment_range->setProperty("Value", kf_br_var);

    QVariantList nc_segment_params;
    QString nc_segment_end("L");
    QString nc_segment_end_number = QString::number(nc_segments.size() + 2);
    nc_segment_params << "A3" << QVariant::fromValue(nc_segment_end + nc_segment_end_number);// 点部网点代码 网点代码 分部 财务中心 应收 账单环节未使用已使用合计使用率 发票环节未使用已使用合计使用率

    QVariant nc_var;
    castListListVariant2Variant(nc_var, nc_segments);

    QAxObject *nc_segment_range = nc_usedRange->querySubObject("Range(QVariant, QVariant)", nc_segment_params);
    nc_segment_range->setProperty("Value", nc_var);

    QVariantList bp_segment_params;
    QString bp_segment_end("K");
    QString bp_segment_end_number = QString::number(bp_segments.size() + 2);
    bp_segment_params << "A3" << QVariant::fromValue(bp_segment_end + bp_segment_end_number);// 分部 分部 财务中心 应收 账单环节未使用已使用合计使用率 发票环节未使用已使用合计使用率

    QVariant bp_var;
    castListListVariant2Variant(bp_var, bp_segments);

    QAxObject *bp_segment_range = bp_usedRange->querySubObject("Range(QVariant, QVariant)", bp_segment_params);
    bp_segment_range->setProperty("Value", bp_var);

    QVariantList fr_segment_params;
    QString fr_segment_end("M");
    QString fr_segment_end_number = QString::number(fr_segments.size() + 2);
    fr_segment_params << "A3" << QVariant::fromValue(fr_segment_end + fr_segment_end_number);// 账务中心 财务中心 账单环节未使用已使用合计使用率目标差距 发票环节未使用已使用合计使用率目标差距

    QVariant fr_var;
    castListListVariant2Variant(fr_var, fr_segments);

    QAxObject *fr_segment_range = fr_usedRange->querySubObject("Range(QVariant, QVariant)", fr_segment_params);
    fr_segment_range->setProperty("Value",fr_var);

    QVariantList re_segment_params;
    QString re_segment_end("M");
    QString re_segment_end_number = QString::number(re_segments.size() + 2);
    re_segment_params << "A3" << QVariant::fromValue(re_segment_end + re_segment_end_number);// 账务中心 财务中心 账单环节未使用已使用合计使用率目标差距 发票环节未使用已使用合计使用率目标差距

    QVariant re_var;
    castListListVariant2Variant(re_var, re_segments);

    QAxObject *re_segment_range = re_usedRange->querySubObject("Range(QVariant, QVariant)", re_segment_params);
    re_segment_range->setProperty("Value", re_var);

//    QList<User_info> userList;
//    userList.reserve(user_row_count);
//    int user_key_follow_row = 4;
//    int user_key_follow_column = 2;
//    while(true){
//        QString month_over_code = user_info_sheet->querySubObject("Cells(int,int)", user_key_follow_row, month_over_code_column)->dynamicCall("Value2()").toString();
//        QString username = user_info_sheet->querySubObject("Cells(int,int)", user_key_follow_row, user_key_follow_column)->dynamicCall("Value2()").toString();
//        if(month_over_code.isEmpty()) break;
//        if(username.isEmpty()) {
//            user_key_follow_row++;
//            continue;
//        }
//        User_info user(user_key_follow_row - 4, 0, username);
//        userList.append(user);
//        user_key_follow_row++;
//    }

//    QList<User_info> temp(userList);
//    temp.reserve(userList.size() - 1);
//    util.mergeSortStruct(userList, 0, userList.size() - 1, temp);

//    int type_column = 1;
//    int key_follow_column = 2;
//    int branch_part_column = 5;
//    int receive_person_column = 4;
//    int bill_auto_sub_column = 14;
//    int receipt_sub_column = 15;

//    //pp
//    int pp_code_row = 4;
//    int pp_net_code_column = 1;
//    int pp_branch_part_column = 2;
//    int pp_finance_person_column = 3;
//    int pp_receive_person_column = 4;
//    int pp_bill_no_use_column = 5;
//    int pp_bill_use_column = 6;
//    int pp_bill_sum_column = 7;
//    int pp_bill_percent_column = 8;
//    int pp_receipt_no_use_column = 9;
//    int pp_receipt_use_column = 10;
//    int pp_receipt_sum_column = 11;
//    int pp_receipt_percent_column = 12;

//    QString net_code_histroy = "";
//    Point_part part(NULL, NULL, NULL, NULL);
//    while (true) {
//        emit send_excel_row_done();
//        QString net_code_str = user_info_sheet->querySubObject("Cells(int,int)", code_row, net_code_column)->dynamicCall("Value2()").toString();
//        QString branch_part_str = user_info_sheet->querySubObject("Cells(int,int)", code_row, branch_part_column)->dynamicCall("Value2()").toString();
//        QString receive_person_str = user_info_sheet->querySubObject("Cells(int,int)", code_row, receive_person_column)->dynamicCall("Value2()").toString();
//        QString finance_person_str = user_info_sheet->querySubObject("Cells(int,int)", code_row, finance_person_column)->dynamicCall("Value2()").toString();
//        QString bill_segment_str = user_info_sheet->querySubObject("Cells(int,int)", code_row, bill_segment_column)->dynamicCall("Value2()").toString();
//        QString receipt_segment_str = user_info_sheet->querySubObject("Cells(int,int)", code_row, receipt_segment_column)->dynamicCall("Value2()").toString();
//        if(net_code_histroy != net_code_str){ // 新网点代码
//            net_code_histroy = net_code_str;
//            if(part.net_code != NULL){// one handle is over, ready to set values to "point_part_sheet"
//                QAxObject *pp_net_code_cell = point_part_sheet->querySubObject("Cells(int,int)", pp_code_row, pp_net_code_column);
//                pp_net_code_cell->setProperty("Value", part.net_code);
//                QAxObject *pp_branch_part_cell = point_part_sheet->querySubObject("Cells(int,int)", pp_code_row, pp_branch_part_column);
//                pp_branch_part_cell->setProperty("Value", part.branch_part);
//                QAxObject *pp_finance_person_cell = point_part_sheet->querySubObject("Cells(int,int)", pp_code_row, pp_finance_person_column);
//                pp_finance_person_cell->setProperty("Value", part.finance_person);
//                QAxObject *pp_receive_person_cell = point_part_sheet->querySubObject("Cells(int,int)", pp_code_row, pp_receive_person_column);
//                pp_receive_person_cell->setProperty("Value", part.receive_person);
//                QAxObject *pp_bill_no_use_cell = point_part_sheet->querySubObject("Cells(int,int)", pp_code_row, pp_bill_no_use_column);
//                pp_bill_no_use_cell->setProperty("Value", part.bill_no_use_count);
//                QAxObject *pp_bill_use_cell = point_part_sheet->querySubObject("Cells(int,int)", pp_code_row, pp_bill_use_column);
//                pp_bill_use_cell->setProperty("Value", part.bill_use_count);
//                QAxObject *pp_bill_sum_cell = point_part_sheet->querySubObject("Cells(int,int)", pp_code_row, pp_bill_sum_column);
//                part.bill_sum = part.bill_no_use_count + part.bill_use_count;
//                pp_bill_sum_cell->setProperty("Value", part.bill_sum);
//                QAxObject *pp_bill_percent_cell = point_part_sheet->querySubObject("Cells(int,int)", pp_code_row, pp_bill_percent_column);
//                part.bill_percent = (float)part.bill_use_count / part.bill_sum;
//                pp_bill_percent_cell->setProperty("Value", part.bill_percent);
//                QAxObject *pp_receipt_no_use_cell = point_part_sheet->querySubObject("Cells(int,int)", pp_code_row, pp_receipt_no_use_column);
//                pp_receipt_no_use_cell->setProperty("Value", part.receipt_no_use_count);
//                QAxObject *pp_receipt_use_cell = point_part_sheet->querySubObject("Cells(int,int)", pp_code_row, pp_receipt_use_column);
//                pp_receipt_use_cell->setProperty("Value", part.receipt_use_count);
//                part.receipt_sum = part.receipt_no_use_count + part.receipt_use_count;
//                QAxObject *pp_receipt_sum_cell = point_part_sheet->querySubObject("Cells(int,int)", pp_code_row, pp_receipt_sum_column);
//                pp_receipt_sum_cell->setProperty("Value", part.receipt_sum);
//                part.receipt_percent = (float)part.receipt_use_count / part.receipt_sum;
//                QAxObject *pp_receipt_percent_cell = point_part_sheet->querySubObject("Cells(int,int)", pp_code_row, pp_receipt_percent_column);
//                pp_receipt_percent_cell->setProperty("Value", part.receipt_percent);
//                clearPP(part);
//                if(net_code_str.isEmpty()){
//                    break;
//                }
//                ++ pp_code_row;
//            }
//            part.net_code = net_code_str;
//            part.branch_part = branch_part_str;
//            part.receive_person = receive_person_str;
//            part.finance_person = finance_person_str;
//        }
//        if(bill_segment_str == NO_USE){
//            ++ part.bill_no_use_count;
//        }else if(bill_segment_str == HAVE_USE){
//            ++ part.bill_use_count;
//        }
//        if(receipt_segment_str == NO_USE){
//            ++ part.receipt_no_use_count;
//        }else if(receipt_segment_str == HAVE_USE){
//            ++ part.receipt_use_count;
//        }
//        ++ code_row;
//    }

//    // 点部（网点代码）
//    QAxObject *point_part_sheet = work_sheets1->querySubObject("Item(int)", 4);
//    int net_code_column = 6;
//    int branch_part_column = 5;
//    int receive_person_column = 4;
//    int finance_person_column = 3;
//    int bill_segment_column = 12;
//    int receipt_segment_column = 13;

//    //pp
//    int pp_code_row = 4;
//    int pp_net_code_column = 1;
//    int pp_branch_part_column = 2;
//    int pp_finance_person_column = 3;
//    int pp_receive_person_column = 4;
//    int pp_bill_no_use_column = 5;
//    int pp_bill_use_column = 6;
//    int pp_bill_sum_column = 7;
//    int pp_bill_percent_column = 8;
//    int pp_receipt_no_use_column = 9;
//    int pp_receipt_use_column = 10;
//    int pp_receipt_sum_column = 11;
//    int pp_receipt_percent_column = 12;

//    QString net_code_histroy = "";
//    Point_part part(NULL, NULL, NULL, NULL);
//    while (true) {
//        emit send_excel_row_done();
//        QString net_code_str = user_info_sheet->querySubObject("Cells(int,int)", code_row, net_code_column)->dynamicCall("Value2()").toString();
//        QString branch_part_str = user_info_sheet->querySubObject("Cells(int,int)", code_row, branch_part_column)->dynamicCall("Value2()").toString();
//        QString receive_person_str = user_info_sheet->querySubObject("Cells(int,int)", code_row, receive_person_column)->dynamicCall("Value2()").toString();
//        QString finance_person_str = user_info_sheet->querySubObject("Cells(int,int)", code_row, finance_person_column)->dynamicCall("Value2()").toString();
//        QString bill_segment_str = user_info_sheet->querySubObject("Cells(int,int)", code_row, bill_segment_column)->dynamicCall("Value2()").toString();
//        QString receipt_segment_str = user_info_sheet->querySubObject("Cells(int,int)", code_row, receipt_segment_column)->dynamicCall("Value2()").toString();
//        if(net_code_histroy != net_code_str){ // 新网点代码
//            net_code_histroy = net_code_str;
//            if(part.net_code != NULL){// one handle is over, ready to set values to "point_part_sheet"
//                QAxObject *pp_net_code_cell = point_part_sheet->querySubObject("Cells(int,int)", pp_code_row, pp_net_code_column);
//                pp_net_code_cell->setProperty("Value", part.net_code);
//                QAxObject *pp_branch_part_cell = point_part_sheet->querySubObject("Cells(int,int)", pp_code_row, pp_branch_part_column);
//                pp_branch_part_cell->setProperty("Value", part.branch_part);
//                QAxObject *pp_finance_person_cell = point_part_sheet->querySubObject("Cells(int,int)", pp_code_row, pp_finance_person_column);
//                pp_finance_person_cell->setProperty("Value", part.finance_person);
//                QAxObject *pp_receive_person_cell = point_part_sheet->querySubObject("Cells(int,int)", pp_code_row, pp_receive_person_column);
//                pp_receive_person_cell->setProperty("Value", part.receive_person);
//                QAxObject *pp_bill_no_use_cell = point_part_sheet->querySubObject("Cells(int,int)", pp_code_row, pp_bill_no_use_column);
//                pp_bill_no_use_cell->setProperty("Value", part.bill_no_use_count);
//                QAxObject *pp_bill_use_cell = point_part_sheet->querySubObject("Cells(int,int)", pp_code_row, pp_bill_use_column);
//                pp_bill_use_cell->setProperty("Value", part.bill_use_count);
//                QAxObject *pp_bill_sum_cell = point_part_sheet->querySubObject("Cells(int,int)", pp_code_row, pp_bill_sum_column);
//                part.bill_sum = part.bill_no_use_count + part.bill_use_count;
//                pp_bill_sum_cell->setProperty("Value", part.bill_sum);
//                QAxObject *pp_bill_percent_cell = point_part_sheet->querySubObject("Cells(int,int)", pp_code_row, pp_bill_percent_column);
//                part.bill_percent = (float)part.bill_use_count / part.bill_sum;
//                pp_bill_percent_cell->setProperty("Value", part.bill_percent);
//                QAxObject *pp_receipt_no_use_cell = point_part_sheet->querySubObject("Cells(int,int)", pp_code_row, pp_receipt_no_use_column);
//                pp_receipt_no_use_cell->setProperty("Value", part.receipt_no_use_count);
//                QAxObject *pp_receipt_use_cell = point_part_sheet->querySubObject("Cells(int,int)", pp_code_row, pp_receipt_use_column);
//                pp_receipt_use_cell->setProperty("Value", part.receipt_use_count);
//                part.receipt_sum = part.receipt_no_use_count + part.receipt_use_count;
//                QAxObject *pp_receipt_sum_cell = point_part_sheet->querySubObject("Cells(int,int)", pp_code_row, pp_receipt_sum_column);
//                pp_receipt_sum_cell->setProperty("Value", part.receipt_sum);
//                part.receipt_percent = (float)part.receipt_use_count / part.receipt_sum;
//                QAxObject *pp_receipt_percent_cell = point_part_sheet->querySubObject("Cells(int,int)", pp_code_row, pp_receipt_percent_column);
//                pp_receipt_percent_cell->setProperty("Value", part.receipt_percent);
//                clearPP(part);
//                if(net_code_str.isEmpty()){
//                    break;
//                }
//                ++ pp_code_row;
//            }
//            part.net_code = net_code_str;
//            part.branch_part = branch_part_str;
//            part.receive_person = receive_person_str;
//            part.finance_person = finance_person_str;
//        }
//        if(bill_segment_str == NO_USE){
//            ++ part.bill_no_use_count;
//        }else if(bill_segment_str == HAVE_USE){
//            ++ part.bill_use_count;
//        }
//        if(receipt_segment_str == NO_USE){
//            ++ part.receipt_no_use_count;
//        }else if(receipt_segment_str == HAVE_USE){
//            ++ part.receipt_use_count;
//        }
//        ++ code_row;
//    }

//    // 分部
//    QAxObject *branch_part_sheet = work_sheets1->querySubObject("Item(int)", 3);

//    code_row = 4;

//    // part
//    int part_code_row = 4;
//    int part_branch_column = 1;
//    int part_finance_person_column = 2;
//    int part_receive_person_column = 3;
//    int part_bill_no_use_column = 4;
//    int part_bill_use_column = 5;
//    int part_bill_sum_column = 6;
//    int part_bill_percent_column = 7;
//    int part_receipt_no_use_column = 8;
//    int part_receipt_use_column = 9;
//    int part_receipt_sum_column = 10;
//    int part_receipt_percent_column = 11;

//    QString part_branch_history = "";
//    clearPP(part);
//    while (true) {
//        emit send_excel_row_done();
//        QString branch_part_str = user_info_sheet->querySubObject("Cells(int,int)", code_row, branch_part_column)->dynamicCall("Value2()").toString();
//        QString receive_person_str = user_info_sheet->querySubObject("Cells(int,int)", code_row, receive_person_column)->dynamicCall("Value2()").toString();
//        QString finance_person_str = user_info_sheet->querySubObject("Cells(int,int)", code_row, finance_person_column)->dynamicCall("Value2()").toString();
//        QString bill_segment_str = user_info_sheet->querySubObject("Cells(int,int)", code_row, bill_segment_column)->dynamicCall("Value2()").toString();
//        QString receipt_segment_str = user_info_sheet->querySubObject("Cells(int,int)", code_row, receipt_segment_column)->dynamicCall("Value2()").toString();
//        if(part_branch_history != branch_part_str){ // 新分部
//            part_branch_history = branch_part_str;
//            if(part.branch_part != NULL){// one handle is over, ready to set values to "point_part_sheet"
//                QAxObject *pp_branch_part_cell = branch_part_sheet->querySubObject("Cells(int,int)", part_code_row, part_branch_column);
//                pp_branch_part_cell->setProperty("Value", part.branch_part);
//                QAxObject *pp_finance_person_cell = branch_part_sheet->querySubObject("Cells(int,int)", part_code_row, part_finance_person_column);
//                pp_finance_person_cell->setProperty("Value", part.finance_person);
//                QAxObject *pp_receive_person_cell = branch_part_sheet->querySubObject("Cells(int,int)", part_code_row, part_receive_person_column);
//                pp_receive_person_cell->setProperty("Value", part.receive_person);
//                QAxObject *pp_bill_no_use_cell = branch_part_sheet->querySubObject("Cells(int,int)", part_code_row, part_bill_no_use_column);
//                pp_bill_no_use_cell->setProperty("Value", part.bill_no_use_count);
//                QAxObject *pp_bill_use_cell = branch_part_sheet->querySubObject("Cells(int,int)", part_code_row, part_bill_use_column);
//                pp_bill_use_cell->setProperty("Value", part.bill_use_count);
//                QAxObject *pp_bill_sum_cell = branch_part_sheet->querySubObject("Cells(int,int)", part_code_row, part_bill_sum_column);
//                part.bill_sum = part.bill_no_use_count + part.bill_use_count;
//                pp_bill_sum_cell->setProperty("Value", part.bill_sum);
//                QAxObject *pp_bill_percent_cell = branch_part_sheet->querySubObject("Cells(int,int)", part_code_row, part_bill_percent_column);
//                part.bill_percent = (float)part.bill_use_count / part.bill_sum;
//                pp_bill_percent_cell->setProperty("Value", part.bill_percent);
//                QAxObject *pp_receipt_no_use_cell = branch_part_sheet->querySubObject("Cells(int,int)", part_code_row, part_receipt_no_use_column);
//                pp_receipt_no_use_cell->setProperty("Value", part.receipt_no_use_count);
//                QAxObject *pp_receipt_use_cell = branch_part_sheet->querySubObject("Cells(int,int)", part_code_row, part_receipt_use_column);
//                pp_receipt_use_cell->setProperty("Value", part.receipt_use_count);
//                part.receipt_sum = part.receipt_no_use_count + part.receipt_use_count;
//                QAxObject *pp_receipt_sum_cell = branch_part_sheet->querySubObject("Cells(int,int)", part_code_row, part_receipt_sum_column);
//                pp_receipt_sum_cell->setProperty("Value", part.receipt_sum);
//                part.receipt_percent = (float)part.receipt_use_count / part.receipt_sum;
//                QAxObject *pp_receipt_percent_cell = branch_part_sheet->querySubObject("Cells(int,int)", part_code_row, part_receipt_percent_column);
//                pp_receipt_percent_cell->setProperty("Value", part.receipt_percent);
//                clearPP(part);
//                if(branch_part_str.isEmpty()){
//                    break;
//                }
//                ++ part_code_row;
//            }
//            part.branch_part = branch_part_str;
//            part.receive_person = receive_person_str;
//            part.finance_person = finance_person_str;
//        }
//        if(bill_segment_str == NO_USE){
//            ++ part.bill_no_use_count;
//        }else if(bill_segment_str == HAVE_USE){
//            ++ part.bill_use_count;
//        }
//        if(receipt_segment_str == NO_USE){
//            ++ part.receipt_no_use_count;
//        }else if(receipt_segment_str == HAVE_USE){
//            ++ part.receipt_use_count;
//        }
//        ++ code_row;
//    }

    work_book1->dynamicCall("Save()");  //保存文件（为了对比test与下面的test2文件，这里不做保存操作） work_book->dynamicCall("SaveAs(const QString&)", "E:\\test2.xlsx");  //另存为另一个文件
    work_book1->dynamicCall("Close(Boolean)", false);  //关闭文件
    excel1->dynamicCall("Quit(void)");  //退出
    delete excel1;

    emit send_export_over_signal(path1);
}

void WorkThread::close(){
    if(work_book1 != NULL){
        work_book1->dynamicCall("Save()");  //保存文件（为了对比test与下面的test2文件，这里不做保存操作） work_book->dynamicCall("SaveAs(const QString&)", "E:\\test2.xlsx");  //另存为另一个文件
        work_book1->dynamicCall("Close(Boolean)", false);  //关闭文件
    }
    if(excel1 != NULL){
        excel1->dynamicCall("Quit(void)");  //退出
        delete excel1;
    }

    if(work_book2 != NULL){
        work_book2->dynamicCall("Save()");
        work_book2->dynamicCall("Close(Boolean)", false);
    }
    if(excel2 != NULL){
        excel2->dynamicCall("Quit(void)");
        delete excel2;
    }

    if(work_book3 != NULL){
        work_book3->dynamicCall("Save()");
        work_book3->dynamicCall("Close(Boolean)", false);
    }
    if(excel3 != NULL){
        excel3->dynamicCall("Quit(void)");
        delete excel3;
    }

    if(work_book4 != NULL){
        work_book4->dynamicCall("Save()");
        work_book4->dynamicCall("Close(Boolean)", false);
    }
    if(excel4 != NULL){
        excel4->dynamicCall("Quit(void)");
        delete excel4;
    }

    if(work_book5 != NULL){
        work_book5->dynamicCall("Save()");
        work_book5->dynamicCall("Close(Boolean)", false);
    }
    if(excel5 != NULL){
        excel5->dynamicCall("Quit(void)");
        delete excel5;
    }
    if(work_book6 != NULL){
        work_book6->dynamicCall("Save()");
        work_book6->dynamicCall("Close(Boolean)", false);
    }
    if(excel6 != NULL){
        excel6->dynamicCall("Quit(void)");
        delete excel6;
    }
}

//void WorkThread::send()
//{
//    emit send_export_signal();
//}





