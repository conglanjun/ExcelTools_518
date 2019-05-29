#include "sortutils.h"

SortUtils::SortUtils()
{

}

void SortUtils::merge(QList<int> &resultList, int low, int mid, int high, int* tempList)
{
    int i = 0;
    int j = 0;
    int k = 0;
    for(i = low, j = mid + 1, k = i; i <= mid && j <= high; k++){
        if(resultList.at(i) <= resultList.at(j)){
            tempList[k] = resultList.at(i++);
        }else{
            tempList[k] = resultList.at(j++);
        }
    }
    while(i <= mid){
        tempList[k++] = resultList.at(i++);
    }
    while(j <= high){
        tempList[k++] = resultList.at(j++);
    }
}

void SortUtils::mergeSort(QList<int> &resultList, int low, int high, int* tempList)
{
    if(low < high){
        int mid = (low + high) / 2;
        mergeSort(resultList, low, mid, tempList);
        mergeSort(resultList, mid + 1, high, tempList);
        merge(resultList, low, mid, high, tempList);
        for(int i = low;i <= high;++i)
            resultList.replace(i, tempList[i]);
    }
}


void SortUtils::mergeStruct(QList<User_info> &resultList, int low, int mid, int high, QList<User_info> &tempList){
    int i = 0;
    int j = 0;
    int k = 0;
    for(i = low, j = mid + 1, k = i; i <= mid && j <= high; k++){
        if(QString::compare(resultList.at(i).username, resultList.at(j).username) < 0){
            tempList.replace(k, resultList.at(i++));
        }else{
            tempList.replace(k, resultList.at(j++));
        }
    }
    while(i <= mid){
        tempList.replace(k++, resultList.at(i++));
    }
    while(j <= high){
        tempList.replace(k++, resultList.at(j++));
    }
}

void SortUtils::mergeSortStruct(QList<User_info> &resultList, int low, int high, QList<User_info> &tempList){
    if(low < high){
        int mid = (low + high) / 2;
        mergeSortStruct(resultList, low, mid, tempList);
        mergeSortStruct(resultList, mid + 1, high, tempList);
        mergeStruct(resultList, low, mid, high, tempList);
        for(int i = low;i <= high;++i)
            resultList.replace(i, tempList.at(i));
    }
}

int SortUtils::binary_search(QList<User_info> userList, int key){
    int low = 0, high = userList.size() - 1, mid;
    while(low <= high){
        mid = (low + high) / 2;
        if(userList.at(mid).month_over_code == key){
            return mid;
        }else if(userList.at(mid).month_over_code > key){
            high = mid - 1;
        }else{
            low = mid + 1;
        }
    }
    return -1;
}

int SortUtils::binary_search2(QVariantList DM_rows, int column, int data_row, int count, qint64 key){
    int low = data_row, high = count, mid;
    while(low <= high){
        mid = (low + high) / 2;
        QVariantList tempList_col = DM_rows [mid].value<QVariantList>();
        qint64 cell_key = tempList_col[column].toLongLong();
        if(cell_key == key){
            return mid;
        }else if(cell_key > key){
            high = mid - 1;
        }else{
            low = mid + 1;
        }
    }
    return -1;
}






