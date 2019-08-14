package com.cm.easyexcel.util.bean;

import com.alibaba.excel.metadata.BaseRowModel;
import com.alibaba.excel.metadata.TableStyle;
import com.cm.easyexcel.util.constant.ExcelContant;
import org.apache.poi.ss.formula.functions.T;

import java.io.Serializable;
import java.util.List;

/**
 * @description:
 * @author: caomian
 * @data: 2019/8/13 20:07
 */
public class Excel implements Serializable {
    String path;
    int sheetNo = 1;
    int headLineMun = 0;
    String sheetName = String.valueOf(Math.ceil(Math.random() * 1000000));
    String excelSuffix = ExcelContant.ExcelSuffix.XLSX.getName();
    int tableNo = 1;
    List<List<String>> tableHeadList;
    List<List<Object>> tableBodyList;
    TableStyle tableStyle;

    public String getPath() {
        return path;
    }

    public void setPath(String path) {
        this.path = path;
    }

    public int getSheetNo() {
        return sheetNo;
    }

    public void setSheetNo(int sheetNo) {
        this.sheetNo = sheetNo;
    }

    public int getHeadLineMun() {
        return headLineMun;
    }

    public void setHeadLineMun(int headLineMun) {
        this.headLineMun = headLineMun;
    }

    public String getExcelSuffix() {
        return excelSuffix;
    }

    public void setExcelSuffix(String excelSuffix) {
        this.excelSuffix = excelSuffix;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public int getTableNo() {
        return tableNo;
    }

    public void setTableNo(int tableNo) {
        this.tableNo = tableNo;
    }

    public List<List<String>> getTableHeadList() {
        return tableHeadList;
    }

    public void setTableHeadList(List<List<String>> tableHeadList) {
        this.tableHeadList = tableHeadList;
    }

    public List<List<Object>> getTableBodyList() {
        return tableBodyList;
    }

    public void setTableBodyList(List<List<Object>> tableBodyList) {
        this.tableBodyList = tableBodyList;
    }

    public TableStyle getTableStyle() {
        return tableStyle;
    }

    public void setTableStyle(TableStyle tableStyle) {
        this.tableStyle = tableStyle;
    }
}
