package com.cm.easyexcel.util;

import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Font;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.metadata.Table;
import com.alibaba.excel.metadata.TableStyle;
import com.cm.easyexcel.util.bean.Excel;
import com.cm.easyexcel.util.bean.ExcelTableStyle;
import com.cm.easyexcel.util.constant.ExcelContant;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.Assert;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

/**
 * @description: Excel表格生成工具类   可以生成后缀名为.csv .xls  .xlsx .pdf 的Excel表格数据文件
 * @author: caomian
 * @data: 2019/8/13 19:14
 */
public class ExcelTool {
    private static final Logger logger = LoggerFactory.getLogger(ExcelTool.class);

    /**
     * 导出
     *
     * @param excel Excel表格对象
     * @throws IOException
     */
    public static void Export(Excel excel) throws IOException {
        Assert.notNull(excel.getPath(), "输出位置不能为空");
        //文件输出位置
        OutputStream out = new FileOutputStream(excel.getPath() + "\\" + excel.getSheetName() + excel.getExcelSuffix());

        ExcelWriter writer = EasyExcelFactory.getWriter(out);
        // 动态添加表头
        Sheet sheet = new Sheet(excel.getSheetNo(), excel.getHeadLineMun());

        sheet.setSheetName(excel.getSheetName());

        // 创建表格
        Table table = new Table(excel.getTableNo());
        // 动态添加表头
        table.setHead(excel.getTableHeadList());
        // 自定义表格样式
        if (excel.getTableStyle() != null) {
            table.setTableStyle(excel.getTableStyle());
        }

        // 写数据
        writer.write1(excel.getTableBodyList(), sheet, table);
        // 上下文的最终outputStream写入文件
        writer.finish();
        out.close();
    }

    /**
     * Excel表格的风格
     *
     * @param excelTableStyle excel表格风格对象
     * @return
     */
    public static TableStyle createTableStyle(ExcelTableStyle excelTableStyle) {
        TableStyle tableStyle = new TableStyle();
        Font headFont = new Font();
        //是否加粗
        headFont.setBold(excelTableStyle.isHeadIsBold());
        // 字体大小
        headFont.setFontHeightInPoints(excelTableStyle.getHeadFontSize());
        headFont.setFontName(excelTableStyle.getHeadFontName());
        tableStyle.setTableHeadFont(headFont);
        tableStyle.setTableHeadBackGroundColor(excelTableStyle.getHeadColor());

        Font bodyFont = new Font();
        bodyFont.setBold(excelTableStyle.isBodyIsBold());
        bodyFont.setFontHeightInPoints(excelTableStyle.getBodyFontSize());
        bodyFont.setFontName(excelTableStyle.getBodyFontName());
        tableStyle.setTableContentFont(bodyFont);
        tableStyle.setTableContentBackGroundColor(excelTableStyle.getBodyColor());
        return tableStyle;
    }

    /**
     * 创建表头
     *
     * @param object 表头显示的不定参数
     * @return
     */
    public static List<List<String>> createTableHeadList(String... object) {
        List<List<String>> tableHeadList = new ArrayList<>();
        List<String> col = null;
        for (int i = ExcelContant.RESULT_ZERO; i < object.length; i++) {
            col = new ArrayList<>();
            col.add(object[i]);
            tableHeadList.add(col);
        }
        return tableHeadList;
    }

    /**
     * 获得excel表格的数据集合
     * 利用反射机制获取参数的大小和属性值
     *
     * @param data excel的数据
     * @return
     */
    public static List<List<Object>> listTableBodyData(List<?> data) {
        List<List<Object>> tableBodyList = new ArrayList<>();
        List<Object> row = null;
        Field[] fields = null;
        Field field = null;
        for (int i = ExcelContant.RESULT_ZERO; i < data.size(); i++) {
            row = new ArrayList<>();
            Object o = data.get(i);
            fields = o.getClass().getDeclaredFields();
            for (int j = ExcelContant.RESULT_ZERO; j < fields.length; j++) {
                try {
                    field = o.getClass().getDeclaredField(fields[j].getName());
                    field.setAccessible(true);
                    row.add(field.get(o));
                } catch (Exception e) {
                    logger.error("excel反射添加行数据出现异常：{}", e.getMessage());
                }
            }
            tableBodyList.add(row);
        }
        return tableBodyList;
    }
}
