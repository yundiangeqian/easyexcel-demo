package com.cm.easyexcel.util.constant;

import com.cm.easyexcel.util.bean.Excel;

/**
 * @description:
 * @author: caomian
 * @data: 2019/8/13 20:23
 */
public class ExcelContant {
    public static final int RESULT_ZERO = 0;

    public enum ExcelSuffix {
        XLSX(1, ".xlsx"), XLS(2, ".xls"), CSV(3, ".csv"), PDF(4, ".pdf");
        private int type;
        private String name;

        public int getType() {
            return type;
        }

        public String getName() {
            return name;
        }

        private ExcelSuffix(int type, String name) {
            this.type = type;
            this.name = name;
        }
    }
}
