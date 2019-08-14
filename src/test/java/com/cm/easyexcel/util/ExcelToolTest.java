package com.cm.easyexcel.util;

import com.alibaba.excel.metadata.BaseRowModel;
import com.cm.easyexcel.domain.entity.User;
import com.cm.easyexcel.util.bean.Excel;
import com.cm.easyexcel.util.bean.ExcelTableStyle;
import com.cm.easyexcel.util.constant.ExcelContant;
import org.apache.poi.ss.formula.functions.T;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import static org.junit.Assert.*;

@SpringBootTest
@RunWith(SpringRunner.class)
public class ExcelToolTest {
    private static final Logger logger = LoggerFactory.getLogger(ExcelToolTest.class);

    @Test
    public void export() {
        List<List<String>> tableHeadList = ExcelTool.createTableHeadList("姓名", "年龄", "性别");
        //模拟从数据库查询集合数据
        User user = null;
        List<User> userList = new ArrayList<>();
        for (int i = 0; i < 100; i++) {
            user = new User();
            user.setName("张三" + i);
            user.setAge(i);
            user.setGender(i % 2 == 0 ? "男" : "女");
            userList.add(user);
        }

        List<List<Object>> tableBodyList = ExcelTool.listTableBodyData(userList);
        Excel excel = new Excel();
        excel.setPath("C:\\Users\\caomian\\Desktop");
        excel.setTableHeadList(tableHeadList);
        excel.setTableBodyList(tableBodyList);
        excel.setSheetName("用户信息");
        excel.setExcelSuffix(ExcelContant.ExcelSuffix.CSV.getName());
        excel.setTableStyle(ExcelTool.createTableStyle(new ExcelTableStyle()));
        try {
            ExcelTool.Export(excel);
            logger.info("成功导出{}Excel表", excel.getSheetName());
        } catch (IOException e) {
            logger.error("excel导出出现IO异常", e.getMessage());
        }
    }
}