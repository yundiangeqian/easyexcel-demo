# easyexcel-demo
使用该工具类，需要先导入阿里的easyexcel包

导出很快，提高性能和效率，避免内存泄漏

数据量很大很大时，分页导出，你懂的

10行之内就可以成功导出数据

该工具类类进行了封装，调用非常简单
例如：
    @Test
    public void export() {
        //使用工具类封装好的创建表头方法，只需将表头以不定参数的方式传递即可
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
        //根据从数据库里查询的数据，使用工具类的获取表体数据
        List<List<Object>> tableBodyList = ExcelTool.listTableBodyData(userList);
        //设置Excel的内容  path、tableheadList、tableBodyList必填
        Excel excel = new Excel();
        excel.setPath("C:\\Users\\caomian\\Desktop");
        excel.setTableHeadList(tableHeadList);
        excel.setTableBodyList(tableBodyList);
        excel.setSheetName("用户信息");
        //常量类ExcelContant中的ExcelSuffix枚举定义了 当前的Excel的所有后缀名
        excel.setExcelSuffix(ExcelContant.ExcelSuffix.CSV.getName());
        //可以通过工具类的createTableStyle方法自定义设置样式
        excel.setTableStyle(ExcelTool.createTableStyle(new ExcelTableStyle()));
        try {
            //使用工具类导出
            ExcelTool.Export(excel);
            logger.info("成功导出{}Excel表", excel.getSheetName());
        } catch (IOException e) {
            logger.error("excel导出出现IO异常", e.getMessage());
        }
    }
