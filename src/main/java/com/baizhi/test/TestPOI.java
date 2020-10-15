package com.baizhi.test;

import com.baizhi.entity.Emp;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.CellType;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.lang.reflect.Array;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class TestPOI {
    /**
     * 导出
     */
    @Test
    public void test0() throws Exception{
        // 1. 创建一个文档对象
        HSSFWorkbook workbook = new HSSFWorkbook();
        // 2. 创建工作簿
        HSSFSheet workbookSheet = workbook.createSheet("员工信息");
        // 3. 创建行
        HSSFRow headRow = workbookSheet.createRow(0);

        HSSFRow bodyRow = workbookSheet.createRow(1);
        // 4. 创建单元格
        /**
         * 创建表头 单元格 并设置内容
         */
        HSSFCell headRowCell0 = headRow.createCell(0);
        headRowCell0.setCellValue("编号");
        HSSFCell headRowCell1 = headRow.createCell(1);
        headRowCell1.setCellValue("名字");
        HSSFCell headRowCell2 = headRow.createCell(2);
        headRowCell2.setCellValue("工资");
        HSSFCell headRowCell3 = headRow.createCell(3);
        headRowCell3.setCellValue("年龄");

        /**
         * 创建表体单元格并设置内容
         */
        HSSFCell bodyRowCell0 = bodyRow.createCell(0);
        bodyRowCell0.setCellValue("1");
        HSSFCell bodyRowCell1 = bodyRow.createCell(1);
        bodyRowCell1.setCellValue("Tom");
        HSSFCell bodyRowCell2 = bodyRow.createCell(2);
        bodyRowCell2.setCellValue("1000");
        HSSFCell bodyRowCell3 = bodyRow.createCell(3);
        bodyRowCell3.setCellValue("18");

        // 5.导出excel
        workbook.write(new FileOutputStream(new File("d:/emp.xls")));
        workbook.close();

    }
    @Test
    public void test1() throws Exception{
        // 获取数据集合
        List<Emp> emps = getEmps();


        // 1. 创建一个文档对象
        HSSFWorkbook workbook = new HSSFWorkbook();
        // 2. 创建工作簿
        HSSFSheet workbookSheet = workbook.createSheet("员工信息");
        // 3. 创建行
        HSSFRow headRow = workbookSheet.createRow(0);
        /**
         * 创建样式对象
         */
        HSSFCellStyle cellStyle = workbook.createCellStyle();

        // 设置字体样式
        HSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setColor((short) 127 );
        font.setFontName("微软雅黑");
        cellStyle.setFont(font);

        // 设置日期格式 , 单独创建一个样式对象
        HSSFCellStyle cellStyleBody = workbook.createCellStyle();
        HSSFDataFormat dataFormat = workbook.createDataFormat();
        short format = dataFormat.getFormat("yyyy年MM月dd日");

        cellStyleBody.setDataFormat(format);


        // 4. 创建单元格
        /**
         * 创建表头 单元格 并设置内容
         */
        Class<Emp> empClass = Emp.class;
        Field[] declaredFields = empClass.getDeclaredFields();
        for (int i = 0; i < declaredFields.length; i++) {
            HSSFCell headRowCell0 = headRow.createCell(i);
            // 获取实体类的属性名，设置为表头的名字
            Field declaredField = declaredFields[i];
            declaredField.setAccessible(true);
            String name = declaredField.getName();
            headRowCell0.setCellValue(name);
            // 样式设置到单元格中
            headRowCell0.setCellStyle(cellStyle);
        }


        /**
         * 创建表体单元格并设置内容
         */
        for (int i = 1; i < emps.size(); i++) {
            HSSFRow bodyRow = workbookSheet.createRow(i);

            Emp emp = emps.get(i-1);
            HSSFCell bodyRowCell0 = bodyRow.createCell(0);
            bodyRowCell0.setCellValue(emp.getId());
            HSSFCell bodyRowCell1 = bodyRow.createCell(1);
            bodyRowCell1.setCellValue(emp.getName());
            HSSFCell bodyRowCell2 = bodyRow.createCell(2);
            bodyRowCell2.setCellValue(emp.getSalary());
            HSSFCell bodyRowCell3 = bodyRow.createCell(3);
            bodyRowCell3.setCellValue(emp.getAge());
            // 日期格式单元格演示
            HSSFCell bodyRowCell4 = bodyRow.createCell(4);
            bodyRowCell4.setCellStyle(cellStyleBody);
            bodyRowCell4.setCellValue(emp.getDate());
        }

        // 5.导出excel
        workbook.write(new FileOutputStream(new File("d:/emp.xls")));
        workbook.close();

    }

    /**
     * 导入
     *
     * @return
     */
    @Test
    public void test2() throws Exception{
        // 1. 创建一个文档
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(new File("d://emp001.xls")));

        // 2. 获取工作簿对象
        HSSFSheet sheet = workbook.getSheetAt(0);

        // 3. 获取内容
        /**
         * sheet.getLastRowNum()  获取当前工作簿最后一行下标
         */
        System.out.println(sheet.getLastRowNum());
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            HSSFRow row = sheet.getRow(i);

            // 获取当前行最后一列下标
            short lastCellNum = row.getLastCellNum();
            for (int j = 0; j < lastCellNum; j++) {
                HSSFCell cell = row.getCell(j);
                // 获取当前单元格的数据类型
                CellType type = cell.getCellType();
                if ("STRING".equals(type.name())) {
                    System.out.print(cell.getStringCellValue()+"\t");
                } else if ("NUMERIC".equals(type.name())) {
                    System.out.print(cell.getNumericCellValue()+"\t");
                }
            }
            System.out.println();
        }

    }


    private List<Emp> getEmps() {
        List<Emp> emps = new ArrayList<>();
        emps.add(new Emp(1, "jack1", 1000.0, 20,new Date()));
        emps.add(new Emp(2, "jack2", 1000.0, 20,new Date()));
        emps.add(new Emp(3, "jack3", 1000.0, 20,new Date()));
        emps.add(new Emp(4, "jack4", 1000.0, 20,new Date()));
        emps.add(new Emp(5, "jack5", 1000.0, 20,new Date()));
        emps.add(new Emp(6, "jack6", 1000.0, 20,new Date()));
        emps.add(new Emp(7, "jack7", 1000.0, 20,new Date()));
        emps.add(new Emp(8, "jack8", 1000.0, 20,new Date()));
        emps.add(new Emp(9, "jack9", 1000.0, 20,new Date()));
        emps.add(new Emp(10, "jack10", 1000.0, 20,new Date()));

        return emps;
    }
}






















