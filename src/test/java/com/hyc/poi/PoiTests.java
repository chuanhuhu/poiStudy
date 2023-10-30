package com.hyc.poi;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.*;

public class PoiTests {
    public static final String PATH = "D:\\Project-HYC\\code\\java springboot 后台开发通用模板\\easyexcel\\src\\test\\java\\com\\hyc\\excel\\";

    /**
     * POI写操作基本用法
     */
    @Test
    public void HSSFWorkbook03() {
        FileOutputStream fileOutputStream = null;
        try {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.createSheet("测试一");
            HSSFRow row = sheet.createRow(0);
            HSSFCell cell = row.createCell(0);
            cell.setCellValue("66666");
            fileOutputStream = new FileOutputStream(PATH + "HSSFWorkbook03.xls");
            workbook.write(fileOutputStream);
        } catch (Exception e) {
            throw new RuntimeException(e);
        } finally {
            try {
                if (fileOutputStream != null) {
                    fileOutputStream.close();
                }
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }
    }

    @Test
    public void XSSFWorkbook07() {
        FileOutputStream fileOutputStream = null;
        try {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet();
            XSSFRow row = sheet.createRow(0);
            XSSFCell cell = row.createCell(0);
            cell.setCellValue("66666");
            fileOutputStream = new FileOutputStream(PATH + "XSSFWorkbook07.xlsx");
            workbook.write(fileOutputStream);
            workbook.close();
        } catch (Exception e) {
            throw new RuntimeException(e);
        } finally {
            try {
                if (fileOutputStream != null) {
                    fileOutputStream.close();
                }
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }
    }

    @Test
    public void SXSSFWorkbook07() {
        FileOutputStream fileOutputStream = null;
        try {
            SXSSFWorkbook workbook = new SXSSFWorkbook();
            SXSSFSheet sheet = workbook.createSheet();
            SXSSFRow row = sheet.createRow(0);
            SXSSFCell cell = row.createCell(0);
            cell.setCellValue("66666");
            fileOutputStream = new FileOutputStream(PATH + "SXSSFWorkbook07.xlsx");
            workbook.write(fileOutputStream);
        } catch (Exception e) {
            throw new RuntimeException(e);
        } finally {
            try {
                if (fileOutputStream != null) {
                    fileOutputStream.close();
                }
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }
    }

    /**
     * POI读操作基本用法
     */
    @Test
    public void ExcelRead07() throws Exception {
        FileInputStream fileInputStream = new FileInputStream(PATH + "SXSSFWorkbook07.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheetAt = workbook.getSheetAt(0);
        XSSFRow row = sheetAt.getRow(0);
        XSSFCell cell = row.getCell(0);
        String stringCellValue = cell.getStringCellValue();
        System.out.println(stringCellValue);
    }

    @Test
    public void ListExcelRead07() throws Exception {
        FileInputStream fileInputStream = new FileInputStream(PATH + "text.xlsx");
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);
        int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
        for (int i = 0; i < physicalNumberOfRows; i++) {
            Row row = sheet.getRow(i);
            int physicalNumberOfCells = row.getPhysicalNumberOfCells();
            for (int j = 0; j < physicalNumberOfCells; j++) {
                Cell cell = row.getCell(j);
                switch (cell.getCellType()) {
                    //字符串
                    case STRING:
                        System.out.println(cell.getStringCellValue());
                        break;
                    //布尔值
                    case BOOLEAN:
                        cell.getBooleanCellValue();
                        break;
                    //空值
                    case BLANK:
                        System.out.println("空值");
                        break;
                    //数字类型
                    case NUMERIC:
                        System.out.println(DateUtil.isCellDateFormatted(cell));
                        if (DateUtil.isCellDateFormatted(cell)) {
                            System.out.println(cell.getDateCellValue());
                        } else {
                            System.out.println(cell.getNumericCellValue());
                        }
                        break;
                    default:
                        System.out.println("位置类型");
                }
            }
        }
        fileInputStream.close();

    }

}
