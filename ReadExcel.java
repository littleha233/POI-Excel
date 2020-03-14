package com;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;

public class ReadExcel {


    public  static LinkedList<Firm> firmLinkedList = new LinkedList<>();


    //存入firm数组
    public static void getAllFirm() throws IOException {
        String filePath = "/Users/luckymeng/IdeaProjects/Excel_RAW/src/main/resources/demo.xlsx";

        //获取工作簿对象
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(filePath);
        //获取工作表对象
        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
    //    System.out.println("...");
//        //获取行
//        for(Row row : sheet){
//            //获取单元格
//            for(Cell cell : row){
//                String value = cell.getStringCellValue();
//                System.out.println(value);
//            }
//        }
        int lastRowNum = sheet.getLastRowNum();



        for(int i = 1;i <= lastRowNum; i++){
            Firm firm = new Firm();
            XSSFRow row = sheet.getRow(i);
            if(row != null){
                int lastCellNum = row.getLastCellNum();
                for(int j = 0;j <=lastCellNum; j++){

                    //第一列
                    XSSFCell cell1 = row.getCell(0);
                    if(cell1 !=null){
                        double cellValue = cell1.getNumericCellValue();
                        firm.firm_id = (int) cellValue;
                    }

                    //第二列
                    XSSFCell cell2 = row.getCell(1);
                    if(cell2 !=null){
                        firm.firm_date= cell2.getStringCellValue();
                    }

                    //第三列
                    XSSFCell cell3 = row.getCell(2);
                    if(cell3 !=null){
                        double cellValue = cell3.getNumericCellValue();
                        firm.firm_sales = (long) cellValue;
                    }

                    //第四列
                    XSSFCell cell4 = row.getCell(3);
                    if(cell4 != null){
                        double cellValue = cell4.getNumericCellValue();
                        firm.firm_cost = (long) cellValue;
                    }

                    //第五列
                    XSSFCell cell5 = row.getCell(4);
                    if(cell5 != null){
                        double cellValue = cell5.getNumericCellValue();
                        firm.SOE  = (int) cellValue;
                    }

                    //第六列
                    XSSFCell cell6 = row.getCell(5);
                    if(cell6 != null){
                        firm.firm_a = cell6.getStringCellValue();
                    }

                    //第七列
                    XSSFCell cell7 = row.getCell(6);
                    if(cell7 != null){
                        firm.firm_b = cell7.getStringCellValue();
                    }

                    //第八列
                    XSSFCell cell8 = row.getCell(7);
                    if(cell8 != null){
                        firm.firm_c = cell8.getStringCellValue();
                    }

                    //第九列
                    XSSFCell cell9 = row.getCell(8);
                    if(cell9 != null){
                        double cellValue = cell9.getNumericCellValue();
                        firm.firm_interest = (int) cellValue;
                    }

                    //第十列
                    XSSFCell cell10 = row.getCell(9);
                    if(cell10 != null){
                        double cellValue = cell10.getNumericCellValue();
                        firm.firm_employeeNum = (long) cellValue;
                    }

                    //第十一列
                    XSSFCell cell11 = row.getCell(10);
                    if(cell11 != null){
                        double cellValue = cell11.getNumericCellValue();
                        firm.firm_asset = (long) cellValue;
                    }

                    //第十二列  营业税金及附加
                    XSSFCell cell12 = row.getCell(11);
                    if(cell12 != null){
                        double cellValue = cell12.getNumericCellValue();
                        firm.firm_tax = (long) cellValue;
                    }

                    //第十三列   业务及管理费
                    XSSFCell cell13 = row.getCell(12);
                    if(cell13 != null){
                        double cellValue = cell13.getNumericCellValue();
                        firm.firm_work_fee = (long) cellValue;
                    }

                    //第十四列  销售费用
                    XSSFCell cell14 = row.getCell(13);
                    if(cell14 != null){
                        double cellValue = cell14.getNumericCellValue();
                        firm.firm_sales_fee = (long) cellValue;
                    }

                    //第十五列  管理费用
                    XSSFCell cell15 = row.getCell(14);
                    if(cell15 != null){
                        double cellValue = cell15.getNumericCellValue();
                        firm.firm_manage_fee = (long) cellValue;
                    }

                    //第十六列 财务费用
                    XSSFCell cell16 = row.getCell(15);
                    if(cell16 != null){
                        double cellValue = cell16.getNumericCellValue();
                        firm.firm_eco_fee = (long) cellValue;
                    }

                }
                firmLinkedList.add(firm);
            }
        }


        //释放资源
        xssfWorkbook.close();
    }


}
