package com;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.LinkedList;

import static com.ReadExcel.firmLinkedList;
import static com.ReadExcel.getAllFirm;


public class WriteExcel {

    //筛选数据
    public static LinkedList<Firm> selectData() throws IOException {


        getAllFirm();
        LinkedList<Firm> firms = firmLinkedList;
        System.out.println(firms.size());
        System.out.println("==========================");

        for(int i = 0;i<firms.size();i++){
            firms.get(i).SAVED = true;
            //sale<0  cost<0
            if(firms.get(i).firm_sales<=0 || firms.get(i).firm_cost<=0){
                firms.get(i).SAVED = false;
            }
        }
        for(int j = 1;j<firms.size();j++){
            //sale变化超50%
            if(getChange(firms.get(j-1).firm_sales,firms.get(j).firm_sales)>0.5){
                firms.get(j).SAVED = false;
            }
            //非同增同减的即 变化率相乘为负
            else if(getChange(firms.get(j-1).firm_sales,firms.get(j).firm_sales)*getChange(firms.get(j-1).firm_cost,firms.get(j).firm_cost)<0){
                firms.get(j).SAVED = false;
            }
        }
        for(int k = 0;k<firms.size();k++){
            if(!firms.get(k).SAVED){
                firms.remove(k);
            }
        }
        return firms;
    }

    //计算涨幅
    public static  float getChange(long a,long b){
        return (float)(b-a)/b;
    }

    public static void main(String[] args) throws IOException {
        LinkedList<Firm> firms = selectData();
        System.out.println(firms.size());


        //创建工作簿
        XSSFWorkbook workbook = new XSSFWorkbook();
        //创建工作表
        XSSFSheet sheet = workbook.createSheet("筛选后的数据");

        for(int i = 0; i < firms.size(); i++){
            //创建行
            XSSFRow row = sheet.createRow(i);
            //创建单元格
            row.createCell(0).setCellValue(firms.get(i).firm_id);
            row.createCell(1).setCellValue(firms.get(i).firm_date);
            row.createCell(2).setCellValue(firms.get(i).firm_sales);
            row.createCell(3).setCellValue(firms.get(i).firm_cost);
            row.createCell(4).setCellValue(firms.get(i).SOE);
            row.createCell(5).setCellValue(firms.get(i).firm_a);
            row.createCell(6).setCellValue(firms.get(i).firm_b);
            row.createCell(7).setCellValue(firms.get(i).firm_c);
            row.createCell(8).setCellValue(firms.get(i).firm_interest);
            row.createCell(9).setCellValue(firms.get(i).firm_employeeNum);
            row.createCell(10).setCellValue(firms.get(i).firm_asset);
            row.createCell(11).setCellValue(firms.get(i).firm_tax);
            row.createCell(12).setCellValue(firms.get(i).firm_work_fee);
            row.createCell(13).setCellValue(firms.get(i).firm_sales_fee);
            row.createCell(14).setCellValue(firms.get(i).firm_manage_fee);
            row.createCell(15).setCellValue(firms.get(i).firm_eco_fee);

        }

        //输出流
        FileOutputStream outputStream = new FileOutputStream("/Users/luckymeng/IdeaProjects/Excel_RAW/src/main/resources/result.xlsx");
        workbook.write(outputStream);

        //刷新，并释放资源
        outputStream.flush();
        outputStream.close();
        workbook.close();
        System.out.println("写入成功");


    }
}
