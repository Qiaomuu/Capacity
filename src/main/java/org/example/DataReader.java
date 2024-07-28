package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class DataReader {

    private static LocalDate endTime;
    private static LocalDate startTime;

    private static String st;
    private static String et;

    /*public DataReader() {
        endTime = getCurrentTime();
        startTime = endTime.minusYears(5);
    }
*/
    public static void main(String[] args) throws IOException {
        HashMap<String, HashMap> map = new HashMap<>();
        endTime = getCurrentTime();
        startTime = endTime.minusYears(5);
        String fileName = "E:\\test.xlsx";
        readDate(map, fileName);

    }

    public static LocalDate getCurrentTime() {
        LocalDate date = LocalDate.now();
        return date;
    }

    public static void readDate(HashMap<String, HashMap> map, String fileName) throws IOException {
        File file = new File(fileName);
        if (!file.exists()) {
            System.out.println("文件不存在");
            return;
        }
        /**
         * 判断文件后缀名是否合法；
         */
        FileInputStream inputStream = new FileInputStream(file);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);
        /**
         * 第一列为编码，第二列为开始时间，第三列结束时间；
         */
        Integer num = sheet.getLastRowNum();
        for (int i = 1; i <= num; i++) {
            Row row = sheet.getRow(i);
            Cell bom = row.getCell(0);
            Cell stCell = row.getCell(1);
            st = getValue(stCell);
            Cell etCell = row.getCell(2);
            et = getValue(etCell);

            if(st.compareTo(startTime.toString())<0){
                st=startTime.toString();
            }
            if(et.compareTo(endTime.toString())>0){
                et=endTime.toString();
            }
            //如果编码的维保时间未在我们统计的时间内，则跳过该条数据；
            if(st.compareTo(et)>0){
                continue;
                /*System.out.print(bom.getStringCellValue()+" ");
                System.out.print(st+" ");
                System.out.println(et);*/
            }
            match(bom.getStringCellValue(),st,et,map);

        }


    }
    public static String getValue(Cell cell){
        String cellValue = "";
        if(cell != null){
            CellType cellType = cell.getCellType();
            switch(cellType){
                case STRING:
                    cellValue = cell.getStringCellValue();
                    break;
                case NUMERIC:
                    /*cellValue = String.valueOf(cell.getNumericCellValue());*/
                    if (DateUtil.isCellDateFormatted(cell)) { //日期
                        Date date = cell.getDateCellValue();
                        cellValue = new DateTime(date).toString("yyyy-MM-dd");
                    } else {
                        cell.setCellType(CellType.STRING);
                        cellValue = cell.toString();
                    }
                    break;
            }
        }
        return cellValue;
    }
    /*public static Boolean timeProcessor(String st,String et){
        *//**
         *如果st<startTime,st=startTime
         * 如果et>endTime,et=endTime
         *//*
        if(st.compareTo(startTime.toString())<0){
            st=startTime.toString();
            System.out.println("更新维保开始时间"+st);
        }
        if(et.compareTo(endTime.toString())>0){
            et=endTime.toString();
            System.out.println("更新维保结束时间"+et);
        }
        if(st.compareTo(et)>0){
            return false;
        }
        return true;
    }*/
    public static void match(String bom,String st,String et,HashMap<String, HashMap> map){
        //先看看bom有没有在map中
        if(!map.containsKey(bom)){
            //先针对该BOM创建一个空map;
            map.put(bom,new HashMap<String,Integer>());
            //取出来该map，填充它，填好之后再放进总map中;
            HashMap<String,Integer> bomMap = map.get(bom);
            LocalDate sy = LocalDate.parse(st, DateTimeFormatter.ofPattern("yyyy-MM-dd"));
            LocalDate ey = LocalDate.parse(et, DateTimeFormatter.ofPattern("yyyy-MM-dd"));
            for(LocalDate ld=sy;ld.isBefore(ey) || ld.isEqual(ey);ld=ld.plusMonths(1)){
                System.out.println("维保期内年月日"+ld.toString());
            }
            System.out.println("------------------------");
        }
        /**
         * 先按年统计
         * 计算出需要统计的几个年份
         * 写一个方法，计算出开始时间-结束时间包含的年份
         */
        /*Integer[] arr = new Integer[]{2019,2020,2021,2022,2023,2024};
        List<Integer> list = new ArrayList<>();
        Collections.addAll(list,arr);
        */
        /*LocalDate parse = LocalDate.parse(st, DateTimeFormatter.ofPattern("yyyy"));*/



    }
}

