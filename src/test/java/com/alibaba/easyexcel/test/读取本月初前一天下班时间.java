package com.alibaba.easyexcel.test;

import com.alibaba.easyexcel.test.util.CellUtil;
import com.alibaba.easyexcel.test.util.FileUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

/**
 * Created by hcb on 2019/3/31.
 */
public class 读取本月初前一天下班时间 {

    private static final String FILE_NAME = "./表B.xlsx";

    public static Map<Long, Cell> lastWorkDayCells = new LinkedHashMap<>();

    static {
        Workbook workbook = null;
        try {

            InputStream fileInputStream = FileUtil.getResourcesFileInputStream(FILE_NAME);
            workbook = new XSSFWorkbook(fileInputStream);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        Sheet sheet = workbook.getSheetAt(0);
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            Cell codeCell = row.getCell(0);
            Cell lastWorkDayCell = row.getCell(13);
            try {
                lastWorkDayCells.put(CellUtil.getLongValue(codeCell), lastWorkDayCell);
            } catch (Throwable t) {
                System.err.println("error for reading:"+codeCell+";"+t.toString());
            }
        }
    }

    public static Cell getLastWorkDate(Long code) {
        return lastWorkDayCells.get(code);
    }

    public static void main(String[] args) {
        System.out.println(lastWorkDayCells.size());
        lastWorkDayCells.entrySet().stream().forEach(p->{
            System.out.println(p.getKey()+":"+p.getValue().toString());
        });
    }
}
