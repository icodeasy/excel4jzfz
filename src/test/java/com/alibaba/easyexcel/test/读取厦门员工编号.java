package com.alibaba.easyexcel.test;

import com.alibaba.easyexcel.test.util.CellUtil;
import com.alibaba.easyexcel.test.util.FileUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashSet;
import java.util.Set;

/**
 * Created by hcb on 2019/3/31.
 */
public class 读取厦门员工编号 {

    private static final String FILE_NAME = "./厦门名单.xlsx";

    public static Set<Long> codes = new HashSet<>();

    static {
        Workbook workbook = null;
        try {

            InputStream fileInputStream = FileUtil.getResourcesFileInputStream(FILE_NAME);
            workbook = WorkbookFactory.create(fileInputStream);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }

        Sheet sheet = workbook.getSheet("员工入职");
        System.out.println("last row num --->:"+sheet.getLastRowNum());
        for (int i = 0; i < sheet.getLastRowNum()+1; i++) {
            Row row = sheet.getRow(i);
            Cell cell = row.getCell(1);
            try {
                Long code = CellUtil.getLongValue(cell);
                codes.add(code);
            } catch (Throwable t) {
                System.out.println(t.toString());
            }
        }
    }

    public static boolean contains(Long code) {
        return codes.contains(code);
    }

    public static void main(String[] args) {
        System.out.println(codes);
        System.out.println(codes.size());
    }
}
