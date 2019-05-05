package com.alibaba.easyexcel.test;

import com.alibaba.easyexcel.test.listen.ExcelListener;
import com.alibaba.easyexcel.test.util.FileUtil;
import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Sheet;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.*;
import java.util.List;

/**
 * Created by hcb on 2019/4/14.
 */
public class BirthdayBonusGenerator {

    private static final String SOURCE_FILE_NAME = "./员工入职.xlsx";
    private static final String GENERATE_FILE_NAME = "生日礼金.xlsx";
    public static void main(String[] args) throws IOException {
        // 读取'考勤结果明细'
        Workbook workbook = null;
        try {

            InputStream fileInputStream = FileUtil.getResourcesFileInputStream(SOURCE_FILE_NAME);
            workbook = WorkbookFactory.create(fileInputStream);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }

        org.apache.poi.ss.usermodel.Sheet birthday = workbook.createSheet("生日礼金");
        org.apache.poi.ss.usermodel.Sheet originalSheet = workbook.getSheetAt(0);



        FileOutputStream fileOut = new FileOutputStream(GENERATE_FILE_NAME);
        workbook.write(fileOut);
        fileOut.flush();
        fileOut.close();

    }

    private void generateBirthdaySheet(Sheet origin, Sheet birthday) {

    }
}
