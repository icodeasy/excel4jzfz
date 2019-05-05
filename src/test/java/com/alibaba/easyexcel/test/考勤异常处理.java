package com.alibaba.easyexcel.test;


import com.alibaba.easyexcel.test.util.CellUtil;
import com.alibaba.easyexcel.test.util.FileUtil;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

/**
 * Created by hcb on 2019/3/29.
 */
public class 考勤异常处理 {

    private static final String FILE_NAME = "./表A.xlsx";
    private static final String OUT_FILE_NAME = "result.xlsx";

    private static final Set<Integer> columnNums2copy = new HashSet() {
        {
            add(0);// 员工编码
            add(1); // 姓名
            add(2); // 考勤日期
            add(4); // 组织名称
            add(7); // 班次名称
            add(8); // 上下班时间
            add(10); // 第一段上班时间
            add(13); // 第二段下班时间
            add(19); // 缺卡次数
            add(20); // 补卡次数
            add(21); // 出差次数
            add(27); // 迟到分钟
            add(29); // 早退分钟
            add(33); // 旷工天数
            add(44); // 请假次数
        }

    };

    public static void main(String[] args) throws IOException {
        // 读取'考勤结果明细'
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

        Sheet copy2sheet = workbook.createSheet("总数据");
        Sheet originalSheet = workbook.getSheetAt(0);
        List<Row> sortedRows = 排序以及过滤考勤明细(originalSheet);

        拷贝到新的表格_同时进行必要的处理(copy2sheet, sortedRows);

        FileOutputStream fileOut = new FileOutputStream(OUT_FILE_NAME);
        workbook.write(fileOut);
        fileOut.flush();
        fileOut.close();

    }

    private static List<Row> 排序以及过滤考勤明细(Sheet originalSheet) {
        // rid = 6 表示从table body开始
        List<Row> bodyRows = new ArrayList<>();
        for (int rId = 5; rId <= originalSheet.getLastRowNum(); rId++) {
            Row row = originalSheet.getRow(rId);

//            if (!读取厦门员工编号.contains(Long.valueOf(row.getCell(0).getStringCellValue()))) {
//                // 排除非厦门的员工
//                continue;
//            }

            if (排除周末没有加班的记录(row)) {
                continue;
            }

            bodyRows.add(row);
        }

        Collections.sort(bodyRows, (o1, o2) -> {
            // 按照组织名称，班次，姓名，考勤日期进行排序
            int idx4org = 4;
            int idx4name = 1;
            int idx4kaoqin = 2;
            int orgRst = o1.getCell(idx4org).getStringCellValue().compareTo(o2.getCell(idx4org).getStringCellValue());
            if (orgRst == 0) {
                int nameRst = o1.getCell(idx4name).getStringCellValue().compareTo(o2.getCell(idx4name).getStringCellValue());
                if (nameRst == 0) {
                    return o1.getCell(idx4kaoqin).getDateCellValue().compareTo(o2.getCell(idx4kaoqin).getDateCellValue());
                } else {
                    return nameRst;
                }
            } else {
                return orgRst;
            }
        });

        // 最后把表头标题加上
        bodyRows.add(0, originalSheet.getRow(4));

        return bodyRows;
    }

    private static void 拷贝到新的表格_同时进行必要的处理(Sheet copy2sheet, List<Row> fromSortedRows) {

        int copiedRowId = 0;

        for (int rId = 0; rId < fromSortedRows.size(); rId++) {

            Row row = fromSortedRows.get(rId);

            if (排除考勤正常的记录(rId, row)) {
                continue;
            }

            Row finalSheetRow = copy2sheet.createRow(copiedRowId++);
            int copiedCellId = 0;
            for (int cID = 0; cID < row.getLastCellNum(); cID++) {
                if (!columnNums2copy.contains(cID)) {
                    continue;
                }
                Cell cOld = row.getCell(cID);
                if (cID != 4) {
                    if (cOld != null) {
                        Cell cNew = finalSheetRow.createCell(copiedCellId++, cOld.getCellTypeEnum());
                        cloneCell(cNew, cOld);
                    }
                    if (cID == 13) { // 特殊处理
                        新增前一天下班时间(fromSortedRows, finalSheetRow, rId, copiedCellId++);
                    }
                } else {
                    copiedCellId = 拆分组织名称(finalSheetRow, cOld, copiedCellId);
                }

            }
        }
    }

    /**
     * 15.	筛选“缺卡次数”为空白，“迟到分钟数”为空白，“旷工天数”为空白，“早退分钟数”为空白，删除这些行（此为考勤没有异常的，不需记录）；
     *
     * @param originSheetRow
     * @return
     */
    private static boolean 排除考勤正常的记录(int idx, Row originSheetRow) {
        if (idx > 0) {
            double qukaValue = originSheetRow.getCell(19).getNumericCellValue();
            double chidaoValue = originSheetRow.getCell(26).getNumericCellValue();
            double zaotuiValue = originSheetRow.getCell(28).getNumericCellValue();
            double kuanggongValue = originSheetRow.getCell(32).getNumericCellValue();
            if (qukaValue == 0.0d && chidaoValue == 0.0d && zaotuiValue == 0.0d && kuanggongValue == 0.0d) {
                return true;
            }
        }

        return false;
    }

    private static boolean 排除周末没有加班的记录(Row originSheetRow) {
        Cell on1stCell = originSheetRow.getCell(10);
        double on1st = on1stCell != null ? on1stCell.getNumericCellValue() : 0;
        Cell off1stCell = originSheetRow.getCell(11);
        double off1st = off1stCell != null ? off1stCell.getNumericCellValue() : 0;
        Cell on2ndCell = originSheetRow.getCell(12);
        double on2nd = on2ndCell != null ? on2ndCell.getNumericCellValue() : 0;
        Cell off2ndCell = originSheetRow.getCell(13);
        double off2nd = off2ndCell != null ? off2ndCell.getNumericCellValue() : 0;
        if (on1st == 0.0d && off1st == 0.0d && on2nd == 0.0d && off2nd == 0.0d) {
            return true;
        } else {
            return false;
        }
    }

    private static void 新增前一天下班时间(List<Row> fromSortedRows, Row finalSheetRow, int rId, int copiedCellId) {
        Cell cell = finalSheetRow.createCell(copiedCellId);
        if (rId >= 1) {
            long code = CellUtil.getLongValue(fromSortedRows.get(rId).getCell(0));
            if (rId != 1) {
                long lastCode = CellUtil.getLongValue(fromSortedRows.get(rId - 1).getCell(0));
                if (lastCode == code) {
                    // 如果不是当前员工考勤的第一天
                    Cell oldCell = fromSortedRows.get(rId - 1).getCell(13);
                    cloneCell(cell, oldCell);
                } else {
                    // 考勤第一天（如果是月初，有些员工本月入职则不需要）需要从表b获取前一天下班时间
//                    Date day1st = fromSortedRows.get(rId).getCell(2).getDateCellValue();
//                    Calendar cal = Calendar.getInstance();
//                    cal.setTime(day1st);
//                    int dayOfMonth = cal.get(Calendar.DAY_OF_MONTH);
//                    if (dayOfMonth == 1) {
                        Cell lastWorkDate = 读取本月初前一天下班时间.getLastWorkDate(code);
                        if (lastWorkDate != null) {
                            cloneCell(cell, lastWorkDate);
                        }
//                    }
                }
            } else {
                // 如果是第一行，直接从表b获取前一天下班时间
                Cell lastWorkDate = 读取本月初前一天下班时间.getLastWorkDate(code);
                cloneCell(cell, lastWorkDate);
            }
        } else {
            cell.setCellValue("前一天下班时间");

        }

    }

    private static int 拆分组织名称(Row finalSheetRow, Cell cOld, int copiedCellId) {
        String[] split = null;
        if (cOld.getStringCellValue().contains("_")) {
            split = cOld.getStringCellValue().substring(5).split("_");
        } else {
            split = new String[3];
            split[0] = "公司";
            split[1] = "部";
            split[2] = "室";
        }


        for (int i = 0; i < 3; i++) {
            Cell cNew = finalSheetRow.createCell(copiedCellId++, cOld.getCellTypeEnum());
            cNew.setCellComment(cOld.getCellComment());
            cNew.setCellStyle(cOld.getCellStyle());
            if (split.length < 3 && i == 2) {
                cNew.setCellValue("");
            } else {
                cNew.setCellValue(split[i]);
            }
        }

        return copiedCellId;
    }

    private static void 工号转数字(Iterator<Row> iterator, CellStyle style) {
        while (iterator.hasNext()) {

            Row currentRow = iterator.next();
            Cell currentRowCell = currentRow.getCell(0);

            try {

                String stringCellValue = currentRowCell.getStringCellValue();
                Double numericCellValue = Double.parseDouble(stringCellValue);
                Cell cell = currentRow.createCell(0);
                cell.setCellValue(numericCellValue);
                // 设置居中
                cell.setCellStyle(style);
            } catch (Throwable t) {
                System.out.println("ignore text cell");
            }
        }
    }

    private static void cloneCell(Cell cNew, Cell cOld) {
        cNew.setCellComment(cOld.getCellComment());
        try {
            cNew.setCellStyle(cOld.getCellStyle());
        } catch (Throwable t) {
            //   Exception in thread "main" java.lang.IllegalArgumentException: This Style does not belong to the supplied Workbook Stlyes Source. Are you trying to assign a style from one workbook to the cell of a differnt workbook?

        }

        if (CellType.BOOLEAN == cNew.getCellTypeEnum()) {
            cNew.setCellValue(cOld.getBooleanCellValue());
        } else if (CellType.NUMERIC == cNew.getCellTypeEnum()) {
            cNew.setCellValue(cOld.getNumericCellValue());
        } else if (CellType.STRING == cNew.getCellTypeEnum()) {
            cNew.setCellValue(cOld.getStringCellValue());
        } else if (CellType.ERROR == cNew.getCellTypeEnum()) {
            cNew.setCellValue(cOld.getErrorCellValue());
        } else if (CellType.FORMULA == cNew.getCellTypeEnum()) {
            cNew.setCellValue(cOld.getCellFormula());
        } else if (HSSFDateUtil.isCellDateFormatted(cOld)) {
            cNew.setCellValue(cOld.getDateCellValue());
        } else {
            cOld.setCellType(CellType.NUMERIC);
            cNew.setCellValue(cOld.getNumericCellValue());
        }
    }

}
