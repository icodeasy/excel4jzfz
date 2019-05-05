package com.alibaba.easyexcel.test.util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

/**
 * Created by hcb on 2019/4/7.
 */
public class CellUtil {
    public static Long getLongValue(Cell cell) {
        if (CellType.NUMERIC == cell.getCellTypeEnum()) {
            return (long) cell.getNumericCellValue();
        } else if (CellType.STRING == cell.getCellTypeEnum()) {
            return Long.parseLong(cell.getStringCellValue());
        } else {
            return 0l;
        }
    }
}
