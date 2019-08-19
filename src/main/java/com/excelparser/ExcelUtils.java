package com.excelparser;

import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class ExcelUtils {
    static void filterRows(Sheet datatypeSheet) {
        int LastRowNum = datatypeSheet.getLastRowNum();
        for (int RowNum = 0; RowNum < LastRowNum - 1; RowNum++) {
            Row currentRow = datatypeSheet.getRow(RowNum);
            if (!containsKey(currentRow)) {
                datatypeSheet.removeRow(currentRow);
                continue;
            }
        }
    }

    static void removeEmptyRows(Sheet datatypeSheet) {
        boolean isRowEmpty = false;
        for (int rowIndex = 0; rowIndex < datatypeSheet.getLastRowNum(); rowIndex++) {
            if (datatypeSheet.getRow(rowIndex) == null) {
                isRowEmpty = true;
                datatypeSheet.shiftRows(rowIndex + 1, datatypeSheet.getLastRowNum(), -1);
                rowIndex--;
                continue;
            }
            isRowEmpty = isRowEmpty(datatypeSheet, rowIndex);
            if (isRowEmpty == true) {
                datatypeSheet.shiftRows(rowIndex + 1, datatypeSheet.getLastRowNum(), -1);
                rowIndex--;
            }
        }
    }

    private static boolean isRowEmpty(Sheet datatypeSheet, int rowIndex) {
        boolean isRowEmpty = false;
        for (int j = 0; j < datatypeSheet.getRow(rowIndex).getLastCellNum(); j++) {
            if (datatypeSheet.getRow(rowIndex).getCell(j).toString().trim().equals("")) {
                isRowEmpty = true;
            } else {
                isRowEmpty = false;
                break;
            }
        }
        return isRowEmpty;
    }

    private static boolean containsKey(Row currentRow) {
        boolean containsKey = false;
        Iterator<Cell> cellIterator = currentRow.iterator();
        while (cellIterator.hasNext()) {
            Cell currentCell = cellIterator.next();
            if (currentCell.getCellTypeEnum() == CellType.STRING) {
                String TextInCell = currentCell.toString();
                if (TextInCell.equalsIgnoreCase(App.KEY)) {
                    containsKey = true;
                }
            }
        }
        return containsKey;
    }
}