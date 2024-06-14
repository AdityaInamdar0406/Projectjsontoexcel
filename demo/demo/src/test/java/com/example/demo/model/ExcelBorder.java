package com.example.demo.model;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

public class ExcelBorder {

    public static void addBordersToSheet(Sheet sheet) {
        Workbook workbook = sheet.getWorkbook();
        CellStyle borderStyle = createBorderStyle(workbook);

        // Iterate over all rows and cells
        for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row == null) {
                row = sheet.createRow(rowNum);
            }
            for (int cellNum = 0; cellNum < row.getLastCellNum(); cellNum++) {
                Cell cell = row.getCell(cellNum);
                if (cell == null) {
                    cell = row.createCell(cellNum);
                }
                cell.setCellStyle(borderStyle);
            }
        }

        // Iterate over merged regions and apply borders to each cell in the region.
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress region = sheet.getMergedRegion(i);
            for (int row = region.getFirstRow(); row <= region.getLastRow(); row++) {
                Row r = sheet.getRow(row);
                if (r == null) {
                    r = sheet.createRow(row);
                }
                for (int col = region.getFirstColumn(); col <= region.getLastColumn(); col++) {
                    Cell cell = r.getCell(col);
                    if (cell == null) {
                        cell = r.createCell(col);
                    }
                    cell.setCellStyle(borderStyle);
                }
            }
        }
    }

    private static CellStyle createBorderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        return style;
    }

}