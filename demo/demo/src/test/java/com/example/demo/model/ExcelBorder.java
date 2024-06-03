package com.example.demo.model;

import org.apache.poi.ss.usermodel.*;

public class ExcelBorder {

    public static void addBordersToSheet(Sheet sheet) {
        Workbook workbook = sheet.getWorkbook();
        CellStyle borderStyle = createBorderStyle(workbook);

        // Apply border style to all cells in the sheet
        for (Row row : sheet) {
            for (Cell cell : row) {
                cell.setCellStyle(borderStyle);
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
