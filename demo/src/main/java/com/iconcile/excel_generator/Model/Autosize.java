package com.iconcile.excel_generator.Model;

import org.apache.poi.ss.usermodel.*;

public class Autosize {

    public static void autoSizeColumns(Sheet sheet) {
        // Adjust column widths for header columns
        Row headerRow = sheet.getRow(0);
        if (headerRow != null) {
            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                sheet.autoSizeColumn(i);
                int columnWidth = sheet.getColumnWidth(i);
                // Set a maximum width for the column (optional).
                if (columnWidth > 25000) {
                    columnWidth = 25000;    
                }
                sheet.setColumnWidth(i, columnWidth);
            }
        }

        // Adjust column widths for data columns
        for (int i = 0; i < sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    Cell cell = row.getCell(j);
                    if (cell != null) {
                        int columnWidth = sheet.getColumnWidth(j);
                        int newColumnWidth = (int) ((cell.toString().length() + 1) * 256); // Adjust the multiplier as needed
                        if (newColumnWidth > columnWidth) {
                            sheet.setColumnWidth(j, newColumnWidth);
                        }
                    }
                }
            }
        }
    }
}