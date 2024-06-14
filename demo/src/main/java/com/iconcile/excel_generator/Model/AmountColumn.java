package com.iconcile.excel_generator.Model;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

public class AmountColumn {

    public static void handleAmountColumn(Row row, int colIdx, Object value, Workbook workbook, String currencySymbol) {
        Cell cell = row.createCell(colIdx);
        if (value != null) {
            // Add currency symbol only to non-null values
            cell.setCellValue(currencySymbol + value.toString());

            // Set cell format as currency
            CellStyle currencyStyle = workbook.createCellStyle();
            currencyStyle.setDataFormat((short) 7); // Currency format code
            cell.setCellStyle(currencyStyle);
        } else {
            // Handle null value
            cell.setCellValue((String) null);
        }
    }
}