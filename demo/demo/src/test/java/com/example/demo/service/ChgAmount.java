
package com.example.demo.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

public class ChgAmount {

    public static void handleChgAmount(Row row, int colIdx, Object value, Workbook workbook) {
        Cell cell = row.createCell(colIdx);
        if (value != null) {
            // Add "$" symbol only to non-null values
            cell.setCellValue("$" + value.toString());

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



