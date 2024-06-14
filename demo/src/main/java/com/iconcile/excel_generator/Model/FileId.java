package com.iconcile.excel_generator.Model;

import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

public class FileId {

    public static void handleFileId(Row row, int colIdx, Object value, Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setDataFormat(workbook.createDataFormat().getFormat("0.00")); // Custom number format

        Cell cell = row.createCell(colIdx);
        if (value != null) {
            cell.setCellValue(Double.parseDouble(value.toString())); // Assuming FILE_ID is numeric
        }
        cell.setCellStyle(style);
    }

    public Date getId() {
        // TODO Auto-generated method stub
        throw new UnsupportedOperationException("Unimplemented method 'getId'");
    }
}
