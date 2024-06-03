package com.example.demo.model;
import org.apache.poi.ss.usermodel.*;
import com.fasterxml.jackson.databind.JsonNode;

public class ExcelBackgroundFill {

    public static void applyBackgroundFill(Sheet sheet, JsonNode rootNode) {
        JsonNode backgroundFillNode = rootNode.path("formatting").path("BACKGROUND_FILL");
        if (backgroundFillNode.isArray()) {
            for (JsonNode fillCondition : backgroundFillNode) {
                String colour = fillCondition.path("COLOR").asText();
                String column = fillCondition.has("COLUMN") ? fillCondition.path("COLUMN").asText() : null;
                int row = fillCondition.has("ROW") && !fillCondition.path("ROW").isNull() ? fillCondition.path("ROW").asInt() : -1;
                String condition = fillCondition.path("CONDITION").asText();

                if (column != null) {
                    applyFillToColumn(sheet, colour, column, condition);
                } else if (row != -1) {
                    applyFillToRow(sheet, colour, row, condition);
                }
            }
        }
    }

    private static void applyFillToColumn(Sheet sheet, String colour, String columnName, String condition) {
        int columnIndex = getColumnIndex(sheet, columnName);
        if (columnIndex != -1) {
            for (Row row : sheet) {
                Cell cell = row.getCell(columnIndex);
                if (cell != null && condition.equals(cell.getStringCellValue())) {
                    applyFill(colour, cell);
                }
            }
        }
    }

    private static void applyFillToRow(Sheet sheet, String colour, int rowIndex, String condition) {
        Row row = sheet.getRow(rowIndex);
        if (row != null) {
            for (Cell cell : row) {
                if (condition.equals(cell.getStringCellValue())) {
                    applyFill(colour, cell);
                }
            }
        }
    }

    private static void applyFill(String colour, Cell cell) {
        CellStyle style = cell.getCellStyle();
        if (style == null) {
            style = cell.getSheet().getWorkbook().createCellStyle();
        }
        style.setFillForegroundColor(IndexedColors.valueOf(colour).getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cell.setCellStyle(style);
    }

    private static int getColumnIndex(Sheet sheet, String columnName) {
        Row firstRow = sheet.getRow(0);
        for (Cell cell : firstRow) {
            if (columnName.equals(cell.getStringCellValue())) {
                return cell.getColumnIndex();
            }
        }
        return -1; // Column not found
    }
}
