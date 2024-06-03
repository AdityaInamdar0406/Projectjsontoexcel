package com.example.demo.model;
import org.apache.poi.ss.usermodel.*;
import com.fasterxml.jackson.databind.JsonNode;

public class ExcelBackgroundFill {

    public static void applyBackgroundFill(Sheet sheet, JsonNode rootNode) {
        JsonNode backgroundFillNode = rootNode.path("formatting").path("BACKGROUND_FILL");
        if (backgroundFillNode.isArray()) {
            for (JsonNode fillCondition : backgroundFillNode) {
                String colour = fillCondition.path("COLOUR").asText();
                String column = fillCondition.has("COLUMN") ? fillCondition.path("COLUMN").asText() : null;
                Integer row = fillCondition.has("ROW") && !fillCondition.path("ROW").isNull() ? fillCondition.path("ROW").asInt() : null;
                String condition = fillCondition.path("CONDITION").asText();

                System.out.println("Processing fill condition:");
                System.out.println("COLOUR: " + colour);
                System.out.println("COLUMN: " + column);
                System.out.println("ROW: " + row);
                System.out.println("CONDITION: " + condition);

                if (column != null && row == null) {
                    applyFillToColumn(sheet, colour, column, condition);
                } else if (column == null && row != null) {
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
                if (cell != null) {
                    String cellValue = cell.getStringCellValue();
                    System.out.println("Cell value: " + cellValue);
                    if (condition == null || condition.equals(cellValue)) {
                        applyFill(colour, cell);
                    }
                }
            }
        }
    }

    private static void applyFillToRow(Sheet sheet, String colour, int rowIndex, String condition) {
        Row row = sheet.getRow(rowIndex);
        if (row != null) {
            for (Cell cell : row) {
                if (cell != null) {
                    String cellValue = cell.getStringCellValue();
                    System.out.println("Cell value: " + cellValue);
                    if (condition == null || condition.equals(cellValue)) {
                        applyFill(colour, cell);
                    }
                }
            }
        }
    }

    private static void applyFill(String colour, Cell cell) {
        Workbook workbook = cell.getSheet().getWorkbook();
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.valueOf(colour.toUpperCase()).getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cell.setCellStyle(style);
    }

    private static int getColumnIndex(Sheet sheet, String columnName) {
        Row firstRow = sheet.getRow(0);
        if (firstRow != null) {
            for (Cell cell : firstRow) {
                if (columnName.equals(cell.getStringCellValue())) {
                    return cell.getColumnIndex();
                }
            }
        }
        return -1; // Column not found
    }
}

