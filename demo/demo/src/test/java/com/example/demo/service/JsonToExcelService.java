package com.example.demo.service;

import java.awt.Color;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import com.example.demo.model.Camelcase;
import com.example.demo.model.ExcelBorder;
import com.example.demo.model.FreezeRowColumn;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

@Service
public class JsonToExcelService {

    private LabelName labelName;

    public ByteArrayInputStream convertJsonToExcel(String jsonString) throws IOException {
        try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet("Data");

            // Parse JSON string into JsonNode
            ObjectMapper objectMapper = new ObjectMapper();
            JsonNode rootNode = objectMapper.readTree(jsonString);

            // Extract REMOVE_COLUMNS list
            Set<String> removeColumns = RemoveColumn.extractRemovedColumns(rootNode);

            // Initialize LabelName and extract LABEL_NAME mapping
            labelName = new LabelName();
            labelName.extractLabelNameMapping(rootNode);

            List<Map<String, Object>> jsonData = objectMapper.convertValue(rootNode.get("data"), List.class);

            if (jsonData != null && !jsonData.isEmpty()) {
                // Generate header row
                Map<String, Integer> headerMapping = createHeaderRow(sheet, jsonData.get(0), removeColumns, workbook);

                // Generate data rows
                createDataRows(sheet, jsonData, headerMapping, removeColumns, workbook);

                // Autosize columns after filling data
                autoSizeColumns(sheet);

                // Extract freeze information from JSON data
                int freezeRows = rootNode.has("formatting") && rootNode.get("formatting").has("ROW_FREEZE")
                        ? rootNode.get("formatting").get("ROW_FREEZE").asInt()
                        : 0;
                int freezeColumns = rootNode.has("formatting") && rootNode.get("formatting").has("COLUMN_FREEZE")
                        ? rootNode.get("formatting").get("COLUMN_FREEZE").asInt()
                        : 3;

                // Freeze rows and columns
                FreezeRowColumn.freezeRowsAndColumns(sheet, freezeRows, freezeColumns);

                // Add border to all cells
                ExcelBorder.addBordersToSheet(sheet);

                

                // Apply background colors to the sheet
                JsonNode backgroundFill = rootNode.path("formatting").path("BACKGROUND_FILL");
                applyBackgroundColorToSheet(sheet, workbook, backgroundFill);
            }

            workbook.write(out);

            return new ByteArrayInputStream(out.toByteArray());
        }
        
        catch (IOException e) {
            System.out.println("IOException: " + e.getMessage());
            e.printStackTrace();
            throw e;
        }

    }

    
    private void applyBackgroundColorToSheet(Sheet sheet, Workbook workbook, JsonNode backgroundFill) {
        if (workbook instanceof XSSFWorkbook) {
            for (JsonNode fillConfig : backgroundFill) {
                String colorHex = fillConfig.path("COLOUR").asText();
                Color awtColor = getColorFromHex(colorHex);
                if (awtColor == null) {
                    System.out.println("Skipping invalid color: " + colorHex);
                    continue; // Skip if color is invalid
                }
    
                XSSFColor xssfColor = new XSSFColor(awtColor, new DefaultIndexedColorMap());
                String column = fillConfig.path("COLUMN").asText(null);
                Integer row = fillConfig.path("ROW").isInt() ? fillConfig.path("ROW").asInt() - 1 : null; // Adjust row index to 0-based
                String condition = fillConfig.path("CONDITION").asText(null);
    
                try {
                    // Apply color based on the configuration
                    if (column != null && !column.isEmpty()) {
                        int columnIndex = -1;
                        // Find the correct column index dynamically based on the header text
                        for (Row firstRow : sheet) {
                            for (Cell headerCell : firstRow) {
                                if (headerCell.getStringCellValue().trim().equalsIgnoreCase(column.trim())) {
                                    columnIndex = headerCell.getColumnIndex();
                                    break;
                                }
                            }
                            if (columnIndex != -1) {
                                break;
                            }
                        }
    
                        if (columnIndex == -1) {
                            throw new IllegalArgumentException("Column '" + column + "' not found in the sheet.");
                        }
                        for (Row currentRow : sheet) {
                            Cell cell = currentRow.getCell(columnIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                            if (isCellMatchingCondition(cell, condition)) {
                                System.out.println("Applying color to column " + column + " at row " + currentRow.getRowNum());
                                applyCellStylePreservingBorders(cell, xssfColor, workbook);
                            }
                        }
                    } else if (row != null) {
                        Row currentRow = sheet.getRow(row);
                        if (currentRow != null) {
                            for (Cell cell : currentRow) {
                                if (isCellMatchingCondition(cell, condition)) {
                                    System.out.println("Applying color to row " + row + " at column " + cell.getColumnIndex());
                                    applyCellStylePreservingBorders(cell, xssfColor, workbook);
                                }
                            }
                        }
                    }
                } catch (IllegalArgumentException e) {
                    // Handle invalid column index exception
                    System.out.println("Invalid column index exception: " + e.getMessage());
                    e.printStackTrace(); // Print stack trace for debugging
                } catch (Exception e) {
                    // Handle other exceptions
                    e.printStackTrace();
                }
            }
        }
    }
    

    private String getCellValueAsString(Cell cell) {
        DataFormatter formatter = new DataFormatter();
        return formatter.formatCellValue(cell);
    }

    
    private boolean isCellMatchingCondition(Cell cell, String condition) {
        if (condition == null || condition.trim().isEmpty()) {
            return true;
        }
        if (cell == null) {
            return false;
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim().equals(condition.trim());
            case NUMERIC:
                try {
                    double numericCondition = Double.parseDouble(condition.trim());
                    return cell.getNumericCellValue() == numericCondition;
                } catch (NumberFormatException e) {
                    return false;
                }
            case BOOLEAN:
                boolean booleanCondition = Boolean.parseBoolean(condition.trim());
                return cell.getBooleanCellValue() == booleanCondition;
            case FORMULA:
                try {
                    return cell.getStringCellValue().trim().equals(condition.trim());
                } catch (IllegalStateException e) {
                    try {
                        double numericCondition = Double.parseDouble(condition.trim());
                        return cell.getNumericCellValue() == numericCondition;
                    } catch (NumberFormatException e2) {
                        return false;
                    }
                }
            default:
                return false;
        }
    }
    
    

private void applyCellStylePreservingBorders(Cell cell, XSSFColor xssfColor, Workbook workbook) {
    XSSFCellStyle newCellStyle = (XSSFCellStyle) workbook.createCellStyle();
    XSSFCellStyle existingCellStyle = (XSSFCellStyle) cell.getCellStyle();

    // Clone the existing cell style
    newCellStyle.cloneStyleFrom(existingCellStyle);

    // Set the new background color
    newCellStyle.setFillForegroundColor(xssfColor);
    newCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

    cell.setCellStyle(newCellStyle);
}

private Color getColorFromHex(String colorHex) {
    try {
        return Color.decode(colorHex);
    } catch (NumberFormatException e) {
        return null;
    }
}

    
   
        private void autoSizeColumns(Sheet sheet) {
            int numColumns = sheet.getRow(0).getPhysicalNumberOfCells();
            for (int i = 0; i < numColumns; i++) {
                sheet.autoSizeColumn(i);
                int currentColumnWidth = sheet.getColumnWidth(i);
                sheet.setColumnWidth(i, Math.max(currentColumnWidth + 100, 7000)); // Adjust minimum width
            }
        }
    



    // Method to apply dynamic formatting based on JSON configuration
    private void applyDynamicFormatting(Sheet sheet, Workbook workbook, JsonNode formattingConfig) {
        if (formattingConfig.has("CONDITIONS")) {
            JsonNode conditions = formattingConfig.get("CONDITIONS");
            for (JsonNode condition : conditions) {
                String type = condition.path("TYPE").asText();
                if ("DATE_FORMATTING".equalsIgnoreCase(type)) {
                    applyDateFormatting(sheet, workbook, condition);
                }
                // Add more condition types as needed (e.g., numeric formatting, text formatting)
            }
        }
    }

    // Example method to apply date formatting based on JSON configuration
    private void applyDateFormatting(Sheet sheet, Workbook workbook, JsonNode condition) {
        String column = condition.path("COLUMN").asText();
        String dateFormat = condition.path("DATE_FORMAT").asText("dd-MM-yyyy");

        // Find column index dynamically based on column name
        int columnIndex = findColumnIndex(sheet, column);

        if (columnIndex != -1) {
            for (Row row : sheet) {
                Cell cell = row.getCell(columnIndex);
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    String cellValue = cell.getStringCellValue().trim();
                    try {
                        DateTimeFormatter formatter = DateTimeFormatter.ofPattern(dateFormat);
                        LocalDate date = LocalDate.parse(cellValue, formatter);

                        // Apply date formatting to the cell
                        CellStyle cellStyle = workbook.createCellStyle();
                        cellStyle.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat(dateFormat));
                        cell.setCellStyle(cellStyle);

                        System.out.println("Formatted cell (" + row.getRowNum() + ", " + columnIndex + "): " + cellValue);

                    } catch (DateTimeParseException e) {
                        System.out.println("Error parsing date: " + cellValue);
                    }
                }
            }
        }
    }

    // Utility method to find column index by column name
    private int findColumnIndex(Sheet sheet, String columnName) {
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (columnName.equalsIgnoreCase(cell.getStringCellValue().trim())) {
                    return cell.getColumnIndex();
                }
            }
        }
        return -1; // Column not found
    }

    
    
    
    private Map<String, Integer> createHeaderRow(Sheet sheet, Map<String, Object> firstRecord, Set<String> removeColumns, Workbook workbook) {
        Row parentHeaderRow = sheet.createRow(0);
        Row childHeaderRow = sheet.createRow(1);
        int colIdx = 0;
        Map<String, Integer> headerMapping = new HashMap<>();

        for (Map.Entry<String, Object> entry : firstRecord.entrySet()) {
            String key = entry.getKey();
            Object value = entry.getValue();

            // Replace the column name if it's in the labelNameMapping
            String originalKey = key; // Save the original key for data row reference
            String label = labelName.getLabelForColumn(key);
            if (label != null) {
                key = label;
            }

            // If the key is "SYS_VENDOR_CODE", replace it with "Vendor" (case-insensitive)
            if ("SYS_VENDOR_CODE".equalsIgnoreCase(originalKey.trim())) {
                key = "Vendor";
            }

            // Skip columns in the removeColumns set
            if (removeColumns.contains(originalKey)) { // Use original key for removal check
                continue;
            }

            // Create cells for header rows
            if (value instanceof List) {
                @SuppressWarnings("unchecked")
                List<Object> nestedList = (List<Object>) value;
                if (!nestedList.isEmpty() && nestedList.get(0) instanceof Map) {
                    int startColIdx = colIdx;
                    for (Map.Entry<String, Object> nestedEntry : ((Map<String, Object>) nestedList.get(0)).entrySet()) {
                        String nestedKey = nestedEntry.getKey();
                        if (removeColumns.contains(originalKey + "." + nestedKey)) { // Use original key for removal check
                            continue; // Skip nested columns
                        }
                        Cell childCell = childHeaderRow.createCell(colIdx++);
                        childCell.setCellValue(Camelcase.toCamelCaseWithSpaces(nestedKey));
                    }
                    Cell parentCell = parentHeaderRow.createCell(startColIdx);
                    parentCell.setCellValue(Camelcase.toCamelCaseWithSpaces(key));
                    sheet.addMergedRegion(new CellRangeAddress(0, 0, startColIdx, colIdx - 1));
                    headerMapping.put(originalKey, startColIdx); // Map original key to column index
                }
            } else if (value instanceof Map) {
                @SuppressWarnings("unchecked")
                Map<String, Object> nestedMap = (Map<String, Object>) value;
                int startColIdx = colIdx;
                for (Map.Entry<String, Object> nestedEntry : nestedMap.entrySet()) {
                    String nestedKey = nestedEntry.getKey();
                    if (removeColumns.contains(originalKey + "." + nestedKey)) { // Use original key for removal check
                        continue; // Skip nested columns
                    }
                    Cell childCell = childHeaderRow.createCell(colIdx++);
                    childCell.setCellValue(Camelcase.toCamelCaseWithSpaces(nestedKey));
                }
                Cell parentCell = parentHeaderRow.createCell(startColIdx);
                parentCell.setCellValue(Camelcase.toCamelCaseWithSpaces(key));
                sheet.addMergedRegion(new CellRangeAddress(0, 0, startColIdx, colIdx - 1));
                headerMapping.put(originalKey, startColIdx); // Map original key to column index
            } else {
                Cell cell = parentHeaderRow.createCell(colIdx);
                cell.setCellValue(Camelcase.toCamelCaseWithSpaces(key));
                sheet.addMergedRegion(new CellRangeAddress(0, 1, colIdx, colIdx));
                headerMapping.put(originalKey, colIdx); // Map original key to column index
                colIdx++;
            }
        }
        return headerMapping;
    }

    private void createDataRows(Sheet sheet, List<Map<String, Object>> jsonData, Map<String, Integer> headerMapping, Set<String> removeColumns, Workbook workbook) {
        int rowIdx = 2;
        for (Map<String, Object> record : jsonData) {
            Row row = sheet.createRow(rowIdx++);
            for (Map.Entry<String, Object> entry : record.entrySet()) {
                String key = entry.getKey();
                Object value = entry.getValue();

                if (removeColumns.contains(key)) {
                    continue; // Skip columns that are in the removeColumns set
                }

                if (headerMapping.containsKey(key)) {
                    int colIdx = headerMapping.get(key);

                    if ("FILE_ID".equals(key)) {
                        // Use FileId class to handle FILE_ID field
                        FileId.handleFileId(row, colIdx, value, workbook);
                    } else {
                        if (value instanceof List) {
                            @SuppressWarnings("unchecked")
                            List<Object> nestedList = (List<Object>) value;
                            
                            if (key.equals("RECON_STATUS")) {
                                // Handle RECON_STATUS field using RecordStatus class
                                RecordStatus.handleReconStatus(row, colIdx, nestedList);
                            } else {
                                for (Object listItem : nestedList) {
                                    if (listItem instanceof Map) {
                                        colIdx = fillRowWithMap(row, (Map<String, Object>) listItem, colIdx, key, removeColumns, workbook);
                                    }
                                }
                            }
                        } else if (value instanceof Map) {
                            @SuppressWarnings("unchecked")
                            Map<String, Object> nestedMap = (Map<String, Object>) value;
                            colIdx = fillRowWithMap(row, nestedMap, colIdx, key, removeColumns, workbook);
                        } else {

                            if ("CHG_AMOUNT".equals(key)) {
                                // Use ChgAmount class to handle CHG_AMOUNT field
                                ChgAmount.handleChgAmount(row, colIdx, value, workbook);
                            } else {
                                Cell cell = row.createCell(colIdx);
                                if (value != null) {
                                    cell.setCellValue(value.toString());
                                } else {
                                    cell.setCellValue((String) null); // Handle null value
                                }
                            }
                        }
                    }
                }
            }
        }
    }





    private int fillRowWithMap(Row row, Map<String, Object> nestedMap, int colIdx, String parentKey, Set<String> removeColumns, Workbook workbook) {
        for (Map.Entry<String, Object> nestedEntry : nestedMap.entrySet()) {
            String nestedKey = nestedEntry.getKey();
            if (removeColumns.contains(parentKey + "." + nestedKey)) {
                continue; // Skip nested columns that are in the removeColumns set
            }
            Object nestedValue = nestedEntry.getValue();
            Cell cell = row.createCell(colIdx++);
            if (nestedValue != null) {
                cell.setCellValue(nestedValue.toString());
            } else {
                cell.setCellValue((String) null); // Handle null value
            }
        }
        return colIdx;
    }
}
   