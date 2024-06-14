package com.iconcile.excel_generator.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.springframework.stereotype.Service;


import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.iconcile.excel_generator.Model.AmountColumn;
import com.iconcile.excel_generator.Model.Autosize;
import com.iconcile.excel_generator.Model.Camelcase;
import com.iconcile.excel_generator.Model.ExcelBorder;
import com.iconcile.excel_generator.Model.FileId;
import com.iconcile.excel_generator.Model.FreezeRowColumn;
import com.iconcile.excel_generator.Model.LabelName;
import com.iconcile.excel_generator.Model.ReconStatus;
import com.iconcile.excel_generator.Model.RemoveColumn;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.*;


@Service
public class JsonToExcelService {

    public ByteArrayInputStream convertJsonToExcel(String jsonString) throws IOException {
        try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet("Data");

            // Parse JSON string into JsonNode
            ObjectMapper objectMapper = new ObjectMapper();
            JsonNode rootNode = objectMapper.readTree(jsonString);

            // Extract LABEL_NAME mapping
            LabelName.extractLabelNameMapping(rootNode);

            // Extract REMOVE_COLUMNS list
            Set<String> removeColumns = RemoveColumn.extractRemovedColumns(rootNode);

            // Determine JSON format and extract data accordingly
            List<Map<String, Object>> jsonData;
            List<String> columns = new ArrayList<>();
            if (rootNode.has("data") && rootNode.get("data").isArray()) {
                // Original JSON format
                jsonData = objectMapper.convertValue(rootNode.get("data"), List.class);
            } else {
                // New JSON format
                JsonNode dataNode = rootNode.get("data");
                JsonNode columnsNode = dataNode.get("columns");
                for (JsonNode column : columnsNode) {
                    columns.add(column.asText());
                }
                jsonData = extractReportData(dataNode.get("report"), columns);
            }

            // Extract AMOUNT_COLUMNS from JSON data
            Map<String, String> amountColumns = new HashMap<>();
            if (rootNode.has("formatting") && rootNode.get("formatting").has("AMOUNT_COLUMNS")) {
                for (JsonNode amountColumnNode : rootNode.get("formatting").get("AMOUNT_COLUMNS")) {
                    String column = amountColumnNode.get("COLUMN").asText();
                    String currency = amountColumnNode.get("CURRENCY").asText();
                    amountColumns.put(column, currency);
                }
            }

            if (jsonData != null && !jsonData.isEmpty()) {
                // Generate header row
                Map<String, Integer> headerMapping = createHeaderRow(sheet, jsonData.get(0), removeColumns, workbook, columns);

                // Generate data rows
                createDataRows(sheet, jsonData, headerMapping, removeColumns, workbook, amountColumns);

                // Autosize columns after filling data
                Autosize.autoSizeColumns(sheet);

                // Extract freeze information from JSON data
                int freezeRows = rootNode.has("formatting") && rootNode.get("formatting").has("ROW_FREEZE") ? rootNode.get("formatting").get("ROW_FREEZE").asInt() : 0;
                int freezeColumns = rootNode.has("formatting") && rootNode.get("formatting").has("COLUMN_FREEZE") ? rootNode.get("formatting").get("COLUMN_FREEZE").asInt() : 0;

                // Freeze rows and columns
                FreezeRowColumn.freezeRowsAndColumns(sheet, freezeRows, freezeColumns);

                // Add border to all cells
                ExcelBorder.addBordersToSheet(sheet);

                // Apply background color to all cells (optional)
                // applyBackgroundColorToSheet(sheet, workbook, new XSSFColor(new java.awt.Color(220, 230, 241), null)); // Light blue color
            }

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        }
    }

    private List<Map<String, Object>> extractReportData(JsonNode reportNode, List<String> columns) {
        List<Map<String, Object>> data = new ArrayList<>();
        for (JsonNode report : reportNode) {
            Map<String, Object> record = new HashMap<>();
            for (String column : columns) {
                record.put(column, report.get(column));
            }
            data.add(record);
        }
        return data;
    }

    private Map<String, Integer> createHeaderRow(Sheet sheet, Map<String, Object> firstRecord, Set<String> removeColumns, Workbook workbook, List<String> columns) {
        Row parentHeaderRow = sheet.createRow(0);
        Row childHeaderRow = sheet.createRow(1);
        int colIdx = 0;
        Map<String, Integer> headerMapping = new HashMap<>();

        for (String column : columns.isEmpty() ? firstRecord.keySet() : columns) {
            Object value = firstRecord.get(column);
            String originalKey = column; // Save the original key for mapping

            // Get the label for the column
            String label = LabelName.getLabelForColumn(column);

            // Skip columns in the removeColumns set
            if (removeColumns.contains(originalKey)) {
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
                        if (removeColumns.contains(originalKey + "." + nestedKey)) {
                            continue; // Skip nested columns
                        }
                        Cell childCell = childHeaderRow.createCell(colIdx++);
                        childCell.setCellValue(Camelcase.toCamelCaseWithSpaces(nestedKey));
                    }
                    Cell parentCell = parentHeaderRow.createCell(startColIdx);
                    parentCell.setCellValue(Camelcase.toCamelCaseWithSpaces(label));
                    sheet.addMergedRegion(new CellRangeAddress(0, 0, startColIdx, colIdx - 1));
                    headerMapping.put(originalKey, startColIdx); // Map original key to column index
                }
            } else if (value instanceof Map) {
                @SuppressWarnings("unchecked")
                Map<String, Object> nestedMap = (Map<String, Object>) value;
                int startColIdx = colIdx;
                for (Map.Entry<String, Object> nestedEntry : nestedMap.entrySet()) {
                    String nestedKey = nestedEntry.getKey();
                    if (removeColumns.contains(originalKey + "." + nestedKey)) {
                        continue; // Skip nested columns
                    }
                    Cell childCell = childHeaderRow.createCell(colIdx++);
                    childCell.setCellValue(Camelcase.toCamelCaseWithSpaces(nestedKey));
                }
                Cell parentCell = parentHeaderRow.createCell(startColIdx);
                parentCell.setCellValue(Camelcase.toCamelCaseWithSpaces(label));
                sheet.addMergedRegion(new CellRangeAddress(0, 0, startColIdx, colIdx - 1));
                headerMapping.put(originalKey, startColIdx); // Map original key to column index
            } else {
                Cell cell = parentHeaderRow.createCell(colIdx);
                cell.setCellValue(Camelcase.toCamelCaseWithSpaces(label));
                sheet.addMergedRegion(new CellRangeAddress(0, 1, colIdx, colIdx));
                headerMapping.put(originalKey, colIdx); // Map original key to column index
                colIdx++;
            }
        }
        return headerMapping;
    }

    private void createDataRows(Sheet sheet, List<Map<String, Object>> jsonData, Map<String, Integer> headerMapping, Set<String> removeColumns, Workbook workbook, Map<String, String> amountColumns) {
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
                    } else if ("RECON_STATUS".equals(key)) {
                        // Handle RECON_STATUS field using ReconStatus class
                        ReconStatus.handleReconStatus(row, colIdx, (List<Object>) value);
                    } else {
                        // Handle other fields
                        if (value instanceof List) {
                            @SuppressWarnings("unchecked")
                            List<Object> nestedList = (List<Object>) value;
                            for (Object listItem : nestedList) {
                                if (listItem instanceof Map) {
                                    colIdx = fillRowWithMap(row, (Map<String, Object>) listItem, colIdx, key, removeColumns, workbook);
                                }
                            }
                        } else if (value instanceof Map) {
                            @SuppressWarnings("unchecked")
                            Map<String, Object> nestedMap = (Map<String, Object>) value;
                            colIdx = fillRowWithMap(row, nestedMap, colIdx, key, removeColumns, workbook);
                        } else {
                            if (amountColumns.containsKey(key)) {
                                // Use AmountColumn class to handle currency formatting
                                AmountColumn.handleAmountColumn(row, colIdx, value, workbook, amountColumns.get(key));
                            } else {
                                Cell cell = row.createCell(colIdx);
                                if (value != null) {
                                    cell.setCellValue(value.toString());
                                } else {
                                    cell.setCellValue((String) null);
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



