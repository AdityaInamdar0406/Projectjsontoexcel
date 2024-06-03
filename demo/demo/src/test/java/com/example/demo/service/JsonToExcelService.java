package com.example.demo.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import com.example.demo.model.Autosize;
import com.example.demo.model.Camelcase;
import com.example.demo.model.ExcelBackgroundFill;
import com.example.demo.model.ExcelBorder;
import com.example.demo.model.FreezeRowColumn;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

@Service
public class JsonToExcelService {

    public ByteArrayInputStream convertJsonToExcel(String jsonString) throws IOException {
        try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet("Data");

            // Parse JSON string into JsonNode
            ObjectMapper objectMapper = new ObjectMapper();
            JsonNode rootNode = objectMapper.readTree(jsonString);

            

            // Extract REMOVE_COLUMNS list
            Set<String> removeColumns = new HashSet<>();
            if (rootNode.has("formatting") && rootNode.get("formatting").has("REMOVE_COLUMNS")) {
                for (JsonNode column : rootNode.get("formatting").get("REMOVE_COLUMNS")) {
                    removeColumns.add(column.asText());
                }
            }

 
            List<Map<String, Object>> jsonData = objectMapper.convertValue(rootNode.get("data"), List.class);

            if (jsonData != null && !jsonData.isEmpty()) {
                // Generate header row
                Map<String, Integer> headerMapping = createHeaderRow(sheet, jsonData.get(0), removeColumns, workbook);

                // Generate data rows
                createDataRows(sheet, jsonData, headerMapping, removeColumns, workbook);

                // Autosize columns after filling data
                Autosize.autoSizeColumns(sheet);


            // Apply background fill
            ExcelBackgroundFill.applyBackgroundFill(sheet, rootNode);

                // Extract freeze information from JSON data
                int freezeRows = rootNode.has("freezeRows") ? rootNode.get("freezeRows").asInt() : 0;
                int freezeColumns = rootNode.has("freezeColumns") ? rootNode.get("freezeColumns").asInt() : 3;

                // Freeze rows and columns
                FreezeRowColumn.freezeRowsAndColumns(sheet, freezeRows, freezeColumns);

                // Add border to all cells
                ExcelBorder.addBordersToSheet(sheet);
            }

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        }
    }


    private Map<String, Integer> createHeaderRow(Sheet sheet, Map<String, Object> firstRecord, Set<String> removeColumns,Workbook workbook) {
        Row parentHeaderRow = sheet.createRow(0);
        Row childHeaderRow = sheet.createRow(1);
        int colIdx = 0;
        Map<String, Integer> headerMapping = new HashMap<>();

        for (Map.Entry<String, Object> entry : firstRecord.entrySet()) {
            String key = entry.getKey();
            Object value = entry.getValue();

            

            if (removeColumns.contains(key)) {
                continue; // Skip columns that are in the removeColumns set
            }

            CellStyle cellStyle = workbook.createCellStyle();
            
            if (value instanceof List) {
                @SuppressWarnings("unchecked")
                List<Object> nestedList = (List<Object>) value;
                if (!nestedList.isEmpty() && nestedList.get(0) instanceof Map) {
                    int startColIdx = colIdx;
                    for (Map.Entry<String, Object> nestedEntry : ((Map<String, Object>) nestedList.get(0)).entrySet()) {
                       
                        if (removeColumns.contains(key + "." + nestedEntry.getKey())) {
                            continue; // Skip nested columns
                        }
                        Cell childCell = childHeaderRow.createCell(colIdx++);
                        childCell.setCellValue(Camelcase.toCamelCaseWithSpaces(nestedEntry.getKey()));
                        childCell.setCellStyle(cellStyle); // Apply background fill style
                    }
                    Cell parentCell = parentHeaderRow.createCell(startColIdx);
                    parentCell.setCellValue(Camelcase.toCamelCaseWithSpaces(key));
                    parentCell.setCellStyle(cellStyle); // Apply background fill style
                    sheet.addMergedRegion(new CellRangeAddress(0, 0, startColIdx, colIdx - 1));
                    headerMapping.put(key, startColIdx);
                }
            } else if (value instanceof Map) {
                @SuppressWarnings("unchecked")
                Map<String, Object> nestedMap = (Map<String, Object>) value;
                int startColIdx = colIdx;
                for (Map.Entry<String, Object> nestedEntry : nestedMap.entrySet()) {
                    if (removeColumns.contains(key + "." + nestedEntry.getKey())) {
                        continue; // Skip nested columns
                    }
                    Cell childCell = childHeaderRow.createCell(colIdx++);
                    childCell.setCellValue(Camelcase.toCamelCaseWithSpaces(nestedEntry.getKey()));
                    childCell.setCellStyle(cellStyle); // Apply background fill style
                }
                Cell parentCell = parentHeaderRow.createCell(startColIdx);
                parentCell.setCellValue(Camelcase.toCamelCaseWithSpaces(key));
                parentCell.setCellStyle(cellStyle); // Apply background fill style
                sheet.addMergedRegion(new CellRangeAddress(0, 0, startColIdx, colIdx - 1));
                headerMapping.put(key, startColIdx);
            } else {
                Cell cell = parentHeaderRow.createCell(colIdx);
                cell.setCellValue(Camelcase.toCamelCaseWithSpaces(key));
                cell.setCellStyle(cellStyle); // Apply background fill style
                sheet.addMergedRegion(new CellRangeAddress(0, 1, colIdx, colIdx));
                headerMapping.put(key, colIdx);
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

                    CellStyle cellStyle = workbook.createCellStyle();
                 
                    if ("FIELD_ID".equals(key)) {
                        // Apply numbering format to the Field Id column
                        CellStyle style = sheet.getWorkbook().createCellStyle();
                        style.setDataFormat((short) 1); // Number format
                        Cell cell = row.createCell(colIdx);
                        if (value != null) {
                            cell.setCellValue(Double.parseDouble(value.toString())); // Assuming FIELD_ID is numeric
                        }
                        cell.setCellStyle(style);
                    } else {
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
                        } 
                        else {
                            if ("CHG_AMOUNT".equals(key)) {
                                // Update CHG_AMOUNT column to add "$" symbol to non-null values and format as number
                                if ("CHG_AMOUNT".equals(key)) {
                                    Cell cell = row.createCell(colIdx);
                                    if (value != null) {
                                        // Add "$" symbol only to non-null values
                                        cell.setCellValue("$" + value);
    
                                         // Set cell format as currency
                                         CellStyle currencyStyle = sheet.getWorkbook().createCellStyle();
                                         currencyStyle.setDataFormat((short) 7); // Currency format code
                                         cell.setCellStyle(currencyStyle);
                                            
                                    } else {
                                        // If value is null, keep the cell empty
                                        cell.setCellValue("");
                                    }
                                }
                            }
                        else {
                            Cell cell = row.createCell(colIdx);
                            setCellValue(cell, value);
                            cell.setCellStyle(cellStyle); // Apply background fill style
                        }
                    }
                }
            }
        }
    }
}

    private int fillRowWithMap(Row row, Map<String, Object> map, int colIdx, String parentKey, Set<String> removeColumns, Workbook workbook) {
        for (Map.Entry<String, Object> entry : map.entrySet()) {
            String fullKey = parentKey + "." + entry.getKey();
            if (removeColumns.contains(fullKey)) {
                continue; // Skip nested columns
            }
            Cell cell = row.createCell(colIdx++);
            setCellValue(cell, entry.getValue());

            CellStyle cellStyle = workbook.createCellStyle();
            cell.setCellStyle(cellStyle); // Apply background fill style
        }
        return colIdx;
    }

    private void setCellValue(Cell cell, Object value) {
        if (value == null) {
            cell.setCellValue("");
        } else if (value instanceof Number) {
            cell.setCellValue(((Number) value).doubleValue());
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else {
            cell.setCellValue(value.toString());
        }
    }
}
