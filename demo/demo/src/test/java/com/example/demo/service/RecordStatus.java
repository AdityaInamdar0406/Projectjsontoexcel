
package com.example.demo.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.util.List;
import java.util.Map;

public class RecordStatus {

    public static void handleReconStatus(Row row, int colIdx, List<Object> nestedList) {
        Cell titleCell = row.createCell(colIdx); // Assume colIdx is for title
        Cell colorCell = row.createCell(colIdx + 1); // Assume colIdx + 1 is for color
        Cell percentageCell = row.createCell(colIdx + 2); // Assume colIdx + 2 is for percentage

        StringBuilder concatenatedTitle = new StringBuilder();
        StringBuilder concatenatedColor = new StringBuilder();
        StringBuilder concatenatedPercentage = new StringBuilder();

        for (Object listItem : nestedList) {
            if (listItem instanceof Map) {
                Map<String, Object> mapItem = (Map<String, Object>) listItem;
                if (concatenatedTitle.length() > 0) {
                    concatenatedTitle.append(", ");
                    concatenatedColor.append(", ");
                    concatenatedPercentage.append(", ");
                }
                concatenatedTitle.append(mapItem.get("title")); // add title column
                concatenatedColor.append(mapItem.get("color")); // add color column
                concatenatedPercentage.append(mapItem.get("percentage")); // add percentage column
            }
        }

        // Set the cell values
        titleCell.setCellValue(concatenatedTitle.toString());
        colorCell.setCellValue(concatenatedColor.toString());
        percentageCell.setCellValue(concatenatedPercentage.toString());
    }
}
