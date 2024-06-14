package com.iconcile.excel_generator.Model;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.util.List;
import java.util.Map;
import java.util.Set;

public class ReconStatus {

    public static void handleReconStatus(Row row, int colIdx, List<Object> nestedList) {
        // Check if the nested list is not empty
        if (nestedList != null && !nestedList.isEmpty() && nestedList.get(0) instanceof Map) {
            Map<String, Object> firstItem = (Map<String, Object>) nestedList.get(0);

            // Get the keys from the first item.
            Set<String> keys = firstItem.keySet();

            // Iterate over the keys to fill the cells
            int currentColIdx = colIdx;
            for (String key : keys) {
                Cell cell = row.createCell(currentColIdx++);
                StringBuilder concatenatedValues = new StringBuilder();

                // Iterate over the items in the nested list
                for (Object item : nestedList) {
                    if (item instanceof Map) {
                        Map<String, Object> mapItem = (Map<String, Object>) item;

                        // Check if the map contains the key
                        if (mapItem.containsKey(key)) {
                            // Append the value to the concatenated string
                            concatenatedValues.append(mapItem.get(key)).append(" , ");
                        } else {
                            // Append 'N/A' if the key is missing
                            concatenatedValues.append("N/A, ");
                        }
                    }
                }

                // Remove the trailing comma and space
                if (concatenatedValues.length() > 2) {
                    concatenatedValues.setLength(concatenatedValues.length() - 2);
                }

                // Set the cell value
                cell.setCellValue(concatenatedValues.toString());
            }
        }
    }
}