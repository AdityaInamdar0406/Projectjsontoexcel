package com.iconcile.excel_generator.Model;

import com.fasterxml.jackson.databind.JsonNode;
import java.util.HashMap;
import java.util.Map;

public class LabelName {
    private static Map<String, String> labelNameMapping = new HashMap<>();

    public static void extractLabelNameMapping(JsonNode rootNode) {
        if (rootNode.has("formatting") && rootNode.get("formatting").has("LABEL_NAME")) {
            JsonNode labelNameNode = rootNode.get("formatting").get("LABEL_NAME");
            for (JsonNode labelNode : labelNameNode) {
                String column = labelNode.get("COLUMN").asText();
                String label = labelNode.get("LABEL").asText();
                labelNameMapping.put(column, label);
            }
        }
    }

    public static String getLabelForColumn(String column) {
        return labelNameMapping.getOrDefault(column, column); // Return the label or the original column name if no label is found
    }
}