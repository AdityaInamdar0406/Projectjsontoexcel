package com.example.demo.service;

import com.fasterxml.jackson.databind.JsonNode;

import java.util.HashMap;
import java.util.Map;

public class LabelName {

    private Map<String, String> labelNameMapping;

    public LabelName() {
        this.labelNameMapping = new HashMap<>();
    }

    public void extractLabelNameMapping(JsonNode rootNode) {
        if (rootNode.has("LABEL_NAME")) {
            JsonNode labelNameNode = rootNode.get("LABEL_NAME");
            for (JsonNode labelNode : labelNameNode) {
                String column = labelNode.get("COLUMN").asText();
                String label = labelNode.get("LABEL").asText();
                labelNameMapping.put(column, label);
            }
        }
    }

    public String getLabelForColumn(String column) {
        return labelNameMapping.get(column);
    }
}