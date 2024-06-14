package com.example.demo.service;

import com.fasterxml.jackson.databind.JsonNode;

import java.util.HashSet;
import java.util.Set;

public class RemoveColumn {

    public static Set<String> extractRemovedColumns(JsonNode rootNode) {
        Set<String> removeColumns = new HashSet<>();
        if (rootNode.has("formatting") && rootNode.get("formatting").has("REMOVE_COLUMNS")) {
            for (JsonNode column : rootNode.get("formatting").get("REMOVE_COLUMNS")) {
                removeColumns.add(column.asText());
            }
        }
        return removeColumns;
    }
}
