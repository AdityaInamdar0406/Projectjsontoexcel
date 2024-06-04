
package com.example.demo.model;

import java.util.HashMap;
import java.util.Map;

public class Hashmap {
    private Map<String, Object> data = new HashMap<>();

    public Hashmap() {
    }

    public Hashmap(Map<String, Object> data) {
        this.data = data;
    }

    public Map<String, Object> getData() {
        return data;
    }

    public void setData(Map<String, Object> data) {
        this.data = data;
    }

    public Object getValue(String key) {
        return data.get(key);
    }

    public void setValue(String key, Object value) {
        data.put(key, value);
    }

    public Map<String, Object> toHashMap(HashMap<String, Integer> columnMapping) {
        Map<String, Object> hashMap = new HashMap<>();
        for (Map.Entry<String, Integer> entry : columnMapping.entrySet()) {
            String fieldName = entry.getKey();
            if (data.containsKey(fieldName)) {
                hashMap.put(fieldName, data.get(fieldName));
            }
        }
        return hashMap;
    }
}