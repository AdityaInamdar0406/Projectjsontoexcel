package com.example.demo;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.web.client.RestTemplateBuilder;
import org.springframework.context.annotation.Bean;
import org.springframework.web.client.RestTemplate;

@SpringBootApplication
public class DemoApplication {

	public static void main(String[] args) {
		SpringApplication.run(DemoApplication.class, args);

		// Send a POST request to the /convert endpoint with sample JSON data
        RestTemplate restTemplate = new RestTemplate();
        String url = "http://localhost:8080/convert"; // Update with your server URL
        List<Map<String, Object>> jsonData = createSampleJsonData();
        byte[] excelFile = restTemplate.postForObject(url, jsonData, byte[].class);
        // Now you can save or process the excelFile byte array as needed
	}
 @Bean
    public RestTemplate restTemplate(RestTemplateBuilder builder) {
        return builder.build();
    }

    private static List<Map<String, Object>> createSampleJsonData() {
        List<Map<String, Object>> jsonData = new ArrayList<>();

        Map<String, Object> data1 = new HashMap<>();
        data1.put("Name", "John");
        data1.put("Age", 30);
        data1.put("Email", "john@example.com");
        jsonData.add(data1);

        Map<String, Object> data2 = new HashMap<>();
        data2.put("Name", "Alice");
        data2.put("Age", 25);
        data2.put("Email", "alice@example.com");
        jsonData.add(data2);

        return jsonData;
    }
}



