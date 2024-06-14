package com.iconcile.excel_generator.controller;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import com.iconcile.excel_generator.service.JsonToExcelService;

import java.io.ByteArrayInputStream;
import java.io.IOException;

@RestController
@RequestMapping("/api")
public class JsonToExcelController {

    @Autowired
    private JsonToExcelService jsonToExcelService;

    @PostMapping(value = "/convert", consumes = MediaType.APPLICATION_JSON_VALUE)
    public ResponseEntity<byte[]> convertJsonToExcel(@RequestBody String jsonString) throws IOException {
        // Call the service method with the JSON string.
        
        ByteArrayInputStream byteArrayInputStream = jsonToExcelService.convertJsonToExcel(jsonString);

        // Set the response headers
        HttpHeaders headers = new HttpHeaders();
        headers.add("Content-Disposition", "attachment; filename=DATA1.xlsx");
        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);

        // Return the Excel file as a byte array in the response
        return ResponseEntity
                .ok()
                .headers(headers)
                .body(byteArrayInputStream.readAllBytes());
    }
}