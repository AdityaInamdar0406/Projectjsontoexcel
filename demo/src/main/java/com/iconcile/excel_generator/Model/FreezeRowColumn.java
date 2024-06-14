package com.iconcile.excel_generator.Model;

import org.apache.poi.ss.usermodel.Sheet;

import com.fasterxml.jackson.databind.JsonNode;

public class FreezeRowColumn {
    public static void freezeRowsAndColumns(Sheet sheet, int freezeRows, int freezeColumns) {
        sheet.createFreezePane(freezeColumns, freezeRows);
    }

    public static void freezeRowsAndColumns(Sheet sheet, JsonNode rootNode) {
        // TODO Auto-generated method stub
        throw new UnsupportedOperationException("Unimplemented method 'freezeRowsAndColumns'");
    }
}