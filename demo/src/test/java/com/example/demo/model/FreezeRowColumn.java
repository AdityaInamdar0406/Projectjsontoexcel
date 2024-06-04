package com.example.demo.model;

import org.apache.poi.ss.usermodel.Sheet;

public class FreezeRowColumn {
    public static void freezeRowsAndColumns(Sheet sheet, int freezeRows, int freezeColumns) {
        sheet.createFreezePane(freezeColumns, freezeRows);
    }
}