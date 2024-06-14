package com.iconcile.excel_generator.Model;

public class FormattingRule {
    private String colour;
    private String column;
    private Integer row;
    private String condition;

    public FormattingRule(String colour, String column, Integer row, String condition) {
        this.colour = colour;
        this.column = column;
        this.row = row;
        this.condition = condition;
    }

    public String getColour() {
        return colour;
    }

    public String getColumn() {
        return column;
    }

    public Integer getRow() {
        return row;
    }

    public String getCondition() {
        return condition;
    }

    public boolean matches(String column, int row, Object value) {
        boolean columnMatches = this.column == null || this.column.equals(column);
        boolean rowMatches = this.row == null || this.row == row;
        boolean conditionMatches = this.condition == null || value != null && value.toString().contains(this.condition);
        return columnMatches && rowMatches && conditionMatches;
    }
}

