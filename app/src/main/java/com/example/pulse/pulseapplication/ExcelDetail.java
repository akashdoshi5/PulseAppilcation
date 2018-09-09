package com.example.pulse.pulseapplication;

public class ExcelDetail {

    private String question;
    private String sheetName;

    public String getQuestion() {
        return question;
    }

    public void setQuestion(String question) {
        this.question = question;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    @Override
    public String toString() {
        return question;
    }

    @Override
    public boolean equals(Object obj) {
        if (obj instanceof ExcelDetail) {
            ExcelDetail c = (ExcelDetail) obj;
            if (c.getSheetName().equals(sheetName) && c.getQuestion() == question) return true;
        }

        return false;
    }
}
