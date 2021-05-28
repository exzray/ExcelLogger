package com.exzray;

public class Main {

    public static void main(String[] args) {
        ExcelLogger logger = ExcelLogger.newInstance("test.xlsx");
        logger.writeLog("abc123", "just simple error");
    }
}
