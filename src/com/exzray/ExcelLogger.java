package com.exzray;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public final class ExcelLogger {

    private String string_path;
    private final List<String> header_names = new ArrayList<>();


    public static ExcelLogger newInstance(String file_path) {
        ExcelLogger logger = new ExcelLogger();
        logger.string_path = file_path;

        return logger;
    }

    private ExcelLogger() {
        header_names.add("uid");
        header_names.add("description");
        header_names.add("timestamp");
    }

    public void writeLog(String uid, String description) {
        synchronized (this) {
            checkExcel();

            File file = new File(string_path);

            try {
                FileInputStream inputStream = new FileInputStream(file);
                XSSFWorkbook workbook = new XSSFWorkbook(inputStream);

                XSSFSheet sheet = workbook.getSheetAt(0);

                int lastRowIndex = sheet.getLastRowNum();
                int targetRowIndex = lastRowIndex + 1;

                XSSFRow row = sheet.createRow(targetRowIndex);
                XSSFCell cell_uid = row.createCell(0);
                XSSFCell cell_description = row.createCell(1);
                XSSFCell cell_timestamp = row.createCell(2);

                cell_uid.setCellValue(uid);
                cell_description.setCellValue(description);
                cell_timestamp.setCellValue(getCurrentTimestamp());

                OutputStream outputStream = new FileOutputStream(file);
                workbook.write(outputStream);

                workbook.close();
                outputStream.close();

            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private String getCurrentTimestamp() {
        return new SimpleDateFormat("yyyy-MM-dd hh:mm:ss").format(new Date());
    }

    private void checkExcel() {
        File file = new File(string_path);

        if (!file.exists()) {
            try {

                boolean status = file.createNewFile();

                if (status) {
                    createNewExcel(file);
                    System.out.println("start a fresh...");
                }

            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private void createNewExcel(File file) throws IOException {
        OutputStream outputStream = new FileOutputStream(file);

        XSSFWorkbook workbook = new XSSFWorkbook();
        createSheet(workbook);
        workbook.write(outputStream);

        workbook.close();
        outputStream.close();
    }

    private void createSheet(XSSFWorkbook workbook) {
        XSSFSheet sheet = workbook.createSheet();
        XSSFRow row = sheet.createRow(0);

        CellStyle style = workbook.createCellStyle();
        style.setWrapText(true);
        style.setAlignment(HorizontalAlignment.CENTER);

        for (int index = 0; index < header_names.size(); index++) {
            sheet.setColumnWidth(index, 6000);

            XSSFCell cell = row.createCell(index);
            cell.setCellValue(header_names.get(index));
            cell.setCellStyle(style);
        }
    }
}
