package com.example.converter;

import java.io.File;
import java.io.FileOutputStream;
import java.util.List;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import javafx.stage.DirectoryChooser;
import javafx.stage.Stage;

public class ExcelToExcelConverter {

    public static void convertToExcel(List<File> files, List<String> expectedHeaders, Stage stage) {
        DirectoryChooser directoryChooser = new DirectoryChooser();
        directoryChooser.setTitle("Выберите папку для сохранения");
        File outputDirectory = directoryChooser.showDialog(stage);

        if (outputDirectory == null) {
            Controller.showAlert("Папка для сохранения не выбрана!");
            return;
        }

        File outputFile = new File(outputDirectory, "output.xlsx");

        try (Workbook newWorkbook = new XSSFWorkbook()) {
            Sheet newSheet = newWorkbook.createSheet("Combined");
            int rowIndex = 0;
            boolean isFirstFile = true;

            for (File file : files) {
                try (FileInputStream fis = new FileInputStream(file);
                     Workbook workbook = new XSSFWorkbook(fis)) {

                    Sheet sheet = workbook.getSheetAt(0);
                    if (!validateColumns(sheet, expectedHeaders)) {
                        Controller.showAlert("Недопустимый заголовок в: " + file.getName());
                        continue;
                    }

                    for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                        Row row = sheet.getRow(i);
                        if (row == null) continue;

                        if (!isFirstFile && i == 0) continue;

                        Row newRow = newSheet.createRow(rowIndex++);
                        for (int j = 0; j < expectedHeaders.size(); j++) {
                            Cell cell = row.getCell(j);
                            if (cell != null) {
                                switch (cell.getCellType()) {
                                    case STRING:
                                        newRow.createCell(j).setCellValue(cell.getStringCellValue());
                                        break;
                                    case NUMERIC:
                                        if (DateUtil.isCellDateFormatted(cell)) {
                                            newRow.createCell(j).setCellValue(cell.getDateCellValue());
                                        } else {
                                            double value = cell.getNumericCellValue();
                                            if (value == (long) value) {
                                                newRow.createCell(j).setCellValue((long) value);
                                            } else {
                                                newRow.createCell(j).setCellValue(value);
                                            }
                                        }
                                        break;
                                    case BOOLEAN:
                                        newRow.createCell(j).setCellValue(cell.getBooleanCellValue());
                                        break;
                                    default:
                                        newRow.createCell(j).setCellValue("");
                                }
                            } else {
                                newRow.createCell(j).setCellValue("");
                            }
                        }
                    }
                    isFirstFile = false;
                }
            }

            try (FileOutputStream out = new FileOutputStream(outputFile)) {
                newWorkbook.write(out);
            }
        } catch (Exception e) {
            e.printStackTrace();
            Controller.showAlert("Ошибка при конвертации в Excel: " + e.getMessage());
        }
    }

    private static boolean validateColumns(Sheet sheet, List<String> expectedHeaders) {
        Row headerRow = sheet.getRow(0);
        if (headerRow == null) return false;

        for (int i = 0; i < expectedHeaders.size(); i++) {
            Cell cell = headerRow.getCell(i);
            if (cell == null || !cell.toString().equals(expectedHeaders.get(i))) {
                return false;
            }
        }
        return true;
    }
}