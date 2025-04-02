package com.example.converter;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.util.List;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import java.io.FileOutputStream;
import javafx.stage.DirectoryChooser;
import javafx.stage.Stage;

public class ExcelToWordConverter {

    public static File convertToWord(List<File> files, List<String> expectedHeaders, Stage stage) {
        DirectoryChooser directoryChooser = new DirectoryChooser();
        directoryChooser.setTitle("Выберите папку для сохранения");
        File outputDirectory = directoryChooser.showDialog(stage);

        if (outputDirectory == null) {
            Controller.showAlert("Папка для сохранения не выбрана!");
            return null;
        }

        File outputFile = new File(outputDirectory, "output.docx");

        try (XWPFDocument document = new XWPFDocument()) {
            XWPFTable table = document.createTable();
            if (table.getRows().size() > 0) {
                table.removeRow(0); // Удаляем дефолтную строку
            }
            createHeaderRow(table, expectedHeaders); // Создаем заголовок

            for (File file : files) {
                try (FileInputStream fis = new FileInputStream(file);
                     Workbook workbook = new XSSFWorkbook(fis)) {

                    Sheet sheet = workbook.getSheetAt(0);
                    if (!validateColumns(sheet, expectedHeaders)) {
                        Controller.showAlert("Недопустимый заголовок в: " + file.getName());
                        continue;
                    }

                    // Пропускаем заголовок файла (первую строку)
                    for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                        Row row = sheet.getRow(i);
                        if (row == null) continue;

                        XWPFTableRow tableRow = table.createRow();
                        for (int j = 0; j < expectedHeaders.size(); j++) {
                            Cell cell = row.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                            String value = cellToString(cell);

                            if (tableRow.getCell(j) == null) {
                                tableRow.addNewTableCell();
                            }
                            tableRow.getCell(j).setText(value);
                        }
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                    Controller.showAlert("Ошибка при чтении файла: " + file.getName() + " - " + e.getMessage());
                }
            }

            try (FileOutputStream out = new FileOutputStream(outputFile)) {
                document.write(out);
            }
        } catch (Exception e) {
            e.printStackTrace();
            Controller.showAlert("Ошибка при конвертации в Word: " + e.getMessage());
        }
        return outputFile;
    }

    private static void createHeaderRow(XWPFTable table, List<String> headers) {
        XWPFTableRow headerRow = table.createRow();
        for (int i = 0; i < headers.size(); i++) {
            if (headerRow.getCell(i) == null) {
                headerRow.addNewTableCell();
            }
            headerRow.getCell(i).setText(headers.get(i));
        }
    }

    private static String cellToString(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    double value = cell.getNumericCellValue();
                    if (value == (long) value) {
                        return String.format("%d", (long) value);
                    } else {
                        return String.valueOf(value);
                    }
                }
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            default: return "";
        }
    }

    private static boolean validateColumns(Sheet sheet, List<String> expectedHeaders) {
        Row headerRow = sheet.getRow(0);
        if (headerRow == null) return false;

        for (int i = 0; i < expectedHeaders.size(); i++) {
            Cell cell = headerRow.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            String actualHeader = cellToString(cell).trim();
            String expectedHeader = expectedHeaders.get(i).trim();
            if (!actualHeader.equals(expectedHeader)) {
                return false;
            }
        }
        return true;
    }
}