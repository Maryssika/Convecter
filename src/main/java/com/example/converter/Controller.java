package com.example.converter;

import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.stage.FileChooser;
import javafx.stage.DirectoryChooser;
import javafx.stage.Stage;
import java.io.File;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class Controller {

    @FXML
    private Button selectFilesButton;
    @FXML
    private Button selectFolderButton;
    @FXML
    private Label statusLabel;
    @FXML
    private TextField headersField;

    private List<File> selectedFiles;
    private List<String> expectedHeaders = new ArrayList<>();
    private Stage stage;

    @FXML
    public void initialize() {
        // Инициализация (если требуется)
    }

    public void setStage(Stage stage) {
        this.stage = stage;
    }

    @FXML
    private void handleSelectFilesButtonAction() {
        FileChooser fileChooser = new FileChooser();
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel Files", "*.xlsx"));
        selectedFiles = fileChooser.showOpenMultipleDialog(null);
        updateStatus("Было выбрано " + (selectedFiles != null ? selectedFiles.size() : 0) + " файла");
    }

    @FXML
    private void handleSelectFolderButtonAction() {
        DirectoryChooser directoryChooser = new DirectoryChooser();
        File selectedDirectory = directoryChooser.showDialog(null);
        if (selectedDirectory != null) {
            selectedFiles = Arrays.asList(selectedDirectory.listFiles((dir, name) -> name.endsWith(".xlsx")));
            updateStatus("Было выбрано " + selectedFiles.size() + " файлов из папки");
        }
    }

    @FXML
    private void handleSetValidationButtonAction() {
        try {
            expectedHeaders = Arrays.asList(headersField.getText().split("\\s*,\\s*"));
            updateStatus("Установлена проверка заголовков: " + expectedHeaders);
        } catch (Exception e) {
            showAlert("Недопустимый формат! Пример: Header1, Header2");
            updateStatus("Ошибка настройки проверки");
        }
    }

    @FXML
    private void handleConvertToWordButtonAction() {
        if (selectedFiles == null || selectedFiles.isEmpty()) {
            updateStatus("Файл не выбран!");
            return;
        }
        if (expectedHeaders.isEmpty()) {
            showAlert("Сначала установите правила проверки таблиц!");
            return;
        }

        ExcelToWordConverter.convertToWord(selectedFiles, expectedHeaders, stage);
        updateStatus("Конвертация в Word выполнена");
    }

    @FXML
    private void handleConvertToExcelButtonAction() {
        if (selectedFiles == null || selectedFiles.isEmpty()) {
            updateStatus("Файл не выбран!");
            return;
        }
        if (expectedHeaders.isEmpty()) {
            showAlert("Сначала установите правила проверки таблиц!");
            return;
        }

        ExcelToExcelConverter.convertToExcel(selectedFiles, expectedHeaders, stage);
        updateStatus("Конвертация в Excel выполнена");
    }

    private void updateStatus(String message) {
        statusLabel.setText(message);
    }

    public static void showAlert(String message) {
        Platform.runLater(() -> {
            Alert alert = new Alert(Alert.AlertType.WARNING);
            alert.setTitle("Ошибка проверки");
            alert.setHeaderText(null);
            alert.setContentText(message);
            alert.showAndWait();
        });
    }
}