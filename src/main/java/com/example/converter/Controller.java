package com.example.converter;

import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.stage.FileChooser;
import javafx.stage.DirectoryChooser;
import javafx.stage.Stage;
import java.io.File;
import java.io.IOException;
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
    @FXML
    private ListView<File> filesListView;

    private List<File> selectedFiles = new ArrayList<>();
    private List<String> expectedHeaders = new ArrayList<>();
    private Stage stage;

    @FXML
    public void initialize() {
        filesListView.getSelectionModel().setSelectionMode(SelectionMode.SINGLE);
    }

    public void setStage(Stage stage) {
        this.stage = stage;
    }

    @FXML
    private void handleSelectFilesButtonAction() {
        FileChooser fileChooser = new FileChooser();
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel Files", "*.xlsx"));
        List<File> files = fileChooser.showOpenMultipleDialog(stage);
        if (files != null) {
            selectedFiles.addAll(files);
            updateFilesListView();
            updateStatus("Было выбрано " + files.size() + " файла(ов)");
        }
    }

    @FXML
    private void handleSelectFolderButtonAction() {
        DirectoryChooser directoryChooser = new DirectoryChooser();
        File selectedDirectory = directoryChooser.showDialog(stage);
        if (selectedDirectory != null) {
            File[] files = selectedDirectory.listFiles((dir, name) -> name.endsWith(".xlsx"));
            if (files != null) {
                selectedFiles.addAll(Arrays.asList(files));
                updateFilesListView();
                updateStatus("Было выбрано " + files.length + " файлов из папки");
            }
        }
    }

    @FXML
    private void handleRemoveFileButtonAction() {
        File selectedFile = filesListView.getSelectionModel().getSelectedItem();
        if (selectedFile != null) {
            selectedFiles.remove(selectedFile);
            updateFilesListView();
            updateStatus("Файл удален: " + selectedFile.getName());
        } else {
            updateStatus("Файл не выбран для удаления");
        }
    }

    @FXML
    private void handleReplaceFileButtonAction() {
        File selectedFile = filesListView.getSelectionModel().getSelectedItem();
        if (selectedFile != null) {
            FileChooser fileChooser = new FileChooser();
            fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel Files", "*.xlsx"));
            File newFile = fileChooser.showOpenDialog(stage);
            if (newFile != null) {
                int index = selectedFiles.indexOf(selectedFile);
                selectedFiles.set(index, newFile);
                updateFilesListView();
                updateStatus("Файл заменен: " + newFile.getName());
            }
        } else {
            updateStatus("Файл не выбран для замены");
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
        if (selectedFiles.isEmpty()) {
            updateStatus("Файлы не выбраны!");
            return;
        }
        if (expectedHeaders.isEmpty()) {
            showAlert("Сначала установите правила проверки таблиц!");
            return;
        }

        File outputFile = ExcelToWordConverter.convertToWord(selectedFiles, expectedHeaders, stage);
        updateStatus("Конвертация в Word выполнена");
        showConversionSuccessDialog(outputFile);
    }

    @FXML
    private void handleConvertToExcelButtonAction() {
        if (selectedFiles.isEmpty()) {
            updateStatus("Файлы не выбраны!");
            return;
        }
        if (expectedHeaders.isEmpty()) {
            showAlert("Сначала установите правила проверки таблиц!");
            return;
        }

        File outputFile = ExcelToExcelConverter.convertToExcel(selectedFiles, expectedHeaders, stage);
        updateStatus("Конвертация в Excel выполнена");
        showConversionSuccessDialog(outputFile);
    }

    private void updateFilesListView() {
        filesListView.getItems().setAll(selectedFiles);
    }

    private void updateStatus(String message) {
        statusLabel.setText(message);
    }

    private void showConversionSuccessDialog(File outputFile) {
        if (outputFile == null) return;

        Platform.runLater(() -> {
            Alert alert = new Alert(Alert.AlertType.CONFIRMATION);
            alert.setTitle("Конвертация завершена");
            alert.setHeaderText("Файл успешно создан: " + outputFile.getName());
            alert.setContentText("Хотите открыть файл?");

            ButtonType openButton = new ButtonType("Открыть");
            ButtonType cancelButton = new ButtonType("Отмена", ButtonBar.ButtonData.CANCEL_CLOSE);

            alert.getButtonTypes().setAll(openButton, cancelButton);

            alert.showAndWait().ifPresent(type -> {
                if (type == openButton) {
                    try {
                        java.awt.Desktop.getDesktop().open(outputFile);
                    } catch (IOException e) {
                        showAlert("Не удалось открыть файл: " + e.getMessage());
                    }
                }
            });
        });
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