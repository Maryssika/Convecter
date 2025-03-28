module com.example.converter {
    requires javafx.controls;
    requires javafx.fxml;
    requires org.apache.poi.ooxml;
    requires java.desktop;


    opens com.example.converter to javafx.fxml;
    exports com.example.converter;
}