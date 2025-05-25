module com.ind.word_style_controller {
    requires javafx.controls;
    requires javafx.fxml;
    requires org.apache.poi.ooxml;
    requires java.desktop;


    opens com.ind.word_style_controller to javafx.fxml;
    exports com.ind.word_style_controller;
}