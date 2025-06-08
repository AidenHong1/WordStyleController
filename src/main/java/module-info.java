module com.ind.word_style_controller {
    requires javafx.controls;
    requires javafx.fxml;
    requires org.apache.poi.ooxml;
    requires java.desktop;
    requires org.apache.poi.ooxml.schemas;
    opens com.ind to javafx.base;
    opens com.ind.word_style_controller to javafx.fxml;
    exports com.ind.word_style_controller;
    exports com.ind;
}