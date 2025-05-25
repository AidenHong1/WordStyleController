package com.ind.word_style_controller;

import javafx.application.Application;
import javafx.application.Platform;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.stage.Stage;
import java.io.IOException;

public class HelloApplication extends Application {
    @Override
    public void start(Stage stage) throws IOException {
        // 打开hello-view.fxml
        FXMLLoader fxmlLoader = new FXMLLoader(HelloApplication.class.getResource("hello-view.fxml"));
        Scene scene = new Scene(fxmlLoader.load(), 800, 600);
        stage.setTitle("Word Style Controller");
        stage.setScene(scene);
        // 添加窗口关闭事件处理，当主窗口关闭时自动结束程序
        stage.setOnCloseRequest(event -> {
            Platform.exit(); // 结束JavaFX应用程序
            System.exit(0);  // 确保所有线程都被终止
        });
        stage.show();
    }
    public static void main(String[] args) {
        launch();
    }
}