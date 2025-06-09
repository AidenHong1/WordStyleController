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
        FXMLLoader fxmlLoader = new FXMLLoader(HelloApplication.class.getResource("/com/ind/word_style_controller/hello-view.fxml"));
        Scene scene = new Scene(fxmlLoader.load(), 1200, 800);
        
        // 设置窗口标题
        stage.setTitle("Word Style Controller - 样式管理工具");
        
        // 设置窗口大小和限制
        stage.setScene(scene);
        stage.setMinWidth(1000);  // 设置最小宽度
        stage.setMinHeight(700);  // 设置最小高度
        stage.setMaximized(false); // 默认不最大化
        stage.setResizable(true);  // 允许调整窗口大小
        
        // 居中显示窗口
        stage.centerOnScreen();
        
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