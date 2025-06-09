package com.ind.word_style_controller.controller;

import com.ind.StyleModel;
import com.ind.word_style_controller.HelloApplication;
import com.ind.word_style_controller.service.FileWatcherService;
import com.ind.word_style_controller.service.StyleApplicatorService;
import com.ind.word_style_controller.service.StyleLoaderService;
import com.ind.word_style_controller.service.StyleTableController;
import com.ind.word_style_controller.service.RemoveStyleService;
import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ListView;
import javafx.scene.control.TableView;
import javafx.stage.FileChooser;
import javafx.stage.Modality;
import javafx.stage.Stage;

import java.io.File;
import java.io.IOException;

/**
 * 主控制器类，负责处理主界面的导航和菜单功能
 */
public class MainController {
    @FXML
    public ListView<String> menuList;
    @FXML
    public Button injectToDocx;
    @FXML
    public Button removeStyleButton;
    @FXML
    public Label styleCountLabel;
    @FXML
    public TableView<StyleModel> styleTable;
    
    // 服务类实例
    private StyleTableController styleTableController;
    private StyleLoaderService styleLoaderService;
    private FileWatcherService fileWatcherService;
    private StyleApplicatorService styleApplicatorService;
    private RemoveStyleService removeStyleService;
    
    /**
     * 初始化控制器
     */
    public void initialize() {
        // 初始化菜单列表
        menuList.getItems().addAll("Customize", "Import", "About");
        menuList.setOnMouseClicked(event -> {
            String selectedItem = menuList.getSelectionModel().getSelectedItem();
            if ("Customize".equals(selectedItem)) {
                switchToCustomizeForm();
            } else if ("Import".equals(selectedItem)) {
                switchToImportForm();
            } else if ("About".equals(selectedItem)) {
                showAboutDialog();
            }
        });
        
        // 初始化服务类
        styleTableController = new StyleTableController(styleTable);
        styleLoaderService = new StyleLoaderService();
        styleApplicatorService = new StyleApplicatorService();
        removeStyleService = new RemoveStyleService(styleLoaderService, styleTableController);
        
        // 初始化表格
        styleTableController.initializeTable();
        
        // 加载样式数据
        reloadStyles();
        
        // 初始化文件监视服务
        try {
            fileWatcherService = new FileWatcherService(this::reloadStyles);
            // 启动文件监视服务
            fileWatcherService.start();
        } catch (IOException e) {
            e.printStackTrace();
            System.err.println("Failed to initialize file watcher: " + e.getMessage());
        }
    }
    
    /**
     * 重新加载样式数据
     */
    public void reloadStyles() {
        try {
            // 静默加载样式数据，不显示任何弹窗
            styleTableController.setStyleData(styleLoaderService.loadStylesFromXml());
            // 确保UI更新在JavaFX应用线程中执行
            Platform.runLater(this::updateStyleCountLabel);
        } catch (Exception e) {
            // 捕获任何可能的异常，确保应用程序不会因为样式加载问题而崩溃
            System.out.println("Note: Unable to reload styles at this moment. Will retry automatically.");
            // 不打印堆栈跟踪，避免在正常操作（如删除样式）导致的文件变化时显示大量错误信息
        }
    }
    
    /**
     * 更新样式计数标签
     */
    private void updateStyleCountLabel() {
        if (styleCountLabel != null && styleTableController != null) {
            int styleCount = styleTableController.getStyleData().size();
            styleCountLabel.setText(styleCount + " 个样式");
        }
    }
    
    /**
     * 切换到自定义表单
     */
    private void switchToCustomizeForm() {
        try {
            // 加载 customize-form 的 FXML 文件
            FXMLLoader fxmlLoader = new FXMLLoader(HelloApplication.class.getResource("/com/ind/word_style_controller/Customize-form.fxml"));
            Stage stage = new Stage();
            // 设置新的场景
            stage.initModality(Modality.APPLICATION_MODAL);
            stage.setTitle("Customize Form");
            stage.setScene(new Scene(fxmlLoader.load(), 800, 600));
            stage.showAndWait();
            
            // 表单关闭后重新加载样式
            reloadStyles();
        } catch (IOException e) {
            e.printStackTrace();
            System.err.println("Failed to load customize_form.fxml");
        }
    }
    
    /**
     * 切换到导入表单
     */
    private void switchToImportForm() {
        try {
            // 加载 ImportStyle 的 FXML 文件
            FXMLLoader fxmlLoader = new FXMLLoader(HelloApplication.class.getResource("/com/ind/word_style_controller/ImportStyle.fxml"));
            Stage stage = new Stage();
            // 设置新的场景
            stage.initModality(Modality.APPLICATION_MODAL);
            stage.setTitle("Import Styles from DOCX");
            stage.setScene(new Scene(fxmlLoader.load(), 800, 600));
            stage.showAndWait();
            
            // 表单关闭后重新加载样式
            reloadStyles();
        } catch (IOException e) {
            e.printStackTrace();
            System.err.println("Failed to load ImportStyle.fxml");
        }
    }
    
    /**
     * 将样式注入到DOCX文件
     */
    @FXML
    public void injectToDocx() {
        // 打开文件选择对话框，让用户选择目标 .docx 文件
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Select Target DOCX File");
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Word Files", "*.docx"));
        File selectedFile = fileChooser.showOpenDialog(injectToDocx.getScene().getWindow());

        if (selectedFile != null) {
            // 获取选中的样式
            var selectedStyles = styleTable.getSelectionModel().getSelectedItems();
            // 应用样式到Word文档
            styleApplicatorService.applyStylesToWord(selectedFile.getAbsolutePath(), selectedStyles);
        }
    }
    
    /**
     * 删除选中的样式
     */
    @FXML
    public void removeSelectedStyle() {
        removeStyleService.removeSelectedStyle(
            styleTableController.getSelectedStyles(),
            styleTable,
            this::updateStyleCountLabel
        );
    }
    
    /**
     * 显示关于对话框
     */
    private void showAboutDialog() {
        Alert aboutAlert = new Alert(Alert.AlertType.INFORMATION);
        aboutAlert.setTitle("关于");
        aboutAlert.setHeaderText("Word样式控制器");
        
        String aboutContent = "程序名称：Word样式控制器\n" +
                             "版本：1.0\n" +
                             "作者：Aiden\n" +
                             "联系方式：aidenhong916@outlook.com\n" +
                             "功能：管理和应用Word文档样式\n\n" +
                             "© 2024 Aiden. All rights reserved.";
        
        aboutAlert.setContentText(aboutContent);
        aboutAlert.setResizable(true);
        aboutAlert.getDialogPane().setPrefWidth(400);
        aboutAlert.showAndWait();
    }
    
    /**
     * 清理资源
     */
    public void cleanup() {
        if (fileWatcherService != null) {
            fileWatcherService.stop();
        }
    }
}