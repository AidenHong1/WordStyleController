package com.ind.word_style_controller;

import com.ind.StyleModel;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.ListView;
import javafx.scene.control.TableView;
import javafx.stage.FileChooser;
import javafx.stage.Modality;
import javafx.stage.Stage;

import java.awt.Desktop;
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
    public TableView<StyleModel> styleTable;
    
    // 服务类实例
    private StyleTableController styleTableController;
    private StyleLoaderService styleLoaderService;
    private FileWatcherService fileWatcherService;
    private StyleApplicatorService styleApplicatorService;
    
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
            }
        });
        
        // 初始化服务类
        styleTableController = new StyleTableController(styleTable);
        styleLoaderService = new StyleLoaderService();
        fileWatcherService = new FileWatcherService(this::reloadStyles);
        styleApplicatorService = new StyleApplicatorService();
        
        // 初始化表格
        styleTableController.initializeTable();
        
        // 加载样式数据
        reloadStyles();
        
        // 启动文件监视服务
        fileWatcherService.startWatching();
    }
    
    /**
     * 重新加载样式数据
     */
    public void reloadStyles() {
        styleTableController.setStyleData(styleLoaderService.loadStylesFromXml());
    }
    
    /**
     * 切换到自定义表单
     */
    private void switchToCustomizeForm() {
        try {
            // 加载 customize-form 的 FXML 文件
            FXMLLoader fxmlLoader = new FXMLLoader(HelloApplication.class.getResource("Customize-form.fxml"));
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
            FXMLLoader fxmlLoader = new FXMLLoader(HelloApplication.class.getResource("ImportStyle.fxml"));
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
     * 清理资源
     */
    public void cleanup() {
        if (fileWatcherService != null) {
            fileWatcherService.stopWatching();
        }
    }
}