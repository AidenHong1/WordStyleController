package com.ind.word_style_controller.controller;

import com.ind.StyleModel;
import com.ind.word_style_controller.HelloApplication;
import com.ind.word_style_controller.service.FileWatcherService;
import com.ind.word_style_controller.service.StyleApplicatorService;
import com.ind.word_style_controller.service.StyleLoaderService;
import com.ind.word_style_controller.service.StyleTableController;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.control.Button;
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
        styleApplicatorService = new StyleApplicatorService();
        
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
        } catch (Exception e) {
            // 捕获任何可能的异常，确保应用程序不会因为样式加载问题而崩溃
            System.out.println("Note: Unable to reload styles at this moment. Will retry automatically.");
            // 不打印堆栈跟踪，避免在正常操作（如删除样式）导致的文件变化时显示大量错误信息
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
        // 获取所有选中的样式
        javafx.collections.ObservableList<StyleModel> selectedStyles = styleTableController.getSelectedStyles();

        if (selectedStyles == null || selectedStyles.isEmpty()) {
            // 如果没有选中任何样式，显示警告
            javafx.scene.control.Alert alert = new javafx.scene.control.Alert(javafx.scene.control.Alert.AlertType.WARNING);
            alert.setTitle("警告");
            alert.setHeaderText(null);
            alert.setContentText("请先选择要删除的样式！");
            alert.showAndWait();
            return;
        }

        // 构建确认消息
        String confirmMessage;
        if (selectedStyles.size() == 1) {
            confirmMessage = "确定要删除样式 \"" + selectedStyles.get(0).getName() + "\" 吗？此操作不可撤销。";
        } else {
            confirmMessage = "确定要删除选中的 " + selectedStyles.size() + " 个样式吗？此操作不可撤销。";
        }

        // 确认删除
        javafx.scene.control.Alert confirmAlert = new javafx.scene.control.Alert(javafx.scene.control.Alert.AlertType.CONFIRMATION);
        confirmAlert.setTitle("确认删除");
        confirmAlert.setHeaderText(null);
        confirmAlert.setContentText(confirmMessage);

        // 显示确认对话框并等待用户响应
        java.util.Optional<javafx.scene.control.ButtonType> result = confirmAlert.showAndWait();

        // 如果用户确认删除
        if (result.isPresent() && result.get() == javafx.scene.control.ButtonType.OK) {
            // 收集所有样式ID
            java.util.List<String> styleIds = new java.util.ArrayList<>();
            for (StyleModel style : selectedStyles) {
                styleIds.add(style.getId());
            }

            // 从XML文件中批量删除样式
            int xmlRemovedCount = styleLoaderService.removeStylesByIds(styleIds);

            // 从表格中批量删除样式
            int tableRemovedCount = styleTableController.removeStylesByIds(styleIds);

            // 删除后清空表格选中，防止死循环
            styleTable.getSelectionModel().clearSelection();

            System.out.print("Removed " + tableRemovedCount + " styles from the table, " + xmlRemovedCount + " styles from XML.");
            // 显示操作结果
            if (tableRemovedCount > 0) {
                javafx.scene.control.Alert successAlert = new javafx.scene.control.Alert(javafx.scene.control.Alert.AlertType.INFORMATION);
                successAlert.setTitle("成功");
                successAlert.setHeaderText(null);

                if (xmlRemovedCount == tableRemovedCount) {
                    successAlert.setContentText("已成功删除 " + xmlRemovedCount + " 个样式！");
                } else if (xmlRemovedCount > 0) {
                    successAlert.setContentText("已删除 " + tableRemovedCount + " 个样式，但其中只有 " + xmlRemovedCount + " 个样式从XML文件中删除成功。");
                } else {
                    successAlert.setContentText("已从表格中删除 " + tableRemovedCount + " 个样式，但无法从XML文件中删除。请检查文件权限或格式。");
                }

                successAlert.showAndWait();
            } else {
                javafx.scene.control.Alert errorAlert = new javafx.scene.control.Alert(javafx.scene.control.Alert.AlertType.ERROR);
                errorAlert.setTitle("错误");
                errorAlert.setHeaderText(null);
                errorAlert.setContentText("删除样式时出现错误！无法从表格中删除样式。");
                errorAlert.showAndWait();
            }
        }
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