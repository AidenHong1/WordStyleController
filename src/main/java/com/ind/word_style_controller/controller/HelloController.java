package com.ind.word_style_controller.controller;

import com.ind.StyleModel;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ListView;
import javafx.scene.control.TableView;

/**
 * HelloController类，作为原始控制器的代理，将功能委托给新的控制器类
 * 这个类保留是为了保持与现有FXML文件的兼容性
 */
public class HelloController {
    @FXML
    public ListView<String> menuList;
    @FXML
    public Button injectToDocx;
    @FXML
    public Button removeStyleButton;
    @FXML
    public Label styleCountLabel;
    @FXML
    private TableView<StyleModel> styleTable;
    
    // 主控制器实例
    private MainController mainController;
    
    /**
     * 初始化控制器
     */
    public void initialize() {
        // 创建主控制器并初始化
        mainController = new MainController();
        
        // 设置主控制器的UI组件引用
        mainController.menuList = this.menuList;
        mainController.injectToDocx = this.injectToDocx;
        mainController.removeStyleButton = this.removeStyleButton;
        mainController.styleCountLabel = this.styleCountLabel;
        mainController.styleTable = this.styleTable;
        
        // 初始化主控制器
        mainController.initialize();
    }
    
    /**
     * 将样式注入到DOCX文件
     */
    @FXML
    public void injectToDocx() {
        mainController.injectToDocx();
    }
    
    /**
     * 删除选中的样式
     */
    @FXML
    public void removeSelectedStyle() {
        mainController.removeSelectedStyle();
    }
}
