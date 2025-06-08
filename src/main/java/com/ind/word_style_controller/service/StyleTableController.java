package com.ind.word_style_controller.service;

import com.ind.StyleModel;
import javafx.beans.property.SimpleStringProperty;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.geometry.Pos;
import javafx.scene.control.*;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.layout.HBox;
import javafx.scene.paint.Color;
import javafx.scene.shape.Rectangle;

/**
 * 样式表格控制器，负责处理样式表格的显示和配置
 */
public class StyleTableController {
    private final TableView<StyleModel> styleTable;
    private final ObservableList<StyleModel> styleData = FXCollections.observableArrayList();
    
    /**
     * 构造函数
     * @param styleTable 样式表格视图
     */
    public StyleTableController(TableView<StyleModel> styleTable) {
        this.styleTable = styleTable;
    }
    
    /**
     * 初始化表格
     */
    public void initializeTable() {
        // 获取表格列
        TableColumn<StyleModel, String> valColumn = (TableColumn<StyleModel, String>) styleTable.getColumns().get(0);
        TableColumn<StyleModel, String> fontColumn = (TableColumn<StyleModel, String>) styleTable.getColumns().get(1);
        TableColumn<StyleModel, String> typeColumn = (TableColumn<StyleModel, String>) styleTable.getColumns().get(2);
        TableColumn<StyleModel, String> colorColumn = (TableColumn<StyleModel, String>) styleTable.getColumns().get(3);
        TableColumn<StyleModel, String> fontSizeColumn = (TableColumn<StyleModel, String>) styleTable.getColumns().get(4);
        TableColumn<StyleModel, String> alignmentColumn = (TableColumn<StyleModel, String>) styleTable.getColumns().get(5);
        TableColumn<StyleModel, Double> paragraphBeforeSpacingColumn = (TableColumn<StyleModel, Double>) styleTable.getColumns().get(6);
        TableColumn<StyleModel, Double> paragraphSpacingColumn = (TableColumn<StyleModel, Double>) styleTable.getColumns().get(7);
        TableColumn<StyleModel, Double> lineSpacingColumn = (TableColumn<StyleModel, Double>) styleTable.getColumns().get(8);
        
        // 设置列的值工厂
        valColumn.setCellValueFactory(new PropertyValueFactory<>("name"));
        fontColumn.setCellValueFactory(new PropertyValueFactory<>("font"));
        typeColumn.setCellValueFactory(new PropertyValueFactory<>("type"));
        alignmentColumn.setCellValueFactory(new PropertyValueFactory<>("alignment"));
        colorColumn.setCellValueFactory(new PropertyValueFactory<>("color"));
        
        // 为字体大小列使用自定义的CellValueFactory处理类型转换
        fontSizeColumn.setCellValueFactory(cellData -> {
            int fontSize = cellData.getValue().getFontSize();
            return new SimpleStringProperty(String.valueOf(fontSize));
        });
        
        // 为字体大小列添加单位
        fontSizeColumn.setCellFactory(column -> new TableCell<StyleModel, String>() {
            @Override
            protected void updateItem(String fontSize, boolean empty) {
                super.updateItem(fontSize, empty);
                if (empty || fontSize == null) {
                    setText(null);
                } else {
                    setText(fontSize + " pt");
                }
            }
        });
        
        // 为段落前间距列使用自定义的CellValueFactory和CellFactory
        paragraphBeforeSpacingColumn.setCellValueFactory(new PropertyValueFactory<>("paragraphBeforeSpacing"));
        paragraphBeforeSpacingColumn.setCellFactory(column -> new TableCell<StyleModel, Double>() {
            @Override
            protected void updateItem(Double spacing, boolean empty) {
                super.updateItem(spacing, empty);
                if (empty || spacing == null) {
                    setText(null);
                } else {
                    setText(String.format("%.1f pt", spacing));
                }
            }
        });
        
        // 为段落后间距列使用自定义的CellValueFactory和CellFactory
        paragraphSpacingColumn.setCellValueFactory(new PropertyValueFactory<>("paragraphSpacing"));
        paragraphSpacingColumn.setCellFactory(column -> new TableCell<StyleModel, Double>() {
            @Override
            protected void updateItem(Double spacing, boolean empty) {
                super.updateItem(spacing, empty);
                if (empty || spacing == null) {
                    setText(null);
                } else {
                    setText(String.format("%.1f pt", spacing));
                }
            }
        });
        
        // 为行间距列使用自定义的CellValueFactory和CellFactory
        lineSpacingColumn.setCellValueFactory(new PropertyValueFactory<>("lineSpacing"));
        lineSpacingColumn.setCellFactory(column -> new TableCell<StyleModel, Double>() {
            @Override
            protected void updateItem(Double spacing, boolean empty) {
                super.updateItem(spacing, empty);
                if (empty || spacing == null) {
                    setText(null);
                } else {
                    setText(String.format("%.2f 倍", spacing));
                }
            }
        });
        
        // 为颜色列设置自定义的单元格工厂，显示颜色矩形和颜色代码
        colorColumn.setCellFactory(column -> new TableCell<StyleModel, String>() {
            private final Rectangle colorRect = new Rectangle(16, 16);
            private final HBox hbox = new HBox(5);
            private final Label colorLabel = new Label();
            
            {
                hbox.setAlignment(Pos.CENTER_LEFT);
                hbox.getChildren().addAll(colorRect, colorLabel);
            }
            
            @Override
            protected void updateItem(String color, boolean empty) {
                super.updateItem(color, empty);
                if (empty || color == null) {
                    setText(null);
                    setGraphic(null);
                } else {
                    try {
                        // 处理颜色格式，确保是有效的JavaFX颜色
                        String colorCode = color.startsWith("#") ? color : "#" + color;
                        Color fxColor = Color.web(colorCode);
                        colorRect.setFill(fxColor);
                        colorRect.setStroke(Color.BLACK);
                        colorLabel.setText(color);
                        setGraphic(hbox);
                    } catch (Exception e) {
                        // 如果颜色无效，只显示文本
                        setText(color);
                        setGraphic(null);
                    }
                }
            }
        });
        
        // 设置表格数据源
        styleTable.setItems(styleData);
        
        // 设置多选模式
        styleTable.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        
        // 设置表格列的相对宽度百分比
        styleTable.widthProperty().addListener((obs, oldVal, newVal) -> {
            double width = newVal.doubleValue();
            // 设置各列的相对宽度百分比
            valColumn.setPrefWidth(width * 0.15);          // 15%
            fontColumn.setPrefWidth(width * 0.15);         // 15%
            typeColumn.setPrefWidth(width * 0.10);         // 10%
            colorColumn.setPrefWidth(width * 0.10);        // 10%
            fontSizeColumn.setPrefWidth(width * 0.10);     // 10%
            alignmentColumn.setPrefWidth(width * 0.10);    // 10%
            paragraphBeforeSpacingColumn.setPrefWidth(width * 0.10); // 10%
            paragraphSpacingColumn.setPrefWidth(width * 0.10);      // 10%
            lineSpacingColumn.setPrefWidth(width * 0.10);           // 10%
        });
    }
    
    /**
     * 设置样式数据
     * @param newStyleData 新的样式数据
     */
    public void setStyleData(ObservableList<StyleModel> newStyleData) {
        styleData.clear();
        styleData.addAll(newStyleData);
    }
    
    /**
     * 获取样式数据
     * @return 样式数据列表
     */
    public ObservableList<StyleModel> getStyleData() {
        return styleData;
    }
    
    /**
     * 根据样式ID从表格中删除行
     * @param styleId 要删除的样式ID
     * @return 是否删除成功
     */
    public boolean removeStyleById(String styleId) {
        if (styleId == null || styleId.isEmpty()) {
            return false;
        }
        
        // 查找匹配ID的样式
        StyleModel styleToRemove = null;
        for (StyleModel style : styleData) {
            if (styleId.equals(style.getId())) {
                styleToRemove = style;
                break;
            }
        }
        
        // 如果找到了匹配的样式，从表格数据中删除
        if (styleToRemove != null) {
            return styleData.remove(styleToRemove);
        }
        
        return false;
    }
    
    /**
     * 批量删除多个样式
     * @param styles 要删除的样式列表
     * @return 成功删除的样式数量
     */
    public int removeStyles(java.util.List<StyleModel> styles) {
        if (styles == null || styles.isEmpty()) {
            return 0;
        }
        
        int successCount = 0;
        // 创建要删除的样式ID列表
        java.util.List<String> styleIdsToRemove = new java.util.ArrayList<>();
        for (StyleModel style : styles) {
            styleIdsToRemove.add(style.getId());
        }
        
        // 创建要删除的样式对象列表
        java.util.List<StyleModel> stylesToRemove = new java.util.ArrayList<>();
        for (StyleModel style : styleData) {
            if (styleIdsToRemove.contains(style.getId())) {
                stylesToRemove.add(style);
                successCount++;
            }
        }
        
        // 从数据集合中删除样式
        styleData.removeAll(stylesToRemove);
        
        return successCount;
    }
    
    /**
     * 获取当前选中的样式
     * @return 选中的样式，如果没有选中则返回null
     */
    public StyleModel getSelectedStyle() {
        return styleTable.getSelectionModel().getSelectedItem();
    }
    
    /**
     * 获取所有选中的样式
     * @return 选中的样式列表
     */
    public ObservableList<StyleModel> getSelectedStyles() {
        return styleTable.getSelectionModel().getSelectedItems();
    }
}