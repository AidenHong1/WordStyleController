package com.ind.word_style_controller;

import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.ColumnConstraints;
import javafx.scene.layout.Priority;
import javafx.geometry.Pos;
import javafx.geometry.HPos;
import javafx.scene.paint.Color;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class CustomizeController{
    @FXML
    private VBox stylesContainer;

    private List<Map<String, Object>> styleEntries = new ArrayList<>();
    
    // 样式ID计数器，从100开始
    private int styleIdCounter = 100;

    @FXML
    private void handleAddStyle() {
        HBox form = createStyleForm();
        stylesContainer.getChildren().add(form);
    }

    @FXML
    private void handleAppendStyles() {
        // 检查是否添加了任何样式
        if (styleEntries.isEmpty()) {
            showAlert("警告", "您没有添加任何样式！请先添加样式再进行操作。", Alert.AlertType.WARNING);
            return; // 取消函数进程
        }
        try {
            File xmlFile = new File(getClass().getResource("/styles.xml").toURI());
            DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
            DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
            Document doc = dBuilder.parse(xmlFile);
            doc.getDocumentElement().normalize();

            Element root = doc.getDocumentElement();

            for (Map<String, Object> entry : styleEntries) {
                // 创建新的 style 元素
                Element style = doc.createElement("style");

                // 获取并设置 id 属性
                TextField idField = (TextField) entry.get("idField");
                style.setAttribute("id", idField.getText());

                // 获取并设置 name 子元素
                TextField nameField = (TextField) entry.get("nameField");
                if (nameField.getText().isEmpty()) {
                    showAlert("警告", "请为每个样式提供一个名称！", Alert.AlertType.WARNING);
                    return; // 取消函数进程
                }
                addElement(doc, style, "name", nameField.getText());

                // 获取并设置 type 子元素
                ComboBox<String> typeCombo = (ComboBox<String>) entry.get("typeCombo");
                addElement(doc, style, "type", typeCombo.getValue());

                // 获取并设置 color 子元素
                ColorPicker colorPicker = (ColorPicker) entry.get("colorPicker");
                addElement(doc, style, "color", colorToHex(colorPicker.getValue()));

                // 获取并设置 fontSize 子元素
                TextField fontSizeField = (TextField) entry.get("fontSizeField");
                addElement(doc, style, "fontSize", fontSizeField.getText());

                // 将新的 style 元素追加到 root 元素中
                root.appendChild(style);
            }

            // 保存修改后的 XML 文件
            Transformer transformer = TransformerFactory.newInstance().newTransformer();
            transformer.setOutputProperty(OutputKeys.INDENT, "yes");
            transformer.transform(new DOMSource(doc), new StreamResult(xmlFile));
            showAlert("成功", "样式已追加到 XML 文件!", Alert.AlertType.INFORMATION);
        } catch (Exception e) {
            showAlert("错误", "追加失败: " + e.getMessage(), Alert.AlertType.ERROR);
        }
    }
    private HBox createStyleForm() {
        // 创建一个网格布局的面板
        GridPane gridPane = new GridPane();
        gridPane.setHgap(10);
        gridPane.setVgap(10);
        gridPane.setPadding(new javafx.geometry.Insets(5));
        gridPane.setStyle("-fx-background-color: #f8f8f8; -fx-border-color: #ddd; -fx-border-radius: 5;");
        
        // 设置列宽约束
        ColumnConstraints labelColumn = new ColumnConstraints();
        labelColumn.setPrefWidth(80);
        labelColumn.setHalignment(javafx.geometry.HPos.RIGHT);
        
        ColumnConstraints fieldColumn = new ColumnConstraints();
        fieldColumn.setPrefWidth(150);
        fieldColumn.setHgrow(javafx.scene.layout.Priority.ALWAYS);
        
        gridPane.getColumnConstraints().addAll(labelColumn, fieldColumn, labelColumn, fieldColumn);
        
        // 创建并配置控件
        TextField idField = new TextField();
        idField.setText(String.valueOf(styleIdCounter));
        idField.setEditable(false);
        idField.setStyle("-fx-opacity: 0.7; -fx-background-color: #f0f0f0;");
        idField.setPrefWidth(80);
        
        TextField nameField = new TextField();
        nameField.setPromptText("输入样式名称");
        nameField.setPrefWidth(150);
        
        ComboBox<String> typeCombo = new ComboBox<>();
        typeCombo.getItems().addAll("段落", "字符");
        typeCombo.setPrefWidth(150);
        
        ColorPicker colorPicker = new ColorPicker(Color.BLACK);
        colorPicker.setPrefWidth(150);
        
        TextField fontSizeField = new TextField("12");
        fontSizeField.setPromptText("字体大小");
        fontSizeField.setPrefWidth(80);
        
        // 初始化默认值
        typeCombo.getSelectionModel().selectFirst();
        
        // 添加到网格
        Label idLabel = new Label("样式ID:");
        idLabel.setStyle("-fx-font-weight: bold;");
        gridPane.add(idLabel, 0, 0);
        gridPane.add(idField, 1, 0);
        
        Label nameLabel = new Label("样式名称:");
        nameLabel.setStyle("-fx-font-weight: bold;");
        gridPane.add(nameLabel, 2, 0);
        gridPane.add(nameField, 3, 0);
        
        Label typeLabel = new Label("样式类型:");
        typeLabel.setStyle("-fx-font-weight: bold;");
        gridPane.add(typeLabel, 0, 1);
        gridPane.add(typeCombo, 1, 1);
        
        Label colorLabel = new Label("颜色:");
        colorLabel.setStyle("-fx-font-weight: bold;");
        gridPane.add(colorLabel, 2, 1);
        gridPane.add(colorPicker, 3, 1);
        
        Label fontSizeLabel = new Label("字体大小:");
        fontSizeLabel.setStyle("-fx-font-weight: bold;");
        gridPane.add(fontSizeLabel, 0, 2);
        gridPane.add(fontSizeField, 1, 2);
        
        // 创建一个包含网格和删除按钮的水平布局
        HBox hbox = new HBox(10);
        hbox.setAlignment(javafx.geometry.Pos.CENTER_LEFT);
        
        // 添加删除按钮
        Button deleteButton = new Button("删除");
        deleteButton.setStyle("-fx-background-color: #f44336; -fx-text-fill: white;");
        deleteButton.setOnAction(e -> {
            // 从容器和数据列表中移除
            stylesContainer.getChildren().remove(hbox);
            styleEntries.removeIf(entry -> entry.get("idField") == idField);
        });
        
        hbox.getChildren().addAll(gridPane, deleteButton);
        HBox.setHgrow(gridPane, javafx.scene.layout.Priority.ALWAYS);
        
        // 保存数据引用
        Map<String, Object> entry = new HashMap<>();
        entry.put("idField", idField);
        entry.put("nameField", nameField);
        entry.put("typeCombo", typeCombo);
        entry.put("colorPicker", colorPicker);
        entry.put("fontSizeField", fontSizeField);
        styleEntries.add(entry);
        
        // ID计数器自增
        styleIdCounter++;
        
        return hbox;
    }

    private String colorToHex(Color color) {
        if (color == null) {
            return "#000000"; // 返回默认黑色
        }
        try {
            int red = (int) (color.getRed() * 255);
            int green = (int) (color.getGreen() * 255);
            int blue = (int) (color.getBlue() * 255);
            return String.format("#%02X%02X%02X", red, green, blue);
        } catch (Exception e) {
            showAlert("错误", "颜色转换失败: " + e.getMessage(), Alert.AlertType.ERROR);
            return "#000000"; // 返回默认黑色
        }
    }

    private void addElement(Document doc, Element parent, String tagName, String value) {
        Element element = doc.createElement(tagName);
        element.setTextContent(value);
        parent.appendChild(element);
    }

    private void showAlert(String title, String message, Alert.AlertType type) {
        Alert alert = new Alert(type);
        alert.setTitle(title);
        alert.setHeaderText(null);
        alert.setContentText(message);
        alert.showAndWait();
    }
}