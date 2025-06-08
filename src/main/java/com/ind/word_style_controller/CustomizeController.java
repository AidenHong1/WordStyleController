package com.ind.word_style_controller;

import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.ColumnConstraints;
import javafx.scene.paint.Color;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

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
            // 直接使用项目中的styles.xml文件路径
            String stylesPath = "src/main/resources/styles.xml";
            File xmlFile = new File(stylesPath);
            
            // 如果文件不存在，创建一个基本的XML结构
            if (!xmlFile.exists()) {
                // 确保目录存在
                xmlFile.getParentFile().mkdirs();
                
                // 创建基本的XML文件
                DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
                DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
                Document newDoc = dBuilder.newDocument();
                Element rootElement = newDoc.createElement("styles");
                newDoc.appendChild(rootElement);
                
                // 保存基本XML文件
                Transformer transformer = TransformerFactory.newInstance().newTransformer();
                transformer.setOutputProperty(OutputKeys.INDENT, "yes");
                transformer.transform(new DOMSource(newDoc), new StreamResult(xmlFile));
            }
            
            DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
            DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
            Document doc = dBuilder.parse(xmlFile);
            doc.getDocumentElement().normalize();

            Element root = doc.getDocumentElement();

            for (Map<String, Object> entry : styleEntries) {
                // 创建新的 xml-fragment 元素
                Element xmlFragment = doc.createElement("xml-fragment");
                
                ComboBox<String> typeCombo = (ComboBox<String>) entry.get("typeCombo");
                if (typeCombo.getValue() != null && !typeCombo.getValue().isEmpty()) {
                    String type = typeCombo.getValue();
                    String typeValue = "paragraph";
                    switch (type) {
                        case "段落":
                            typeValue = "paragraph";
                            break;
                        case "字符":
                            typeValue = "character";
                            break;
                    }
                    xmlFragment.setAttribute("w:type", typeValue);
                }

                TextField styleNameField = (TextField) entry.get("nameField");
                String styleName = styleNameField.getText().replaceAll("\\s+", ""); // 移除空格作为styleId
                
                // 添加命名空间属性
                xmlFragment.setAttribute("xmlns:mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                xmlFragment.setAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                xmlFragment.setAttribute("xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                xmlFragment.setAttribute("xmlns:w14", "http://schemas.microsoft.com/office/word/2010/wordml");
                xmlFragment.setAttribute("xmlns:w15", "http://schemas.microsoft.com/office/word/2012/wordml");
                xmlFragment.setAttribute("xmlns:w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
                xmlFragment.setAttribute("xmlns:w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
                xmlFragment.setAttribute("xmlns:w16", "http://schemas.microsoft.com/office/word/2018/wordml");
                xmlFragment.setAttribute("xmlns:w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
                xmlFragment.setAttribute("xmlns:w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
                
                // 创建w:name元素
                Element wName = doc.createElement("w:name");
                wName.setAttribute("w:val", styleName);
                xmlFragment.appendChild(wName);
                
                // 创建w:name元素
                Element wId = doc.createElement("w:link");
                wId.setAttribute("w:val", styleName);
                xmlFragment.appendChild(wId);

                // 创建w:qFormat元素
                Element wQFormat = doc.createElement("w:qFormat");
                xmlFragment.appendChild(wQFormat);
                
                // 创建w:pPr元素（段落属性）
                Element wPPr = doc.createElement("w:pPr");
                
                // 添加段落间距
                TextField paragraphSpacingField = (TextField) entry.get("paragraphSpacingField");
                TextField paragraphBeforeSpacingField = (TextField) entry.get("paragraphBeforeSpacingField");
                TextField lineSpacingField = (TextField) entry.get("lineSpacingField");
                if (!paragraphSpacingField.getText().isEmpty() || !paragraphBeforeSpacingField.getText().isEmpty() || !lineSpacingField.getText().isEmpty()) {
                    Element wSpacing = doc.createElement("w:spacing");
                    if (!paragraphSpacingField.getText().isEmpty()) {
                        double spacing = Double.parseDouble(paragraphSpacingField.getText()) * 240; // 转换为twips
                        wSpacing.setAttribute("w:after", String.valueOf((int)spacing));
                    }
                    if (!paragraphBeforeSpacingField.getText().isEmpty()) {
                        double beforeSpacing = Double.parseDouble(paragraphBeforeSpacingField.getText()) * 240; // 转换为twips
                        wSpacing.setAttribute("w:before", String.valueOf((int)beforeSpacing));
                    }
                    if (!lineSpacingField.getText().isEmpty()) {
                        double lineSpacing = Double.parseDouble(lineSpacingField.getText()) * 240; // 转换为twips
                        wSpacing.setAttribute("w:line", String.valueOf((int)lineSpacing));
                        wSpacing.setAttribute("w:lineRule", "auto");
                    }
                    wPPr.appendChild(wSpacing);
                }
                
                // 添加对齐方式
                ComboBox<String> alignmentCombo = (ComboBox<String>) entry.get("alignmentCombo");
                if (alignmentCombo.getValue() != null && !alignmentCombo.getValue().isEmpty()) {
                    Element wJc = doc.createElement("w:jc");
                    String alignment = alignmentCombo.getValue();
                    String alignValue = "left";
                    switch (alignment) {
                        case "居中":
                            alignValue = "center";
                            break;
                        case "右对齐":
                            alignValue = "right";
                            break;
                        case "两端对齐":
                            alignValue = "both";
                            break;
                        default:
                            alignValue = "left";
                            break;
                    }
                    wJc.setAttribute("w:val", alignValue);
                    wPPr.appendChild(wJc);
                }
                
                xmlFragment.appendChild(wPPr);
                
                // 创建w:rPr元素（字符属性）
                Element wRPr = doc.createElement("w:rPr");
                
                // 添加字体
                ComboBox<String> fontCombo = (ComboBox<String>) entry.get("fontCombo");
                ComboBox<String> eastAsiaFontCombo = (ComboBox<String>) entry.get("eastAsiaFontCombo");
                if ((fontCombo.getValue() != null && !fontCombo.getValue().isEmpty()) || 
                    (eastAsiaFontCombo.getValue() != null && !eastAsiaFontCombo.getValue().isEmpty())) {
                    Element wRFonts = doc.createElement("w:rFonts");
                    if (fontCombo.getValue() != null && !fontCombo.getValue().isEmpty()) {
                        wRFonts.setAttribute("w:ascii", fontCombo.getValue());
                        wRFonts.setAttribute("w:hAnsi", fontCombo.getValue());
                    }
                    if (eastAsiaFontCombo.getValue() != null && !eastAsiaFontCombo.getValue().isEmpty()) {
                        wRFonts.setAttribute("w:eastAsia", eastAsiaFontCombo.getValue());
                    }
                    wRPr.appendChild(wRFonts);
                }
                
                // 添加粗体
                CheckBox boldCheckBox = (CheckBox) entry.get("boldCheckBox");
                if (boldCheckBox.isSelected()) {
                    Element wB = doc.createElement("w:b");
                    wRPr.appendChild(wB);
                }
                
                // 添加字体大小
                TextField fontSizeField = (TextField) entry.get("fontSizeField");
                if (!fontSizeField.getText().isEmpty()) {
                    int fontSize = Integer.parseInt(fontSizeField.getText()) * 2; // Word中字体大小是半点
                    Element wSz = doc.createElement("w:sz");
                    wSz.setAttribute("w:val", String.valueOf(fontSize));
                    wRPr.appendChild(wSz);
                    
                    Element wSzCs = doc.createElement("w:szCs");
                    wSzCs.setAttribute("w:val", String.valueOf(fontSize));
                    wRPr.appendChild(wSzCs);
                }
                
                // 添加颜色
                ColorPicker colorPicker = (ColorPicker) entry.get("colorPicker");
                String colorHex = colorToHex(colorPicker.getValue());
                if (!colorHex.equals("#000000")) { // 如果不是默认黑色
                    Element wColor = doc.createElement("w:color");
                    wColor.setAttribute("w:val", colorHex.substring(1)); // 移除#号
                    wRPr.appendChild(wColor);
                }
                
                xmlFragment.appendChild(wRPr);
                
                // 将新的 xml-fragment 元素追加到 root 元素中
                root.appendChild(xmlFragment);
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

        TextField styleNameField = new TextField();
        styleNameField.setPromptText("样式名称");
        styleNameField.setPrefWidth(80);

        ComboBox<String> fontCombo = new ComboBox<>();
        fontCombo.getItems().addAll(javafx.scene.text.Font.getFamilies());
        fontCombo.setPromptText("英文字体");
        fontCombo.setPrefWidth(150);
        // Replace nameField with fontComboBox in the layout

        
        ComboBox<String> typeCombo = new ComboBox<>();
        typeCombo.getItems().addAll("段落", "字符");
        typeCombo.setPrefWidth(150);
        
        ColorPicker colorPicker = new ColorPicker(Color.BLACK);
        colorPicker.setPrefWidth(150);
        
        TextField fontSizeField = new TextField("12");
        fontSizeField.setPromptText("字体大小");
        fontSizeField.setPrefWidth(80);
        
        TextField paragraphSpacingField = new TextField("1.0");
        paragraphSpacingField.setPromptText("段后间距");
        paragraphSpacingField.setPrefWidth(80);
        
        TextField paragraphBeforeSpacingField = new TextField("1.0");
        paragraphBeforeSpacingField.setPromptText("段前间距");
        paragraphBeforeSpacingField.setPrefWidth(80);
        
        TextField lineSpacingField = new TextField("1.0");
        lineSpacingField.setPromptText("行间距");
        lineSpacingField.setPrefWidth(80);
        
        ComboBox<String> alignmentCombo = new ComboBox<>();
        alignmentCombo.getItems().addAll("左对齐", "居中", "右对齐", "两端对齐");
        alignmentCombo.setPromptText("对齐方式");
        alignmentCombo.setPrefWidth(150);
        alignmentCombo.getSelectionModel().selectFirst();
        
        CheckBox boldCheckBox = new CheckBox("粗体");
        boldCheckBox.setPrefWidth(80);
        
        ComboBox<String> eastAsiaFontCombo = new ComboBox<>();
        eastAsiaFontCombo.getItems().addAll(
            "宋体", "黑体", "楷体", "仿宋", "微软雅黑", "SimSun", "SimHei", "KaiTi", "FangSong",
            "Microsoft YaHei", "Microsoft JhengHei", "PingFang SC", "Hiragino Sans GB", "STHeiti",
            "STSong", "STKaiti", "STFangsong", "华文宋体", "华文黑体", "华文楷体", "华文仿宋",
            "方正黑体", "方正宋体", "方正楷体", "思源黑体", "思源宋体"
        );
        eastAsiaFontCombo.setPromptText("中文字体");
        eastAsiaFontCombo.setPrefWidth(150);
        
        // 初始化默认值
        typeCombo.getSelectionModel().selectFirst();


        // 添加到网格

        Label styleNameLabel = new Label("样式名称");
        styleNameLabel.setStyle("-fx-font-weight: bold;");
        gridPane.add(styleNameLabel, 0, 0);
        gridPane.add(styleNameField, 1, 0);

        Label nameLabel = new Label("样式字体:");
        nameLabel.setStyle("-fx-font-weight: bold;");
        gridPane.add(nameLabel, 2, 0);
        gridPane.add(fontCombo, 3, 0);
        
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
        
        Label paragraphSpacingLabel = new Label("段后间距:");
        paragraphSpacingLabel.setStyle("-fx-font-weight: bold;");
        gridPane.add(paragraphSpacingLabel, 2, 2);
        gridPane.add(paragraphSpacingField, 3, 2);
        
        Label paragraphBeforeSpacingLabel = new Label("段前间距:");
        paragraphBeforeSpacingLabel.setStyle("-fx-font-weight: bold;");
        gridPane.add(paragraphBeforeSpacingLabel, 0, 3);
        gridPane.add(paragraphBeforeSpacingField, 1, 3);
        
        Label lineSpacingLabel = new Label("行间距:");
        lineSpacingLabel.setStyle("-fx-font-weight: bold;");
        gridPane.add(lineSpacingLabel, 2, 3);
        gridPane.add(lineSpacingField, 3, 3);
        
        Label alignmentLabel = new Label("对齐方式:");
        alignmentLabel.setStyle("-fx-font-weight: bold;");
        gridPane.add(alignmentLabel, 0, 4);
        gridPane.add(alignmentCombo, 1, 4);
        
        Label boldLabel = new Label("字体样式:");
        boldLabel.setStyle("-fx-font-weight: bold;");
        gridPane.add(boldLabel, 2, 4);
        gridPane.add(boldCheckBox, 3, 4);
        
        Label eastAsiaFontLabel = new Label("中文字体:");
        eastAsiaFontLabel.setStyle("-fx-font-weight: bold;");
        gridPane.add(eastAsiaFontLabel, 0, 5);
        gridPane.add(eastAsiaFontCombo, 1, 5);
        // 创建一个包含网格和删除按钮的水平布局
        HBox hbox = new HBox(10);
        hbox.setAlignment(javafx.geometry.Pos.CENTER_LEFT);
        
        // 添加删除按钮
        Button deleteButton = new Button("删除");
        deleteButton.setStyle("-fx-background-color: #f44336; -fx-text-fill: white;");
        deleteButton.setOnAction(e -> {
            // 从容器和数据列表中移除
            stylesContainer.getChildren().remove(hbox);
            styleEntries.removeIf(entry -> entry.get("nameField") == styleNameField);
        });
        
        hbox.getChildren().addAll(gridPane, deleteButton);
        HBox.setHgrow(gridPane, javafx.scene.layout.Priority.ALWAYS);
        
        // 保存数据引用
        Map<String, Object> entry = new HashMap<>();
        entry.put("nameField", styleNameField);
        entry.put("fontCombo", fontCombo);
        entry.put("typeCombo", typeCombo);
        entry.put("colorPicker", colorPicker);
        entry.put("fontSizeField", fontSizeField);
        entry.put("paragraphSpacingField", paragraphSpacingField);
        entry.put("paragraphBeforeSpacingField", paragraphBeforeSpacingField);
        entry.put("lineSpacingField", lineSpacingField);
        entry.put("alignmentCombo", alignmentCombo);
        entry.put("boldCheckBox", boldCheckBox);
        entry.put("eastAsiaFontCombo", eastAsiaFontCombo);
        styleEntries.add(entry);
        
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

    private void showAlert(String title, String message, Alert.AlertType type) {
        Alert alert = new Alert(type);
        alert.setTitle(title);
        alert.setHeaderText(null);
        alert.setContentText(message);
        alert.showAndWait();
    }
}
