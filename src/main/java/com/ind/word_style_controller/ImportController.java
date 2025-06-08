package com.ind.word_style_controller;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.List;

public class ImportController {
    
    @FXML
    private TextField docxFilePathField;
    
    @FXML
    private ListView<String> styleListView;
    
    @FXML
    private Button browseButton;
    
    @FXML
    private Button exportSelectedButton;
    
    @FXML
    private Button exportAllButton;
    
    @FXML
    private Label statusLabel;
    
    private ObservableList<String> styleNames = FXCollections.observableArrayList();
    private XWPFDocument currentDocument;
    
    public void initialize() {
        styleListView.setItems(styleNames);
        styleListView.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        statusLabel.setText("请选择DOCX文件并加载样式");
    }
    
    @FXML
    private void browseDocxFile() {
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("选择DOCX文件");
        fileChooser.getExtensionFilters().add(
            new FileChooser.ExtensionFilter("Word Documents", "*.docx")
        );
        
        Stage stage = (Stage) browseButton.getScene().getWindow();
        File selectedFile = fileChooser.showOpenDialog(stage);
        
        if (selectedFile != null) {
            docxFilePathField.setText(selectedFile.getAbsolutePath());
            statusLabel.setText("文件已选择，点击'加载样式'按钮");
        }
        loadStyles();
    }
    
    @FXML
    private void loadStyles() {
        String filePath = docxFilePathField.getText();
        if (filePath == null || filePath.trim().isEmpty()) {
            showAlert("错误", "请先选择DOCX文件");
            return;
        }
        
        try {
            FileInputStream fis = new FileInputStream(filePath);
            currentDocument = new XWPFDocument(fis);
            
            styleNames.clear();
            XWPFStyles styles = currentDocument.getStyles();
            
            // 获取样式数量并遍历
            int numberOfStyles = styles.getNumberOfStyles();
            for (int i = 0; i < numberOfStyles; i++) {
                // 通过反射或其他方式获取样式，这里使用文档中所有段落的样式
                // 由于getStyleList()方法不存在，我们需要通过其他方式获取样式
            }
            
            // 另一种方法：遍历文档中的段落和运行来收集使用的样式
            List<XWPFParagraph> paragraphs = currentDocument.getParagraphs();
            java.util.Set<String> usedStyleIds = new java.util.HashSet<>();
            
            for (XWPFParagraph paragraph : paragraphs) {
                String styleId = paragraph.getStyleID();
                if (styleId != null && !styleId.isEmpty()) {
                    usedStyleIds.add(styleId);
                }
            }
            
            // 添加找到的样式到列表
            for (String styleId : usedStyleIds) {
                XWPFStyle style = styles.getStyle(styleId);
                if (style != null) {
                    String displayName = style.getName() != null ? style.getName() : styleId;
                    styleNames.add(displayName + " (" + styleId + ")");
                }
            }
            
            statusLabel.setText("已加载 " + styleNames.size() + " 个样式");
            fis.close();
            
        } catch (IOException e) {
            showAlert("错误", "无法读取DOCX文件: " + e.getMessage());
            e.printStackTrace();
        }
    }
    
    @FXML
    private void exportSelectedStyles() {
        ObservableList<String> selectedItems = styleListView.getSelectionModel().getSelectedItems();
        if (selectedItems.isEmpty()) {
            showAlert("提示", "请选择要导出的样式");
            return;
        }
        
        exportStyles(selectedItems);
    }
    
    @FXML
    private void exportAllStyles() {
        if (styleNames.isEmpty()) {
            showAlert("提示", "没有可导出的样式");
            return;
        }
        
        exportStyles(styleNames);
    }
    
    private void exportStyles(ObservableList<String> stylesToExport) {
        if (currentDocument == null) {
            showAlert("错误", "请先加载DOCX文件");
            return;
        }
        
        try {
            // 读取现有的styles.xml文件内容
            String outputPath = "src/main/resources/styles.xml";
            File outputFile = new File(outputPath);
            
            // 确保目录存在
            outputFile.getParentFile().mkdirs();
            
            // 读取现有的XML内容
            StringBuilder existingContent = new StringBuilder();
            if (outputFile.exists()) {
                java.util.Scanner scanner = new java.util.Scanner(outputFile, "UTF-8");
                while (scanner.hasNextLine()) {
                    existingContent.append(scanner.nextLine()).append("\n");
                }
                scanner.close();
            } else {
                // 如果文件不存在，创建基本结构
                existingContent.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?>\n");
                existingContent.append("<styles>\n");
                existingContent.append("</styles>\n");
            }
            
            // 解析现有内容，找到</styles>标签的位置
            int endTagIndex = existingContent.lastIndexOf("</styles>");
            if (endTagIndex == -1) {
                showAlert("错误", "styles.xml文件格式不正确");
                return;
            }
            
            // 准备新的样式内容
            StringBuilder newStylesContent = new StringBuilder();
            XWPFStyles styles = currentDocument.getStyles();
            int addedStylesCount = 0;
            
            for (String selectedStyleDisplay : stylesToExport) {
                // 从显示名称中提取样式ID
                String styleId = extractStyleId(selectedStyleDisplay);
                
                // 检查样式是否已存在
                if (existingContent.indexOf("w:styleId=\"" + styleId + "\"") != -1) {
                    // 样式已存在，跳过
                    continue;
                }
                
                // 直接通过样式ID获取样式
                XWPFStyle style = styles.getStyle(styleId);
                if (style != null) {
                    CTStyle ctStyle = style.getCTStyle();
                    if (ctStyle != null) {
                        newStylesContent.append(ctStyle.xmlText() + "\n");
                        addedStylesCount++;
                    }
                }
            }
            
            // 如果没有新样式被添加，提示用户
            if (addedStylesCount == 0) {
                showAlert("提示", "没有新的样式被添加，所有选中的样式已存在");
                return;
            }
            
            // 将新样式插入到</styles>标签之前
            existingContent.insert(endTagIndex, newStylesContent.toString());
            
            // 写入更新后的内容到文件
            FileWriter writer = new FileWriter(outputFile);
            writer.write(existingContent.toString());
            writer.close();
            
            statusLabel.setText("已追加 " + addedStylesCount + " 个新样式到 " + outputPath);
            showAlert("成功", "已成功追加 " + addedStylesCount + " 个新样式到 " + outputPath);
            
        } catch (IOException e) {
            showAlert("错误", "导出样式时发生错误: " + e.getMessage());
            e.printStackTrace();
        }
    }
    
    private String extractStyleId(String displayName) {
        // 从 "样式名称 (样式ID)" 格式中提取样式ID
        int startIndex = displayName.lastIndexOf("(");
        int endIndex = displayName.lastIndexOf(")");
        if (startIndex != -1 && endIndex != -1 && startIndex < endIndex) {
            return displayName.substring(startIndex + 1, endIndex);
        }
        return displayName;
    }
    
    private void showAlert(String title, String message) {
        Alert alert = new Alert(Alert.AlertType.INFORMATION);
        alert.setTitle(title);
        alert.setHeaderText(null);
        alert.setContentText(message);
        alert.showAndWait();
    }
}
