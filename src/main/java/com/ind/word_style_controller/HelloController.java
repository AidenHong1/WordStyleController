package com.ind.word_style_controller;

import com.ind.StyleModel;
import com.ind.word_style_controller.utils.CommonUtils;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.layout.HBox;
import javafx.scene.layout.StackPane;
import javafx.scene.paint.Color;
import javafx.scene.shape.Rectangle;
import javafx.stage.FileChooser;
import javafx.stage.Modality;
import javafx.stage.Stage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFont;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.awt.Desktop;
import java.io.*;
import java.math.BigInteger;
import java.nio.file.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.UUID;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.beans.property.SimpleStringProperty;

public class HelloController {
    @FXML
    public ListView<String> menuList;
    public Button injectToDocx;
    private ExecutorService executorService;
    // TableView 相关
    @FXML
    private TableView<StyleModel> styleTable;
    @FXML
    private TableColumn<StyleModel, String> valColumn;
    @FXML
    private TableColumn<StyleModel, String> fontColumn;
    @FXML
    private TableColumn<StyleModel, String> typeColumn;
    @FXML
    private TableColumn<StyleModel, String> colorColumn;
    @FXML
    private TableColumn<StyleModel, String> fontSizeColumn;
    @FXML
    private TableColumn<StyleModel, String> alignmentColumn;
    @FXML
    private TableColumn<StyleModel, Double> paragraphSpacingColumn;
    @FXML
    private TableColumn<StyleModel, Double> lineSpacingColumn;
    @FXML
    private TableColumn<StyleModel, Double> paragraphBeforeSpacingColumn;
    private final ObservableList<StyleModel> styleData = FXCollections.observableArrayList();

    public void initialize() {
        menuList.getItems().addAll("Customize", "Import", "About");
        menuList.setOnMouseClicked(event -> {
            String selectedItem = menuList.getSelectionModel().getSelectedItem();
            if ("Customize".equals(selectedItem)) {
                switchToCustomizeForm();
            } else if ("Import".equals(selectedItem)) {
                switchToImportForm();
            }
        });
        initializeFileWatcher();
        // TableView 列绑定
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
    
        styleTable.setItems(styleData);
        loadStylesFromXml();
        styleTable.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        
        // 设置表格列的相对宽度百分比
        styleTable.widthProperty().addListener((obs, oldVal, newVal) -> {
            double width = newVal.doubleValue();
            // 设置各列的相对宽度百分比
            valColumn.setPrefWidth(width * 0.15);          // 12%
            fontColumn.setPrefWidth(width * 0.15);         // 25% - 最宽
            typeColumn.setPrefWidth(width * 0.10);         // 10%
            colorColumn.setPrefWidth(width * 0.10);        // 6% - 最窄
            fontSizeColumn.setPrefWidth(width * 0.10);
            alignmentColumn.setPrefWidth(width * 0.10);     // 12%
            paragraphBeforeSpacingColumn.setPrefWidth(width * 0.10); // 11%
            paragraphSpacingColumn.setPrefWidth(width * 0.10);      // 11%
            lineSpacingColumn.setPrefWidth(width * 0.10);           // 11%
        });
    }

    private void loadStylesFromXml() {
        try {
            styleData.clear();
            InputStream inputStream = getClass().getClassLoader().getResourceAsStream("styles.xml");
            if (inputStream == null) {
                throw new IOException("styles.xml not found in classpath");
            }
            DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
            DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
            Document doc = dBuilder.parse(inputStream);
            doc.getDocumentElement().normalize();
            NodeList styleNodes = doc.getElementsByTagName("xml-fragment");
            for (int i = 0; i < styleNodes.getLength(); i++) {
                Node styleNode = styleNodes.item(i);
                if (styleNode.getNodeType() == Node.ELEMENT_NODE) {
                    Element styleElement = (Element) styleNode;
                    
                    // 获取样式类型
                    String type = styleElement.getAttribute("w:type");
                    
                    // 获取样式名称
                    NodeList nameNodes = styleElement.getElementsByTagName("w:name");
                    String styleName = "";
                    if (nameNodes.getLength() > 0) {
                        Element nameElement = (Element) nameNodes.item(0);
                        styleName = nameElement.getAttribute("w:val");
                    }
                    // 获取样式ID
                    NodeList IdNodes = styleElement.getElementsByTagName("w:link");
                    String styleId = "";
                    if (IdNodes.getLength() > 0) {
                        Element nameElement = (Element) IdNodes.item(0);
                        styleId = nameElement.getAttribute("w:val");
                    }else{
                        styleId = styleName;
                    }
                    
                    // 获取中英文字体信息
                    String englishFont = "黑体"; // 默认英文字体为黑体
                    String chineseFont = "黑体"; // 默认中文字体为黑体
                    String font = "黑体"; // 默认字体为黑体
                    
                    NodeList rFontsNodes = styleElement.getElementsByTagName("w:rFonts");
                    if (rFontsNodes.getLength() > 0) {
                        Element rFontsElement = (Element) rFontsNodes.item(0);
                        // 获取英文字体 (ascii 或 hAnsi)
                        String asciiFont = rFontsElement.getAttribute("w:ascii");
                        String hAnsiFont = rFontsElement.getAttribute("w:hAnsi");
                        if (!asciiFont.isEmpty()) {
                            englishFont = asciiFont;
                        } else if (!hAnsiFont.isEmpty()) {
                            englishFont = hAnsiFont;
                        }
                        
                        // 获取中文字体 (eastAsia)
                        String eastAsiaFont = rFontsElement.getAttribute("w:eastAsia");
                        if (!eastAsiaFont.isEmpty()) {
                            chineseFont = eastAsiaFont;
                        }
                        
                        // 组合中英文字体
                        if (englishFont.equals(chineseFont)) {
                            font = englishFont;
                        } else {
                            font = chineseFont + " / " + englishFont;
                        }
                    }
                    
                    
                    // 获取颜色信息
                    String color = "#000000"; // 默认黑色
                    NodeList colorNodes = styleElement.getElementsByTagName("w:color");
                    if (colorNodes.getLength() > 0) {
                        Element colorElement = (Element) colorNodes.item(0);
                        String colorVal = colorElement.getAttribute("w:val");
                        if (!colorVal.isEmpty()) {
                            color = "#" + colorVal;
                        }
                    }
                    
                    // 获取字体大小
                    int fontSize = 12; // 默认字体大小
                    NodeList szNodes = styleElement.getElementsByTagName("w:sz");
                    if (szNodes.getLength() > 0) {
                        Element szElement = (Element) szNodes.item(0);
                        String szVal = szElement.getAttribute("w:val");
                        if (!szVal.isEmpty()) {
                            fontSize = Integer.parseInt(szVal) / 2; // Word中字号是实际大小的两倍
                        }
                    }
                    
                    // 获取段落间距（前后）
                    double paragraphSpacing = 0.0; // 段落后间距
                    double paragraphBeforeSpacing = 0.0; // 段落前间距
                    NodeList spacingNodes = styleElement.getElementsByTagName("w:spacing");
                    if (spacingNodes.getLength() > 0) {
                        Element spacingElement = (Element) spacingNodes.item(0);
                        String afterVal = spacingElement.getAttribute("w:after");
                        String beforeVal = spacingElement.getAttribute("w:before");
                        if (!afterVal.isEmpty()) {
                            paragraphSpacing = Double.parseDouble(afterVal) / 20.0; // 转换为磅值
                        }
                        if (!beforeVal.isEmpty()) {
                            paragraphBeforeSpacing = Double.parseDouble(beforeVal) / 20.0; // 转换为磅值
                        }
                    }
                    // 获取对齐方式
                    String alignment = "left"; // 默认左对齐
                    NodeList alignmentNodes = styleElement.getElementsByTagName("w:jc");
                    if (alignmentNodes.getLength() > 0) {
                        Element alignmentElement = (Element) alignmentNodes.item(0);
                        String val = alignmentElement.getAttribute("w:val");
                        if (!val.isEmpty()) {
                            switch (val) {
                                case "left":
                                    alignment = "left";
                                    break;
                                case "center":
                                    alignment = "center";
                                    break;
                                case "right":
                                    alignment = "right";
                                    break;
                                case "both":
                                    alignment = "both";
                                    break;
                                case "distribute":
                                    alignment = "distribute";
                                    break;
                            }
                        }else{
                            alignment = "left";
                        }
                    }
                    // 获取行间距
                    double lineSpacing = 1.0; // 默认单倍行距
                    if (spacingNodes.getLength() > 0) {
                        Element spacingElement = (Element) spacingNodes.item(0);
                        String lineVal = spacingElement.getAttribute("w:line");
                        String lineRule = spacingElement.getAttribute("w:lineRule");
                        if (!lineVal.isEmpty()) {
                            if ("auto".equals(lineRule)) {
                                lineSpacing = Double.parseDouble(lineVal) / 240.0; // 自动行距转换
                            } else {
                                lineSpacing = Double.parseDouble(lineVal) / 20.0; // 固定行距转换为磅值
                            }
                        }
                    }
                    
                    // 添加到样式数据中
                    styleData.add(new StyleModel(styleId, styleName, font, type, color, fontSize,alignment, paragraphSpacing, lineSpacing, paragraphBeforeSpacing));
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
            System.err.println("Failed to load styles.xml: " + e.getMessage());
        }
    }
    private void initializeFileWatcher() {

        executorService = Executors.newSingleThreadExecutor();
        executorService.submit(() -> {
            try {
                Path path = Paths.get("target", "classes");
                WatchService watchService = FileSystems.getDefault().newWatchService();
                path.register(watchService, StandardWatchEventKinds.ENTRY_MODIFY);

                boolean running = true;
                while (running && !Thread.currentThread().isInterrupted()) {
                    try {
                        WatchKey key = watchService.take();
                        for (WatchEvent<?> event : key.pollEvents()) {
                            WatchEvent.Kind<?> kind = event.kind();
                            if (kind == StandardWatchEventKinds.ENTRY_MODIFY) {
                                Path changed = (Path) event.context();
                                if ("styles.xml".equals(changed.toString())) {
                                    // 重新加载 styles.xml 文件
                                    javafx.application.Platform.runLater(this::loadStylesFromXml);
                                }
                            }
                        }
                        boolean valid = key.reset();
                        if (!valid) {
                            break;
                        }
                    } catch (InterruptedException e) {
                        // 线程被中断，优雅地退出循环
                        running = false;
                        Thread.currentThread().interrupt(); // 重新设置中断状态
                        System.out.println("文件监视线程被中断，正在退出...");
                    }
                }
            } catch (Exception e) {
                e.printStackTrace();
                System.err.println("Failed to watch styles.xml");
            }
        });
    }
    private void switchToCustomizeForm() {
        try {
            // 加载 customize-form 的 FXML 文件
            // 获取当前 Stage
            FXMLLoader fxmlLoader2 = new FXMLLoader(HelloApplication.class.getResource("Customize-form.fxml"));
            Stage stage = new Stage();
            // 设置新的场景
            stage.initModality(Modality.APPLICATION_MODAL);
            stage.setTitle("Customize Form");
            stage.setScene(new Scene(fxmlLoader2.load(), 800, 600));
            stage.showAndWait();
        } catch (IOException e) {
            e.printStackTrace();
            System.err.println("Failed to load customize_form.fxml");
        }
    }
    
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
        } catch (IOException e) {
            e.printStackTrace();
            System.err.println("Failed to load ImportStyle.fxml");
        }
    }
    
    @FXML
    private void injectToDocx() {
        // 打开文件选择对话框，让用户选择目标 .docx 文件
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Select Target DOCX File");
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Word Files", "*.docx"));
        File selectedFile = fileChooser.showOpenDialog(injectToDocx.getScene().getWindow());

        if (selectedFile != null) {
            applyStylesToWord(selectedFile.getAbsolutePath());
        }
    }

    /**
     * 将选中的样式应用到 Word 文档
     *
     * @param filePath 目标 Word 文档路径
     */
    private void applyStylesToWord(String filePath) {
        try {
            // 获取选中的样式
            ObservableList<StyleModel> selectedStyles = styleTable.getSelectionModel().getSelectedItems();
            if (selectedStyles.isEmpty()) {
                CommonUtils.showAlert("警告", "请先选择要应用的样式！", Alert.AlertType.WARNING);
                return;
            }

            // 加载目标文档
            File file = new File(filePath);
            FileInputStream fis = new FileInputStream(file);
            XWPFDocument document = new XWPFDocument(fis);
            fis.close();

            // 加载styles.xml文件以获取原始XML片段
            InputStream stylesStream = getClass().getClassLoader().getResourceAsStream("styles.xml");
            if (stylesStream == null) {
                throw new IOException("styles.xml not found in classpath");
            }
            DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
            DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
            Document stylesDoc = dBuilder.parse(stylesStream);
            stylesDoc.getDocumentElement().normalize();
            stylesStream.close();

            // 获取所有xml-fragment节点
            NodeList styleNodes = stylesDoc.getElementsByTagName("xml-fragment");
            Map<String, Node> styleNodeMap = new HashMap<>();
            
            // 创建样式ID到节点的映射
            for (int i = 0; i < styleNodes.getLength(); i++) {
                Node node = styleNodes.item(i);
                Element element = (Element) node;
                NodeList styleIdNodes = element.getElementsByTagName("w:link");
                if (styleIdNodes.getLength() > 0) {
                    String styleId = styleIdNodes.item(0).getAttributes().getNamedItem("w:val").getNodeValue();
                    styleNodeMap.put(styleId, node);
                }
            }

            for (StyleModel styleModel : selectedStyles) {
                try {
                    // 创建新的样式
                    XWPFStyle style = new XWPFStyle(CTStyle.Factory.newInstance());
                    CTStyle ctStyle = style.getCTStyle();
                    
                    // 设置样式属性
                    String newStyleId = "Custom" + UUID.randomUUID().toString().substring(0, 8);
                    style.setStyleId(newStyleId);
                    style.setType(STStyleType.Enum.forString(styleModel.type()));
                    ctStyle.addNewName().setVal(styleModel.name());
                    ctStyle.addNewQFormat();
                    
                    // 配置字体属性
                    CTRPr rpr = ctStyle.addNewRPr();
                    // 设置字体
                    String fontValue = styleModel.font();
                    String englishFont = "黑体";
                    String chineseFont = "黑体";
                    
                    // 解析组合的字体字符串
                    if (fontValue.contains("/")) {
                        String[] fonts = fontValue.split("/");
                        if (fonts.length >= 2) {
                            chineseFont = fonts[0].trim();
                            englishFont = fonts[1].trim();
                        }
                    } else {
                        // 如果只有一个字体，则中英文使用相同字体
                        englishFont = fontValue;
                        chineseFont = fontValue;
                    }
                    
                    // 设置各种字体属性
                    CTFonts rFonts = rpr.addNewRFonts();
                    rFonts.setAscii(englishFont);
                    rFonts.setHAnsi(englishFont);
                    rFonts.setCs(englishFont);
                    rFonts.setEastAsia(chineseFont);
                    rpr.addNewSz().setVal(BigInteger.valueOf(styleModel.fontSize() * 2));
                    
                    // 处理颜色
                    String color = styleModel.color();
                    String correctedColor = color.startsWith("#") ? color.substring(1) : color;
                    rpr.addNewColor().setVal(correctedColor.toUpperCase());
                    
                    CTPPrGeneral pPr = ctStyle.addNewPPr();
                    // 设置对齐方式
                    String alignment = styleModel.alignment();
                    pPr.addNewJc().setVal(STJc.Enum.forString(alignment));
                    
                    // 设置段落间距
                    double paragraphSpacing = styleModel.paragraphSpacing();
                    double paragraphBeforeSpacing = styleModel.paragraphBeforeSpacing();
                    pPr.addNewSpacing().setAfter(BigInteger.valueOf((int) (paragraphSpacing * 20)));
                    pPr.addNewSpacing().setBefore(BigInteger.valueOf((int) (paragraphBeforeSpacing * 20)));

                    // 添加样式到文档
                    document.getStyles().addStyle(style);
                    
                    // 创建应用样式的段落
                    XWPFParagraph paragraph = document.createParagraph();
                    paragraph.setStyle(newStyleId);
                    XWPFRun run = paragraph.createRun();
                    run.setText("样式示例 (备用方法): " + styleModel.name());
                    
                } catch (Exception ex) {
                    System.err.println("备用方法也失败: " + ex.getMessage());
                }
            }

            // 保存文档
            FileOutputStream fos = new FileOutputStream(file);
            document.write(fos);
            fos.close();
            document.close();

            // 打开文档
            Desktop.getDesktop().open(file);

            // 显示成功消息
            CommonUtils.showAlert("成功", "成功应用样式到文档！");

        } catch (Exception e) {
            e.printStackTrace();
            CommonUtils.showAlert("错误", "应用样式时出错: " + e.getMessage(), Alert.AlertType.ERROR);
        }
    }
}
