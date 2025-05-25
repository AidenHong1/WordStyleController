package com.ind.word_style_controller;

import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.ListView;
import javafx.scene.control.SelectionMode;
import javafx.stage.FileChooser;
import javafx.stage.Modality;
import javafx.stage.Stage;
import javafx.application.Platform;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STStyleType;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.awt.*;
import java.io.*;
import java.math.BigInteger;
import java.nio.file.*;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

public class HelloController {
    @FXML
    public ListView<String> menuList;
    public ListView<String> styleList;
    public Button injectToDocx;
    private ExecutorService executorService;
    public void initialize() {
        menuList.getItems().addAll("Customize", "Import", "About");
        menuList.setOnMouseClicked(event -> {
            String selectedItem = menuList.getSelectionModel().getSelectedItem();
            if ("Customize".equals(selectedItem)) {
                switchToCustomizeForm();
            }
        });
        initializeFileWatcher();
        loadStylesFromXml();
        styleList.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
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
    private void loadStylesFromXml() {
        try {
            styleList.getItems().clear();
            InputStream inputStream = getClass().getClassLoader().getResourceAsStream("styles.xml");
            if (inputStream == null) {
                throw new IOException("styles.xml not found in classpath");
            }
            DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
            DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
            Document doc = dBuilder.parse(inputStream);
            doc.getDocumentElement().normalize();

            NodeList styleNodes = doc.getElementsByTagName("style");
            for (int i = 0; i < styleNodes.getLength(); i++) {
                Node styleNode = styleNodes.item(i);
                if (styleNode.getNodeType() == Node.ELEMENT_NODE) {
                    Element styleElement = (Element) styleNode;
                    String id = styleElement.getAttribute("id");
                    String name = styleElement.getElementsByTagName("name").item(0).getTextContent();
                    String type = styleElement.getElementsByTagName("type").item(0).getTextContent();
                    String color = styleElement.getElementsByTagName("color").item(0).getTextContent();
                    String fontSize = styleElement.getElementsByTagName("fontSize").item(0).getTextContent();
                    // 将样式信息格式化为字符串
                    String styleInfo = String.format("ID: %s, Name: %s, Type: %s, Color: %s, Font Size: %s",
                            id, name, type, color, fontSize);

                    // 添加到 styleList
                    styleList.getItems().add(styleInfo);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
            System.err.println("Failed to load styles.xml");
        }
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
            // 加载现有的 Word 文档
            XWPFDocument document = new XWPFDocument(new FileInputStream(filePath));

            // 获取选中的样式信息
            List<String> selectedStyles = styleList.getSelectionModel().getSelectedItems();
            for (String styleInfo : selectedStyles) {
                String[] parts = styleInfo.split(", ");
                String id = parts[0].split(": ")[1];
                String name = parts[1].split(": ")[1];
                String type = parts[2].split(": ")[1];
                String color = parts[3].split(": ")[1];
                String fontSize = parts[4].split(": ")[1];

                // 创建一个新的样式
                XWPFStyle style = new XWPFStyle(CTStyle.Factory.newInstance());
                // 设置样式属性
                CTStyle ctStyle = style.getCTStyle();
                style.setStyleId(id);
                style.setType(STStyleType.PARAGRAPH);
                ctStyle.addNewBasedOn().setVal("Normal");

                // 设置字体
                ctStyle.addNewRPr().addNewRFonts().setAscii(name);
                ctStyle.getRPr().addNewSz().setVal(new BigInteger(fontSize));
                ctStyle.getRPr().addNewSzCs().setVal(new BigInteger(fontSize));

                String correctedColor = "FF" + color.replace("#", "").toUpperCase(); // 添加透明度并转换为大写
                ctStyle.getRPr().addNewColor().setVal(correctedColor);

                // 将样式添加到文档中
                document.getStyles().addStyle(style);
                
                // 创建一个示例段落并应用样式
                XWPFParagraph paragraph = document.createParagraph();
                paragraph.setStyle(id);
                XWPFRun run = paragraph.createRun();
                run.setText("样式示例: " + name + " (ID: " + id + ")");
            }

            // 保存修改后的文档
            try (FileOutputStream out = new FileOutputStream(filePath)) {
                document.write(out);
            }

            document.close();

            // 使用默认程序打开修改后的文档
            Desktop.getDesktop().open(new File(filePath));
        } catch (Exception e) {
            e.printStackTrace();
            System.err.println("Failed to apply styles to the Word document");
        }
    }
//    @FXML
//    private Label welcomeText;
//
//    @FXML
//    protected void onHelloButtonClick() {
//        welcomeText.setText("Welcome to JavaFX Application!");
//
}