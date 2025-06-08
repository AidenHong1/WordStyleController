package com.ind.word_style_controller.utils;

import javafx.scene.control.Alert;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.File;
import java.io.InputStream;
import java.io.StringWriter;

/**
 * 通用工具类，提供共享的功能方法
 */
public class CommonUtils {
    
    /**
     * 显示Alert对话框
     * @param title 标题
     * @param message 消息内容
     * @param type Alert类型
     */
    public static void showAlert(String title, String message, Alert.AlertType type) {
        Alert alert = new Alert(type);
        alert.setTitle(title);
        alert.setHeaderText(null);
        alert.setContentText(message);
        alert.showAndWait();
    }
    
    /**
     * 显示信息类型的Alert对话框
     * @param title 标题
     * @param message 消息内容
     */
    public static void showAlert(String title, String message) {
        showAlert(title, message, Alert.AlertType.INFORMATION);
    }
    
    /**
     * 创建XML文档构建器
     * @return DocumentBuilder实例
     * @throws Exception 创建失败时抛出异常
     */
    public static DocumentBuilder createDocumentBuilder() throws Exception {
        DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
        return dbFactory.newDocumentBuilder();
    }
    
    /**
     * 从classpath加载XML文档
     * @param resourcePath 资源路径
     * @return Document实例
     * @throws Exception 加载失败时抛出异常
     */
    public static Document loadXmlFromClasspath(String resourcePath) throws Exception {
        InputStream inputStream = CommonUtils.class.getClassLoader().getResourceAsStream(resourcePath);
        if (inputStream == null) {
            throw new RuntimeException(resourcePath + " not found in classpath");
        }
        DocumentBuilder dBuilder = createDocumentBuilder();
        return dBuilder.parse(inputStream);
    }
    
    /**
     * 从文件加载XML文档
     * @param file 文件对象
     * @return Document实例
     * @throws Exception 加载失败时抛出异常
     */
    public static Document loadXmlFromFile(File file) throws Exception {
        DocumentBuilder dBuilder = createDocumentBuilder();
        return dBuilder.parse(file);
    }
    
    /**
     * 向XML元素添加子元素
     * @param doc XML文档
     * @param parent 父元素
     * @param tagName 标签名
     * @param value 元素值
     */
    public static void addElement(Document doc, Element parent, String tagName, String value) {
        Element element = doc.createElement(tagName);
        element.setTextContent(value);
        parent.appendChild(element);
    }
    
    /**
     * 获取styles.xml文件路径
     * @return 文件路径
     */
    public static String getStylesXmlPath() {
        return "src/main/resources/styles.xml";
    }
    
    /**
     * 将 DOM 节点转换为字符串
     * @param node DOM 节点
     * @return 节点的 XML 字符串表示
     * @throws Exception 转换失败时抛出异常
     */
    public static String nodeToString(Node node) throws Exception {
        StringWriter sw = new StringWriter();
        TransformerFactory tf = TransformerFactory.newInstance();
        Transformer transformer = tf.newTransformer();
        transformer.transform(new DOMSource(node), new StreamResult(sw));
        return sw.toString();
    }
    
    /**
     * 确保目录存在
     * @param file 文件对象
     */
    public static void ensureDirectoryExists(File file) {
        File parentDir = file.getParentFile();
        if (parentDir != null && !parentDir.exists()) {
            parentDir.mkdirs();
        }
    }
}