package com.ind.word_style_controller.service;

import com.ind.StyleModel;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import java.io.File;


/**
 * 样式加载服务
 * 负责从XML文件中加载样式数据
 */
public class StyleLoaderService {
    
    // 样式文件路径
    private final String stylesXmlPath = "src/main/resources/styles.xml";
    
    /**
     * 从XML文件加载样式数据
     * @return 样式数据列表
     */
    public ObservableList<StyleModel> loadStylesFromXml() {
        ObservableList<StyleModel> styles = FXCollections.observableArrayList();
        
        try {
            // 从文件系统加载styles.xml文件
            File stylesFile = new File(stylesXmlPath);
            if (!stylesFile.exists()) {
                System.err.println("Could not find styles.xml at " + stylesXmlPath);
                return styles;
            }
            
            // 检查文件大小，如果文件为空或太小，可能是新创建的或被清空的
            if (stylesFile.length() < 50) { // 一个基本的XML结构至少需要这么多字节
                System.out.println("styles.xml is empty or contains only basic structure");
                return styles;
            }
            
            // 解析XML文档
            DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
            DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
            Document doc = dBuilder.parse(stylesFile);
            doc.getDocumentElement().normalize();
            
            // 检查是否有xml-fragment节点，如果没有，则认为文件只包含基本结构
            NodeList styleNodes = doc.getElementsByTagName("xml-fragment");
            if (styleNodes.getLength() == 0) {
                System.out.println("styles.xml contains only basic structure without any style definitions");
                return styles;
            }
            
            for (int i = 0; i < styleNodes.getLength(); i++) {
                Node styleNode = styleNodes.item(i);
                if (styleNode.getNodeType() == Node.ELEMENT_NODE) {
                    Element styleElement = (Element) styleNode;
                    
                    // 解析样式属性
                    StyleModel styleModel = parseStyleElement(styleElement);
                    styles.add(styleModel);
                }
            }
        } catch (Exception e) {
            // 不打印堆栈跟踪，只显示简短的错误信息
            // 这样可以避免在正常操作（如删除样式）导致的文件变化时显示大量错误信息
            System.out.println("Note: styles.xml may be in transition due to recent changes. Will retry on next refresh.");
            System.err.println("Failed to load styles.xml: " + e.getMessage());
        }
        
        return styles;
    }
    
    /**
     * 解析样式元素
     * @param styleElement 样式元素
     * @return 样式模型对象
     */
    private StyleModel parseStyleElement(Element styleElement) {
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
        } else {
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
            } else {
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
        // 创建并返回样式模型对象
        return new StyleModel(styleId, styleName, font, type, color, fontSize, alignment, paragraphSpacing, lineSpacing, paragraphBeforeSpacing);
    }
    
}