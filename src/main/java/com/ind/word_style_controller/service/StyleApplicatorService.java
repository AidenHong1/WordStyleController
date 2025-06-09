package com.ind.word_style_controller.service;

import com.ind.StyleModel;
import com.ind.word_style_controller.utils.CommonUtils;
import javafx.collections.ObservableList;
import javafx.scene.control.Alert;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.HashMap;
import java.util.Map;
import java.util.UUID;

/**
 * 样式应用服务，负责将样式应用到Word文档
 */
public class StyleApplicatorService {
    
    /**
     * 将样式应用到Word文档
     * @param filePath 文档路径
     * @param selectedStyles 选中的样式列表
     */
    public void applyStylesToWord(String filePath, ObservableList<StyleModel> selectedStyles) {
        try {
            // 检查是否选择了样式
            if (selectedStyles.isEmpty()) {
                CommonUtils.showAlert("警告", "请先选择要应用的样式！", Alert.AlertType.WARNING);
                return;
            }

            // 加载目标文档
            File file = new File(filePath);
            FileInputStream fis = new FileInputStream(file);
            XWPFDocument document = new XWPFDocument(fis);
            fis.close();

            // 从XML文件中加载样式数据
            String stylesXmlPath = "src/main/resources/styles.xml";
            File stylesFile = new File(stylesXmlPath);
            if (!stylesFile.exists()) {
                throw new IOException("styles.xml not found at " + stylesXmlPath);
            }
            DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
            DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
            Document stylesDoc = dBuilder.parse(stylesFile);
            stylesDoc.getDocumentElement().normalize();

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

            // 应用每个选中的样式
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
                    run.setText("样式示例: " + styleModel.name());
                    
                } catch (Exception ex) {
                    System.err.println("应用样式失败: " + ex.getMessage());
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