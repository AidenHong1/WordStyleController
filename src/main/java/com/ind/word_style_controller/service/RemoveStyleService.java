package com.ind.word_style_controller.service;

import com.ind.StyleModel;
import javafx.collections.ObservableList;
import javafx.scene.control.Alert;
import javafx.scene.control.ButtonType;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

public class RemoveStyleService {
    private final StyleLoaderService styleLoaderService;
    private final StyleTableController styleTableController;

    public RemoveStyleService(StyleLoaderService styleLoaderService, StyleTableController styleTableController) {
        this.styleLoaderService = styleLoaderService;
        this.styleTableController = styleTableController;
    }

    /**
     * 删除选中的样式
     * @param selectedStyles 选中的样式列表
     * @param styleTable JavaFX TableView 控件
     * @param updateStyleCountLabel Runnable 用于更新样式计数标签
     */
    public void removeSelectedStyle(ObservableList<StyleModel> selectedStyles, javafx.scene.control.TableView<StyleModel> styleTable, Runnable updateStyleCountLabel) {
        if (selectedStyles == null || selectedStyles.isEmpty()) {
            Alert alert = new Alert(Alert.AlertType.WARNING);
            alert.setTitle("警告");
            alert.setHeaderText(null);
            alert.setContentText("请先选择要删除的样式！");
            alert.showAndWait();
            return;
        }

        String confirmMessage;
        if (selectedStyles.size() == 1) {
            confirmMessage = "确定要删除样式 \"" + selectedStyles.get(0).getName() + "\" 吗？此操作不可撤销。";
        } else {
            confirmMessage = "确定要删除选中的 " + selectedStyles.size() + " 个样式吗？此操作不可撤销。";
        }

        Alert confirmAlert = new Alert(Alert.AlertType.CONFIRMATION);
        confirmAlert.setTitle("确认删除");
        confirmAlert.setHeaderText(null);
        confirmAlert.setContentText(confirmMessage);
        Optional<ButtonType> result = confirmAlert.showAndWait();

        if (result.isPresent() && result.get() == ButtonType.OK) {
            List<String> styleIds = new ArrayList<>();
            for (StyleModel style : selectedStyles) {
                styleIds.add(style.getId());
            }

            int xmlRemovedCount = this.removeStylesByIds(styleIds);
            int tableRemovedCount = this.removeStylesByIdsFromTable(styleIds, styleTable.getItems());
            styleTable.getSelectionModel().clearSelection();
            System.out.print("Removed " + tableRemovedCount + " styles from the table, " + xmlRemovedCount + " styles from XML.");
            if (updateStyleCountLabel != null) {
                updateStyleCountLabel.run();
            }

            if (tableRemovedCount > 0) {
                Alert successAlert = new Alert(Alert.AlertType.INFORMATION);
                successAlert.setTitle("成功");
                successAlert.setHeaderText(null);
                if (xmlRemovedCount == tableRemovedCount) {
                    successAlert.setContentText("已成功删除 " + xmlRemovedCount + " 个样式！");
                } else if (xmlRemovedCount > tableRemovedCount) {
                    successAlert.setContentText("已删除 " + tableRemovedCount + " 个样式，但其中有 " + xmlRemovedCount + " 个样式从XML文件中删除成功。可能是由于自定义样式的id一致所导致的");
                } else {
                    successAlert.setContentText("已从表格中删除 " + tableRemovedCount + " 个样式，但无法从XML文件中删除。请检查文件权限或格式。");
                }
                successAlert.showAndWait();
            } else {
                Alert errorAlert = new Alert(Alert.AlertType.ERROR);
                errorAlert.setTitle("错误");
                errorAlert.setHeaderText(null);
                errorAlert.setContentText("删除样式时出现错误！无法从表格中删除样式。");
                errorAlert.showAndWait();
            }
        }
    }

    /**
     * 批量删除多个样式（迁移自 StyleLoaderService）
     * @param styleIds 要删除的样式ID列表
     * @return 成功删除的样式数量
     */
    public int removeStylesByIds(List<String> styleIds) {
        if (styleIds == null || styleIds.isEmpty()) {
            return 0;
        }
        int successCount = 0;
        try {
            // 获取styles.xml文件
            java.io.File stylesFile = new java.io.File("src/main/resources/styles.xml");
            if (!stylesFile.exists()) {
                throw new java.io.IOException("styles.xml not found at src/main/resources/styles.xml");
            }
            javax.xml.parsers.DocumentBuilderFactory dbFactory = javax.xml.parsers.DocumentBuilderFactory.newInstance();
            javax.xml.parsers.DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
            org.w3c.dom.Document doc = dBuilder.parse(stylesFile);
            doc.getDocumentElement().normalize();
            org.w3c.dom.NodeList styleNodes = doc.getElementsByTagName("xml-fragment");
            java.util.List<org.w3c.dom.Node> nodesToRemove = new java.util.ArrayList<>();
            for (int i = 0; i < styleNodes.getLength(); i++) {
                org.w3c.dom.Node styleNode = styleNodes.item(i);
                if (styleNode.getNodeType() == org.w3c.dom.Node.ELEMENT_NODE) {
                    org.w3c.dom.Element styleElement = (org.w3c.dom.Element) styleNode;
                    String currentStyleId = null;
                    if (currentStyleId == null || currentStyleId.isEmpty()) {
                        org.w3c.dom.NodeList linkNodes = styleElement.getElementsByTagName("w:link");
                        if (linkNodes.getLength() > 0) {
                            org.w3c.dom.Element linkElement = (org.w3c.dom.Element) linkNodes.item(0);
                            currentStyleId = linkElement.getAttribute("w:val");
                        }
                    }
                    if (currentStyleId == null || currentStyleId.isEmpty()) {
                        org.w3c.dom.NodeList nameNodes = styleElement.getElementsByTagName("w:name");
                        if (nameNodes.getLength() > 0) {
                            org.w3c.dom.Element nameElement = (org.w3c.dom.Element) nameNodes.item(0);
                            currentStyleId = nameElement.getAttribute("w:val");
                        }
                    }
                    if (currentStyleId != null && styleIds.contains(currentStyleId)) {
                        nodesToRemove.add(styleNode);
                        successCount++;
                    }
                }
            }
            for (org.w3c.dom.Node node : nodesToRemove) {
                node.getParentNode().removeChild(node);
            }
            if (!nodesToRemove.isEmpty()) {
                javax.xml.transform.TransformerFactory transformerFactory = javax.xml.transform.TransformerFactory.newInstance();
                javax.xml.transform.Transformer transformer = transformerFactory.newTransformer();
                transformer.setOutputProperty(javax.xml.transform.OutputKeys.INDENT, "no");
                javax.xml.transform.dom.DOMSource source = new javax.xml.transform.dom.DOMSource(doc);
                javax.xml.transform.stream.StreamResult result = new javax.xml.transform.stream.StreamResult(stylesFile);
                transformer.transform(source, result);
            }
            return successCount;
        } catch (Exception e) {
            e.printStackTrace();
            System.err.println("Failed to remove styles: " + e.getMessage());
            return successCount;
        }
    }

    /**
     * 根据 id 列表批量删除样式（迁移自 StyleTableController）
     * @param ids 样式 id 列表
     * @param styleData 样式数据列表
     * @return 实际删除的数量
     */
    public int removeStylesByIdsFromTable(java.util.List<String> ids, javafx.collections.ObservableList<com.ind.StyleModel> styleData) {
        if (ids == null || ids.isEmpty()) return 0;
        java.util.Set<String> idSet = new java.util.HashSet<>(ids);
        styleData.removeIf(item -> idSet.contains(item.getId()));
        return ids.size();
    }
}