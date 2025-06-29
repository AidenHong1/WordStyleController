# Word Style Controller

一个用于批量管理和自定义Word文档样式的JavaFX桌面应用。

## 功能特性
- 支持自定义段落和字符样式，包括字体、字号、颜色、对齐方式、间距等
- 批量添加、删除、导入、导出样式
- 通过可视化界面操作，无需手动编辑XML
- 样式数据以`styles.xml`文件存储，兼容Word格式
- 备注：目前仅支持对DOCX格式文档进行样式管理，暂不支持DOC格式类型文档

## 安装与运行
1. 克隆本项目：
   ```bash
   git clone https://github.com/yourname/word_style_controller.git
   ```
2. 使用内置Maven Wrapper构建：
   ```bash
   cd word_style_controller
   ./mvnw clean package
   ```
3. 运行JavaFX应用：
   ```bash
   cd word_style_controller
   ./mvnw javafx:run
   ```
   或直接运行`quick_boot.bat`批处理脚本。

## 依赖环境
- JDK 17 及以上（建议使用OpenJDK或Oracle JDK）
- 可选：Maven 3.6+（如需源码构建）
- 可选：JavaFX 17+（如需本地开发，已集成于依赖中）

本项目可直接通过双击`quick_boot.bat`直接运行运行，无需单独安装Maven或JavaFX SDK。

## 目录结构
```
├── src/main/java         # Java源代码
├── src/main/resources    # 资源文件（如styles.xml）
├── LICENSE               # MIT开源协议
├── README.md             # 项目说明文档
├── pom.xml               # Maven构建文件
├── quick_boot.bat        # 一键编译运行脚本
```

## 未来计划
- 支持更加完整的样式信息导入，如完善对继承关系的支持
- 支持创建“样式组”，支持将样式组（也可以理解为模板）应用到DOCX文档中
- 支持更加友好的用户界面，添加对段前间距、段后间距的计量单位选择

## 开源协议
本项目采用 [MIT License](./LICENSE) 开源，欢迎自由使用和二次开发。

## 联系方式
如有建议或问题，欢迎提交 issue 或联系作者。
