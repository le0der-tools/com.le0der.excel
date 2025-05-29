# Le0der Excel工具包（Excel Toolkit）

这是一个轻量级的 Unity Excel工具，支持根据excel结构创建对应所有sheet的数据类，并且不同excel表中重复sheet类会自动跳过。根据数据类，解析excel数据，并生成对应的数据类实例，excel更改时自动更新数据类实例。
该工具已按 Unity Package Manager（UPM）规范封装，可通过 Git 地址直接集成到 Unity 项目中。

---

## 📦 包信息
**包名**：`com.le0der.excel`

**最低支持 Unity 版本**：`2020.3` 

**作者**：[Le0der](https://github.com/le0der)

---

## ✨ 功能特色

- ✅ 集成[NPOI](https://github.com/tonyqus/npoi/)工具，不需要其他额外依赖，可独立使用
- ✅ Unity Editor内操作，无需额外软件
- ✅ 自动更新数据类实例，无需手动更新防止遗忘

---

## 📥 安装方式

你可以通过以下任一方式将该工具包集成到你的 Unity 项目中：

---
### ✅ 方法 1：使用 Unity 编辑器内的 Package Manager 添加 Git 地址（推荐）

1. 打开 Unity 的菜单：Window > Package Manager

2. 点击左上角的 + 号按钮

3. 选择 Add package from Git URL...

4. 输入：
```arduino
https://github.com/le0der-tools/com.le0der.excel.git
```
### ✅ 方法 2：使用 Git URL 添加依赖

1. 打开你的 Unity 项目
2. 编辑文件：`Packages/manifest.json`
3. 在 `"dependencies"` 节点中添加如下内容：

```json
"com.le0der.excel": "https://github.com/le0der-tools/com.le0der.excel.git"
```
