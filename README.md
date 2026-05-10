# Markdown2Word

一个基于 Python + Tkinter 的小工具，用于把 Markdown 文本快速转换为 Word 文档（`.docx`）。

## GUI 效果

![GUI 界面](data/GUI.png)

## 目前支持

- 图形界面直接粘贴 Markdown 并导出 Word
- 标题、普通段落、粗体、斜体、下划线、代码块
- 表格、图片、超链接
- 行内公式与独占一行公式
- 块公式居中显示
- 正文首行缩进
- 题注识别与居中
  - 例如 `表3：...`、`图2：...`、`Table 1: ...`
- 有序列表按原文重新从 `1.` 开始
- 转换完成后自动打开生成的 Word 文档
- setting.json中的
  - output_dir，默认输出目录优先取用户的 Downloads，如果没有就尝试 下载，再不行就退回到用户主目录
  - asset_root，默认为当前工作目录，也就是你从哪里启动程序，默认资源根目录就指向哪里
  - title_chars，默认是 12
  - auto_timestamp，默认是 True

## 运行方式

先安装依赖：

```bash
pip install -r requirements.txt
```

启动程序：

```bash
python markdown2word.py
```

或者直接在windows系统中双击markdown2word.exe即可运行

## 项目文件

```text
markdown2word.py    GUI 入口
converter.py        Markdown -> Word 核心逻辑
mml2omml.xsl        公式转换所需 XSLT
settings.json       本地默认配置
data/GUI.png        GUI 截图
data/app_icon.ico   程序图标
```
## exe打包指令

```bash
pyinstaller --noconfirm --clean --windowed --onefile --icon data/app_icon.ico --add-data "mml2omml.xsl;." --collect-data latex2mathml markdown2word.py
```