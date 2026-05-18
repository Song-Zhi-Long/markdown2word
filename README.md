# Markdown2Word

一个基于 Python + Tkinter 的小工具，用于把 Markdown 文本快速转换为 Word 文档（`.docx`）。

项目提供了简单的图形界面，适合将ai的markdown格式的回答整理为word笔记、将ai翻译的论文转化为word文档等场景。当前实现重点放在中文文档排版、公式转换、图片与超链接处理，以及更符合中文写作习惯的段落输出效果。

## GUI 效果

![GUI 界面](data/GUI_and_result.png)

## 目前支持

- 图形界面直接粘贴 Markdown 并导出 Word
- 标题、普通段落、粗体、斜体、下划线、代码块
- 表格、图片、超链接
- 行内公式与独占一行公式
- 独占一行的公式会单独成段并居中显示且不首行缩进
- 块公式居中显示
- 正文首行缩进2字符
- 题注识别与居中
  - 例如 `表3：...`、`图2：...`、`Table 1: ...`
- 有序列表按原文重新从 `1.` 开始
- 转换完成后自动打开生成的 Word 文档
- setting.json中的
  - output_dir（输出目录），默认输出目录优先取用户的 Downloads，如果没有就尝试 下载，再不行就退回到用户主目录
  - asset_root（资源根目录），默认为当前工作目录，也就是你从哪里启动程序，默认资源根目录就指向哪里
  - title_chars（导出文件名截取的标题长度），默认是 12
  - auto_timestamp（是否自动在文件名后追加时间戳），默认是 True

## 后续调优历程

 - 2026年5月11日——添加正文首行缩进开关、修改了列表中字体无法加粗显示的bug、优化了列表的缩进美观、完善对exe程序的读取配置规则
 - 2026年5月18日——针对豆包给的文本存在
    ```text
    ## 图1：实验结果对比
    我们的方法……
    ```
    识别为二级标题而不能正确识别为图表标题的情况，代码添加了预处理阶段，将其合并为一个普通段落，再走 Caption 样式识别

## 适用场景

- 将 Markdown 笔记快速整理成 Word 文档
- 将ai帮忙完成的课程作业、实验报告、论文翻译稿从 Markdown 导出为 `.docx`
- 将ai回答的有用知识从 Markdown 导出为 `.docx`，方便整理笔记
- 将包含公式、图片、表格的中文技术文档转换为可继续编辑的 Word 文件

## 环境要求

- Python 3.10 及以上版本
- Windows 环境下可直接运行图形界面

## 运行方式

先安装依赖：

```bash
pip install -r requirements.txt
```

启动程序：

```bash
python markdown2word.py
```

**或者直接在windows系统中双击markdown2word.exe即可运行**

```text
打开 markdown2word 文件夹双击 markdown2word.exe 文件
（注：settings.json 配置文件需与 exe 文件位于同一文件夹内）
```

## 项目文件

```text
markdown2word.py    GUI 入口
converter.py        Markdown -> Word 核心逻辑
mml2omml.xsl        公式转换所需 XSLT
settings.json       本地默认配置
data/GUI.png        GUI 截图
data/app_icon.ico   程序图标
```

## 后续可扩展方向

- 支持更多 Markdown 扩展语法
- 支持自定义 Word 模板
- 支持更细致的段落、字体、页边距和标题样式控制
- 支持批量导入 Markdown 文件并自动导出

## exe打包指令

```bash
pyinstaller --noconfirm --clean --windowed --onefile --icon data/app_icon.ico --add-data "mml2omml.xsl;." --collect-data latex2mathml markdown2word.py
```