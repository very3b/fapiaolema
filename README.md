# 发票处理工具

一个用于处理发票和支付记录的自动化工具。

## 功能特点

### 1. 文件处理
- **PDF发票处理**
  - 自动提取发票金额保存到CSV中
  - 提取发票号码和开票日期保存到CSV中
  - 支持批量处理多个PDF文件
  - 提取供应商名称保存到CSV中
  - 发票文件名也保存到CSV 发票文件 列中
  - 按文件名排序合并所有发票为一个PDF

- **支付截图处理**
  - 多种图像处理方法提高识别率：
    - 原始灰度图处理
    - CLAHE对比度增强
    - Otsu自适应二值化
    - 高斯自适应二值化
  - 多种OCR识别模式（PSM）：
    - PSM 3: 自动页面分割
    - PSM 4: 单列可变文本
    - PSM 6: 统一文本块
    - PSM 7: 单行文本
    - PSM 8: 单词识别
    - PSM 11: 稀疏文本
    - PSM 12: 稀疏文本和OSD
    - PSM 13: 原始单行文本
  - 智能金额选择：
    - 统计所有识别结果
    - 选择出现频率最高的金额
    - 详细的识别日志
  - 支持多种金额格式：
    - 标准格式：-xx.xx
    - 带空格：- xx.xx
    - 不同的负号：−xx.xx
    - 括号格式：(xx.xx)
    - 带"支付"的格式
    - 带"¥"的格式
    - 带"实付"的格式
    - 带"付款"的格式
    - 带"总计"的格式
  - 仅处理文件名包含"log"的图片
  - 支持 jpg/jpeg/png 格式
  - 自动跳过不包含"log"的图片
  - 自动和发票文件匹配，金额误差小于10%
  - 按文件名排序合并为PDF

### 2. 文件输出
- **PDF合并**
  - 将所有发票PDF合并为一个文件：`merged_{date}.pdf`
  - 将所有支付截图合并为一个文件：`merged_{date}_log.pdf`
  - 所有输出文件统一保存到output目录

- **数据汇总**
  - 前面8列头是：
    -- *名称（使用文件名去掉数字和log字符）
    -- *品牌（一律是NA）
    -- *规格数量（一律是1）
    -- *规格单位（一律是套）
    -- *计量单位（一律是套）
    -- *数量（一律是1）
    -- *存放地点（一律是科技楼1907）
    -- *供应商（从发票提取）
  - 生成包含所有数据的CSV文件
  - 自动匹配发票和支付记录
  - 发票金额单独一列
  - 计算支付差额
  - 包含规格数量（默认为1）
  - csv自动创建所有列项目如果不存在
  - 输出文件：`combined_results.csv`

### 3. 用户界面
- 基于Flask和webview的GUI界面
- 支持文件夹选择
- 实时显示处理日志
- 简洁直观的操作方式

### 4. 打包和发布
 - 用pyinstaller打包
 - 用pyinstaller --onefile main.py
 - 用pyinstaller --onefile --windowed main.py

## 安装说明

1. 安装Python依赖：
```bash
pip install -r requirements.txt
```

2. 安装Tesseract OCR：
   - 下载并安装 [Tesseract](https://github.com/UB-Mannheim/tesseract/wiki)
   - 确保安装中文语言包

## 使用方法

1. 启动程序：
```bash
python app.py
```

2. 在GUI界面中：
   - 点击选择文件夹
   - 选择包含发票和支付截图的文件夹
   - 点击开始处理
   - 等待处理完成

3. 输出文件：
   - `output/merged_{date}.pdf`：合并后的发票文件
   - `output/merged_{date}_log.pdf`：合并后的支付截图
   - `output/combined_results.csv`：数据汇总文件

## 文件命名规则

- 支付截图必须包含"log"在文件名中
- 发票PDF和对应的支付截图应有相关联的文件名
- 示例：
  ```
  invoice_1.pdf
  invoice_1_log.jpg
  ```

## 注意事项

1. 确保文件夹中包含：
   - PDF格式的发票文件
   - 包含"log"的支付截图（jpg/jpeg/png）

2. 图片要求：
   - 支付金额清晰可见
   - 负数金额格式（如：-69.00）

3. 系统要求：
   - Python 3.6+
   - Tesseract OCR 5.0+
   - Windows/Linux/MacOS

## 开发说明

主要文件说明：
- `app.py`：GUI程序入口
- `pdf_image_analyzer.py`：PDF处理核心代码
- `test_image_payment.py`：图片处理核心代码

## 更新日志

### v1.1.0
- 增强了OCR识别功能
- 添加了多种图像处理方法
- 优化了金额提取算法
- 改进了文件合并功能
- 统一了输出目录结构

### v1.0.0
- 基础功能实现
- GUI界面添加
- 文件合并功能
- 数据汇总报告




