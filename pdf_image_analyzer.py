import os
import pdfplumber
import pytesseract
import cv2
import pandas as pd
from datetime import datetime
import re
import numpy as np
from pytesseract import Output
import traceback
import logging

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class DocumentAnalyzer:
    def __init__(self, folder_path):
        self.folder_path = folder_path
        self.results = []
        self.payment_images = {}  # 存储支付图片信息
        self.payment_patterns = [
            r'-\s*(\d+(?:[.,]\d{2})?)',  # 匹配以'-'开头的数字
            r'−\s*(\d+(?:[.,]\d{2})?)',  # 匹配以'−'开头的数字（不同的减号字符）
            r'[\-−]\s*¥?\s*(\d+(?:[.,]\d{2})?)',  # 匹配减号后带￥的数字
        ]
        self.logger = logging.getLogger(__name__)
        
        # 设置输出文件路径
        self.output_dir = folder_path
        self.invoice_results_file = os.path.join(self.output_dir, 'analysis_results.csv')
        self.payment_records_file = os.path.join(self.output_dir, 'payment_records.csv')

    def extract_pdf_info(self, pdf_path):
        """从PDF提取发票信息"""
        try:
            with pdfplumber.open(pdf_path) as pdf:
                text = ''
                for page in pdf.pages:
                    text += page.extract_text() or ''
                
                self.logger.info(f"\n提取的文本内容:")
                self.logger.info("-" * 50)
                self.logger.info(text)
                self.logger.info("-" * 50)
                
                # 提取发票号码
                invoice_patterns = [
                    r'发票号码[:：]\s*(\w+)',
                    r'发票号码\s*[:：]?\s*(\w+)',
                    r'NO[.：]\s*(\w+)',
                    r'发票代码[:：]\s*(\w+)',
                    r'[Nn][Oo]\.?\s*(\w+)'
                ]
                invoice_number = None
                for pattern in invoice_patterns:
                    match = re.search(pattern, text)
                    if match:
                        invoice_number = match.group(1).strip()
                        self.logger.info(f"找到发票号码: {invoice_number}")
                        break
                
                # 提取开票日期
                date_patterns = [
                    r'开票日期[:：]\s*(\d{4}[-年/]\d{1,2}[-月/]\d{1,2})',
                    r'开票日期\s*[:：]?\s*(\d{4}[-年/]\d{1,2}[-月/]\d{1,2})',
                    r'日期[:：]\s*(\d{4}[-年/]\d{1,2}[-月/]\d{1,2})',
                    r'(\d{4}[-年/]\d{1,2}[-月/]\d{1,2})\s*日期'
                ]
                invoice_date = None
                for pattern in date_patterns:
                    match = re.search(pattern, text)
                    if match:
                        date_str = match.group(1)
                        # 统一日期格式
                        date_str = date_str.replace('年', '-').replace('月', '-').replace('日', '').replace('/', '-')
                        invoice_date = date_str
                        self.logger.info(f"找到开票日期: {invoice_date}")
                        break
                
                # 提取供应商名称
                supplier_patterns = [
                    r'名\s*称[:：]\s*([^\n]*)',
                    r'销\s*售\s*方[:：]\s*([^\n]*)',
                    r'供\s*应\s*商[:：]\s*([^\n]*)',
                    r'销售方名称[:：]\s*([^\n]*)',
                    r'公司名称[:：]\s*([^\n]*)'
                ]
                supplier = None
                for pattern in supplier_patterns:
                    match = re.search(pattern, text)
                    if match:
                        supplier = match.group(1).strip()
                        # 清理供应商名称中的特殊字符
                        supplier = re.sub(r'[^\w\s\u4e00-\u9fff]', '', supplier)
                        self.logger.info(f"找到供应商: {supplier}")
                        break
                
                # 提取金额
                amount_patterns = [
                    r'金额[:：]\s*[¥￥]?\s*(\d+[\.,]?\d*)',
                    r'合\s*计[:：]\s*[¥￥]?\s*(\d+[\.,]?\d*)',
                    r'价税合计[:：]\s*[¥￥]?\s*(\d+[\.,]?\d*)',
                    r'小写[:：]\s*[¥￥]?\s*(\d+[\.,]?\d*)',
                    r'[¥￥]\s*(\d+[\.,]?\d*)',
                    r'人民币\s*[¥￥]?\s*(\d+[\.,]?\d*)',
                    r'总额[:：]\s*[¥￥]?\s*(\d+[\.,]?\d*)',
                    r'应付金额[:：]\s*[¥￥]?\s*(\d+[\.,]?\d*)'
                ]
                
                amount = None
                for pattern in amount_patterns:
                    matches = re.finditer(pattern, text)
                    for match in matches:
                        potential_amount = match.group(1).replace(',', '')
                        try:
                            current_amount = float(potential_amount)
                            if amount is None or current_amount > amount:
                                amount = current_amount
                                self.logger.info(f"找到金额: {amount}")
                        except ValueError:
                            continue
                
                # 提取商品名称
                product_patterns = [
                    r'货物或应税劳务、服务名称\s*([^\n]*)',
                    r'商品名称\s*([^\n]*)',
                    r'项目名称\s*([^\n]*)',
                    r'商品或服务名称\s*([^\n]*)'
                ]
                product_name = None
                for pattern in product_patterns:
                    match = re.search(pattern, text)
                    if match:
                        product_name = match.group(1).strip()
                        # 清理商品名称中的特殊字符
                        product_name = re.sub(r'[^\w\s\u4e00-\u9fff]', '', product_name)
                        self.logger.info(f"找到商品名称: {product_name}")
                        break
                
                # 如果没有找到商品名称，尝试从文件名提取
                if not product_name:
                    product_name = self.extract_product_name_from_filename(pdf_path)
                    if product_name:
                        self.logger.info(f"从文件名提取的商品名称: {product_name}")
                
                # 记录提取结果
                self.logger.info("\n发票信息提取结果:")
                self.logger.info(f"发票号码: {invoice_number}")
                self.logger.info(f"开票日期: {invoice_date}")
                self.logger.info(f"供应商: {supplier}")
                self.logger.info(f"金额: {amount}")
                self.logger.info(f"商品名称: {product_name}")
                
                return {
                    'invoice_number': invoice_number,
                    'invoice_date': invoice_date,
                    'supplier': supplier,
                    'price': f"{amount:.2f}" if amount is not None else None,
                    'product_name': product_name,
                    'filename': os.path.basename(pdf_path)
                }
                
        except Exception as e:
            self.logger.error(f"处理PDF文件时出错 {pdf_path}: {str(e)}")
            traceback.print_exc()
            return None

    def match_payment_to_invoice(self):
        """将支付图片与发票匹配并更新结果"""
        for result in self.results:
            if 'price' in result:  # 这是票记录
                invoice_amount = float(result['price'])
                filename_without_ext = os.path.splitext(result['filename'])[0]
                
                # 查找对应的支付图片
                matching_payment = None
                for payment_filename, payment_amount in self.payment_images.items():
                    payment_name_without_ext = os.path.splitext(payment_filename)[0]
                    
                    # 如果文件名（不含扩展名）相同或包含关系，认为是匹配的
                    if (payment_name_without_ext in filename_without_ext or 
                        filename_without_ext in payment_name_without_ext):
                        matching_payment = payment_amount
                        break
                
                # 更新结果，添加实际支付金额
                result['actual_payment'] = matching_payment

    def extract_payment_from_image(self, image_path):
        """Extract payment amount from image using OCR"""
        try:
            # 读取图片
            image = cv2.imdecode(np.fromfile(str(image_path), dtype=np.uint8), cv2.IMREAD_COLOR)
            if image is None:
                print(f"错误：无法读取图片 {image_path}")
                return None
                
            # 转换为灰度图
            gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
            
            # 应用不同的图像预处理方法并获取文本数据
            results = []
            
            # 1. 原始灰度图
            data = pytesseract.image_to_data(gray, lang='chi_sim', output_type=Output.DATAFRAME)
            self.process_ocr_data(data, "原始灰度图", results)
            
            # 2. OTSU 二值化
            _, threshold = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            data = pytesseract.image_to_data(threshold, lang='chi_sim', output_type=Output.DATAFRAME)
            self.process_ocr_data(data, "OTSU二值化", results)
            
            # 3. 自适应阈值
            adaptive = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
                                          cv2.THRESH_BINARY, 11, 2)
            data = pytesseract.image_to_data(adaptive, lang='chi_sim', output_type=Output.DATAFRAME)
            self.process_ocr_data(data, "自适应阈值", results)

            # 分析结果
            if results:
                # 按高度（字体大小）排序
                results.sort(key=lambda x: x[1], reverse=True)
                print("\n所有找到的负数金额 (按字体大小排序):")
                for amount, height, method, pattern in results:
                    print(f"金额: {amount:.2f}, 字体高度: {height}, 来自: {method}")
                
                # 返回字体大的负数金额的绝对值
                return abs(results[0][0])
            else:
                print("\n未找到任何负数金额")
                return None

        except Exception as e:
            print(f"处理图片时出错: {str(e)}")
            return None

    def process_ocr_data(self, data, method_name, results):
        """处理OCR数据，提取数字及其高度信息"""
        # 清理数据
        data = data[data.conf != -1]  # 移除低置信度的结果
        
        # 遍历每个识别出的文本块
        for _, row in data.iterrows():
            if pd.isna(row.text) or str(row.text).isspace():
                continue
                
            text = str(row.text)
            height = row.height
            
            # 查找负数金额
            for pattern in self.payment_patterns:
                matches = re.finditer(pattern, text)
                for match in matches:
                    amount_str = match.group(1).replace(',', '')
                    try:
                        # 将金额转换为负数
                        amount_float = -float(amount_str)
                        if -1000000 <= amount_float <= -1:  # 设置合理的负数金额范围
                            results.append((amount_float, height, method_name, pattern))
                    except ValueError:
                        continue

    def analyze_documents(self):
        """Analyze all documents in the folder"""
        # 首先处理所有图片文件，存储支付信息
        for filename in os.listdir(self.folder_path):
            if filename.lower().endswith(('.png', '.jpg', '.jpeg')):
                file_path = os.path.join(self.folder_path, filename)
                payment_amount = self.extract_payment_from_image(file_path)
                if payment_amount:
                    self.payment_images[filename] = payment_amount
                    print(f"从图片 {filename} 提取到支付金额: {payment_amount:.2f}")
        
        # 然后处理PDF文件
        for filename in os.listdir(self.folder_path):
            if filename.lower().endswith('.pdf'):
                file_path = os.path.join(self.folder_path, filename)
                pdf_info = self.extract_pdf_info(file_path)
                if pdf_info:
                    self.results.append(pdf_info)
        
        # 匹配支付图片和发票
        self.match_payment_to_invoice()

    def save_to_csv(self, output_file='analysis_results.csv'):
        """Save results to CSV file"""
        df = pd.DataFrame(self.results)
        
        # 重新排列列的顺序，确保actual_payment在price旁边
        columns = ['filename', 'price', 'actual_payment', 'invoice_number', 'invoice_date']
        df = df.reindex(columns=columns)
        
        # 添加支付差额列
        df['payment_difference'] = pd.to_numeric(df['actual_payment'], errors='coerce') - pd.to_numeric(df['price'], errors='coerce')
        
        # 重命名列
        df = df.rename(columns={
            'actual_payment': '实际支付金额',
            'price': '发票金额',
            'invoice_number': '发票号码',
            'invoice_date': '开票日期',
            'filename': '文件名',
            'payment_difference': '支付差额'
        })
        
        # 保存结果
        df.to_csv(output_file, index=False, encoding='utf-8-sig')
        print(f"结果已保存到 {output_file}")
        
        # 打印分析摘要
        print("\n分析摘要:")
        print("-" * 50)
        print(f"总计处理发票数: {len(df)}")
        print(f"成功匹配支付记录数: {df['实际支付金额'].notna().sum()}")
        print(f"未匹配支付记录数: {df['实际支付金���'].isna().sum()}")
        
        # 检查金额差异
        df['has_difference'] = abs(df['支付差额']) > 0.01
        differences = df[df['has_difference'] == True]
        if not differences.empty:
            print("\n发现金额不匹配的记录:")
            for _, row in differences.iterrows():
                print(f"发票 {row['文件名']}: 发票金额 {row['发票金额']}, 实际支付 {row['实际支付金额']}, 差额 {row['支付差额']:.2f}")

    def extract_product_name_from_filename(self, filename):
        """从文件名中提取中文关键字作为商品名称"""
        try:
            # 获取文件名（不含扩展名）
            base_name = os.path.splitext(os.path.basename(filename))[0]
            
            # 查找所有中文字符
            chinese_pattern = re.compile(r'[\u4e00-\u9fff]+')
            chinese_words = chinese_pattern.findall(base_name)
            
            if chinese_words:
                # 返回最长的中文词组
                return max(chinese_words, key=len)
            return None
        except Exception as e:
            self.logger.error(f"从文件名提取商品名称时出错: {str(e)}")
            return None

    def process_pdfs(self, input_dir):
        """处理目录下的所有PDF文件"""
        try:
            results = []
            pdf_files = []
            
            # 收集所有PDF文件
            for filename in os.listdir(input_dir):
                if filename.lower().endswith('.pdf'):
                    pdf_path = os.path.join(input_dir, filename)
                    pdf_files.append(pdf_path)
            
            if not pdf_files:
                self.logger.warning("未找到PDF文件")
                return
            
            self.logger.info(f"\n开始处理 {len(pdf_files)} 个PDF文件...")
            
            # 处理每个PDF文件
            for pdf_path in pdf_files:
                self.logger.info(f"\n处理PDF文件: {os.path.basename(pdf_path)}")
                info = self.extract_pdf_info(pdf_path)
                if info:
                    results.append(info)
            
            # 保存结果到CSV
            if results:
                df = pd.DataFrame(results)
                
                # 确保输出目录存在
                output_dir = 'output'
                if not os.path.exists(output_dir):
                    os.makedirs(output_dir)
                
                # 保存到invoice_results.csv
                output_file = os.path.join(output_dir, 'invoice_results.csv')
                df.to_csv(output_file, index=False, encoding='utf-8')
                self.logger.info(f"\n发票处理结果已保存到: {output_file}")
                self.logger.info(f"处理完成的PDF数量: {len(results)}")
                
                # 显示处理结果统计
                self.logger.info("\n处理结果统计:")
                self.logger.info(f"成功提取发票号码: {df['invoice_number'].notna().sum()}/{len(df)}")
                self.logger.info(f"成功提取开票日期: {df['invoice_date'].notna().sum()}/{len(df)}")
                self.logger.info(f"成功提取供应商: {df['supplier'].notna().sum()}/{len(df)}")
                self.logger.info(f"成功提取金额: {df['price'].notna().sum()}/{len(df)}")
                self.logger.info(f"成功提取商品名称: {df['product_name'].notna().sum()}/{len(df)}")
            else:
                self.logger.warning("没有成功提取的发票信息")
            
            # 合并PDF文件
            self.merge_pdfs(input_dir)
                
        except Exception as e:
            self.logger.error(f"处理PDF文件时出错: {str(e)}")
            traceback.print_exc()
            
    def merge_pdfs(self, input_dir):
        """合并所有PDF文件"""
        try:
            import PyPDF2
            
            # 收集所有PDF文件
            pdf_files = []
            for filename in os.listdir(input_dir):
                if filename.lower().endswith('.pdf'):
                    pdf_path = os.path.join(input_dir, filename)
                    pdf_files.append(pdf_path)
            
            if not pdf_files:
                self.logger.warning("未找到PDF文件可供合并")
                return
            
            # 确保输出目录存在
            output_dir = 'output'
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # 创建PDF合并器
            merger = PyPDF2.PdfMerger()
            
            # 按文件名排序，确保合并顺序一致
            pdf_files.sort()
            
            # 添加所有PDF文件
            for pdf_file in pdf_files:
                merger.append(pdf_file)
                self.logger.info(f"添加PDF文件: {os.path.basename(pdf_file)}")
            
            # 保存合并后的文件
            output_file = os.path.join(output_dir, f'merged_{datetime.now().strftime("%Y%m%d")}.pdf')
            merger.write(output_file)
            merger.close()
            
            self.logger.info(f"PDF文件已合并到: {output_file}")
            self.logger.info(f"合并的PDF数量: {len(pdf_files)}")
            
        except Exception as e:
            self.logger.error(f"合并PDF文件时出错: {str(e)}")
            traceback.print_exc()

    def save_combined_results(self):
        """保存合并后的结果到CSV文件"""
        try:
            # 定义所有需要的列
            required_columns = [
                '名称',              # 从文件名提取的中文关键字
                '品牌',              # 固定为NA
                '规格数量',          # 固定为1
                '规格单位',          # 固定为套
                '计量单位',          # 固定为套
                '数量',              # 固定为1
                '存放地点',          # 固定为科技楼1907
                '供应商',            # 从发票提取
                '发票号码',
                '开票日期',
                '发票金额',
                '实际支付金额',
                '差额',
                '发票文件',          # 发票文件名
                '文件名'             # 支付记录文件名
            ]
            
            # 读取之前保存的发票结果
            try:
                invoice_results = pd.read_csv(self.invoice_results_file, encoding='utf-8')
            except FileNotFoundError:
                self.logger.warning(f"未找到发票结果文件，创建新文件")
                invoice_results = pd.DataFrame(columns=['文件名'])
                
            # 读取支付记录
            try:
                payment_records = pd.read_csv(self.payment_records_file, encoding='utf-8')
            except FileNotFoundError:
                self.logger.warning(f"未找到支付记录文件，创建新文件")
                payment_records = pd.DataFrame(columns=['文件名', '实际支付金额'])
            
            # 确保列名一致
            self.logger.info("\n数据列名:")
            self.logger.info(f"发票结果列: {invoice_results.columns.tolist()}")
            self.logger.info(f"支付记录列: {payment_records.columns.tolist()}")
            
            # 合并数据
            combined_df = pd.merge(invoice_results, payment_records, on='文件名', how='outer')
            
            # 设置固定值的列
            combined_df['品牌'] = 'NA'
            combined_df['规格数量'] = 1
            combined_df['规格单位'] = '套'
            combined_df['计量单位'] = '套'
            combined_df['数量'] = 1
            combined_df['存放地点'] = '科技楼1907'
            
            # 重命名列
            column_mapping = {
                'invoice_number': '发票号码',
                'invoice_date': '开票日期',
                'price': '发票金额',
                'payment_amount': '实际支付金额',
                'supplier': '供应商',
                'product_name': '名称',
                'filename': '发票文件'
            }
            combined_df = combined_df.rename(columns=column_mapping)
            
            # 确保所有必需的列都存在
            for col in required_columns:
                if col not in combined_df.columns:
                    self.logger.info(f"创建缺失的列: {col}")
                    combined_df[col] = None
            
            # 计算差额（如果存在必要的列）
            if '发票金额' in combined_df.columns and '实际支付金额' in combined_df.columns:
                combined_df['差额'] = pd.to_numeric(combined_df['实际支付金额'], errors='coerce') - pd.to_numeric(combined_df['发票金额'], errors='coerce')
            
            # 获取所有可用的列
            available_columns = combined_df.columns.tolist()
            self.logger.info(f"合并后可用列: {available_columns}")
            
            # 重新排列列顺序
            final_columns = [col for col in required_columns if col in available_columns]
            combined_df = combined_df[final_columns]
            
            # 保存合并后的结果
            output_file = os.path.join(self.output_dir, 'combined_results.csv')
            combined_df.to_csv(output_file, index=False, encoding='utf-8')
            self.logger.info(f"合并结果已保存到: {output_file}")
            self.logger.info(f"最终列名: {final_columns}")
            
            # 打印分析摘要
            self.logger.info("\n分析摘要:")
            self.logger.info("-" * 50)
            self.logger.info(f"总计处理记录数: {len(combined_df)}")
            self.logger.info(f"成功匹配支付记录数: {combined_df['实际支付金额'].notna().sum()}")
            
            # 检查金额差异
            if '差额' in combined_df.columns:
                # 计算误差百分比
                combined_df['误差百分比'] = abs(combined_df['差额'] / combined_df['发票金额'] * 100)
                differences = combined_df[combined_df['误差百分比'] > 10].dropna(subset=['误差百分比'])
                
                if not differences.empty:
                    self.logger.info("\n发现金额不匹配的记录 (误差>10%):")
                    for _, row in differences.iterrows():
                        self.logger.info(f"文件 {row['文件名']}: "
                                       f"发票金额 {row.get('发票金额', 'N/A')}, "
                                       f"实际支付 {row.get('实际支付金额', 'N/A')}, "
                                       f"差额 {row.get('差额', 'N/A'):.2f}, "
                                       f"误差 {row.get('误差百分比', 'N/A'):.1f}%")
                
                # 删除临时列
                combined_df = combined_df.drop('误差百分比', axis=1)
            
        except Exception as e:
            self.logger.error(f"保存合并结果时出错: {str(e)}")
            traceback.print_exc()

def main():
    # Get the current directory
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Initialize analyzer
    analyzer = DocumentAnalyzer(current_dir)
    
    # Analyze documents
    analyzer.analyze_documents()
    
    # Save results
    analyzer.save_to_csv()

if __name__ == "__main__":
    main() 