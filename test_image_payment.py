import os
import re
import cv2
import numpy as np
import pandas as pd
import pytesseract
from pytesseract import Output
import traceback
from PIL import Image
from datetime import datetime
import logging

class PaymentImageTester:
    def __init__(self, min_font_height=20):
        self.min_font_height = min_font_height
        self.logger = logging.getLogger(__name__)
        
        # 设置Tesseract路径
        if os.name == 'nt':  # Windows
            tesseract_paths = [
                r'C:\Program Files\Tesseract-OCR\tesseract.exe',
                r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
                r'D:\Program Files\Tesseract-OCR\tesseract.exe',
                r'D:\Program Files (x86)\Tesseract-OCR\tesseract.exe'
            ]
            
            # 查找可用的Tesseract路径
            tesseract_found = False
            for path in tesseract_paths:
                if os.path.exists(path):
                    pytesseract.pytesseract.tesseract_cmd = path
                    tesseract_dir = os.path.dirname(path)
                    os.environ['PATH'] = tesseract_dir + os.pathsep + os.environ['PATH']
                    os.environ['TESSDATA_PREFIX'] = os.path.join(tesseract_dir, 'tessdata')
                    tesseract_found = True
                    self.logger.info(f"找到Tesseract: {path}")
                    
                    # 验证语言包
                    tessdata_dir = os.path.join(tesseract_dir, 'tessdata')
                    eng_traineddata = os.path.join(tessdata_dir, 'eng.traineddata')
                    chi_sim_traineddata = os.path.join(tessdata_dir, 'chi_sim.traineddata')
                    
                    self.logger.info(f"检查语言包:")
                    self.logger.info(f"tessdata目录: {tessdata_dir}")
                    self.logger.info(f"英文语言包: {'存在' if os.path.exists(eng_traineddata) else '不存在'}")
                    self.logger.info(f"中文语言包: {'存在' if os.path.exists(chi_sim_traineddata) else '不存在'}")
                    
                    if not os.path.exists(eng_traineddata):
                        self.logger.error("未找到英文语言包!")
                    if not os.path.exists(chi_sim_traineddata):
                        self.logger.error("未找到中文语言包!")
                    
                    break
            
            if not tesseract_found:
                self.logger.error("未找到Tesseract安装，请确保已正确安装Tesseract")
                raise FileNotFoundError("Tesseract未找到")
            
            # 验证版本
            try:
                version_output = os.popen(f'"{pytesseract.pytesseract.tesseract_cmd}" --version').read()
                self.logger.info(f"Tesseract版本信息:\n{version_output}")
                
                # 从版本输出中提取版本号
                version_match = re.search(r'tesseract\s+(\d+\.\d+)', version_output)
                if version_match:
                    version = float(version_match.group(1))
                    if version < 4.0:
                        self.logger.warning(f"Tesseract版本 {version} 较旧，建议升级到5.0+版本")
                
            except Exception as e:
                self.logger.error(f"获取Tesseract版本失败: {str(e)}")
    
    def extract_payment_from_image(self, image_path):
        """从图片中提取支付金额"""
        try:
            # 读取图片
            img = cv2.imread(image_path)
            if img is None:
                self.logger.error(f"无法读取图片: {image_path}")
                return None
            
            # 转换为灰度图
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            
            # 增加对比度
            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
            enhanced = clahe.apply(gray)
            
            # 创建不同的图像处理版本
            images = []
            
            # 原始灰度图
            images.append(("原始灰度图", gray))
            
            # CLAHE增强
            images.append(("CLAHE增强", enhanced))
            
            # Otsu二值化
            _, binary_otsu = cv2.threshold(enhanced, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            images.append(("Otsu二值化", binary_otsu))
            
            # 自适应二值化
            binary_adaptive = cv2.adaptiveThreshold(enhanced, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)
            images.append(("自适应二值化", binary_adaptive))
            
            # 定义不同的PSM模式
            psm_modes = [
                3,   # 自动页面分割，但没有OSD（默认）
                4,   # 假设有一列可变大小的文本
                6,   # 假设为统一的文本块
                7,   # 将图像视为单行文本
                8,   # 将图像视为单词
                11,  # 稀疏文本，需要尽可能多地找到文本
                12,  # 稀疏文本和OSD
                13   # 将图像视为单行文本，不进行任何预处理/OSD
            ]
            
            all_results = []
            
            # 对每个图像版本尝试不同的PSM模式
            for img_name, img_version in images:
                for psm in psm_modes:
                    try:
                        # 配置Tesseract
                        config = f'--oem 3 --psm {psm}'
                        
                        # 进行OCR识别
                        text = pytesseract.image_to_string(img_version, lang='eng', config=config)
                        
                        # 查找负数金额的不同模式
                        amount_patterns = [
                            r'-\d+\.\d{2}',           # 标准格式：-xx.xx
                            r'[-—]\s*\d+\.\d{2}',     # 带空格：- xx.xx
                            r'[-—]\d+\.\d{2}',        # 不同的负号：−xx.xx
                            r'\(\d+\.\d{2}\)',        # 括号格式：(xx.xx)
                            r'支付\s*[-—]?\s*\d+\.\d{2}',  # 带"支付"的格式
                            r'¥\s*[-—]?\d+\.\d{2}',   # 带"¥"的格式
                            r'实付\s*[-—]?\s*\d+\.\d{2}',  # 带"实付"的格式
                            r'付款\s*[-—]?\s*\d+\.\d{2}',  # 带"付款"的格式
                            r'总计\s*[-—]?\s*\d+\.\d{2}'   # 带"总计"的格式
                        ]
                        
                        for pattern in amount_patterns:
                            matches = re.findall(pattern, text)
                            if matches:
                                for match in matches:
                                    try:
                                        # 处理括号格式
                                        if match.startswith('(') and match.endswith(')'):
                                            amount = -float(match[1:-1])
                                        else:
                                            # 替换所有可能的负号为标准负号
                                            cleaned_match = match.replace('—', '-').replace('−', '-').replace(' ', '')
                                            # 提取数字部分
                                            num_match = re.search(r'-?\d+\.\d{2}', cleaned_match)
                                            if num_match:
                                                amount = float(num_match.group())
                                            else:
                                                continue
                                        
                                        if amount < 0:  # 确保是负数
                                            self.logger.info(f"找到支付金额: {amount} (图像处理: {img_name}, PSM: {psm})")
                                            all_results.append((abs(amount), img_name, psm))
                                    except ValueError:
                                        continue
                    
                    except Exception as e:
                        self.logger.error(f"OCR处理失败 (图像处理: {img_name}, PSM: {psm}): {str(e)}")
                        continue
            
            # 如果找到结果，返回最常见的金额
            if all_results:
                # 统计每个金额出现的次数
                amount_counts = {}
                for amount, img_name, psm in all_results:
                    amount_str = f"{amount:.2f}"
                    if amount_str not in amount_counts:
                        amount_counts[amount_str] = 0
                    amount_counts[amount_str] += 1
                
                # 找出出现次数最多的金额
                most_common_amount = max(amount_counts.items(), key=lambda x: x[1])
                amount = float(most_common_amount[0])
                
                self.logger.info(f"最终选择的支付金额: {amount:.2f} (出现次数: {most_common_amount[1]})")
                return amount
            
            self.logger.warning(f"未找到支付金额")
            return None
                
        except Exception as e:
            self.logger.error(f"处理图片时出错: {str(e)}")
            traceback.print_exc()
            return None
            
    def process_payment_images(self, input_dir):
        """处理目录下的所有支付截图，返回支付记录列表"""
        try:
            results = []
            
            # 验证Tesseract版本
            try:
                version = pytesseract.get_tesseract_version()
                self.logger.info(f"\nTesseract版本: {version}")
            except Exception as e:
                self.logger.warning(f"无法获取Tesseract版本: {str(e)}")
            
            # 创建输出目录
            output_dir = 'output'
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # 处理所有图片
            for filename in os.listdir(input_dir):
                if 'log' in filename.lower() and filename.lower().endswith(('.jpg', '.jpeg', '.png')):
                    image_path = os.path.join(input_dir, filename)
                    self.logger.info(f"\n正在处理图片：{image_path}")
                    amount = self.extract_payment_from_image(image_path)
                    
                    if amount is not None:
                        # 尝试从文件名中找到对应的发票文件
                        base_name = filename.lower().replace('_log', '').replace('.jpg', '').replace('.jpeg', '').replace('.png', '')
                        invoice_file = None
                        
                        # 查找可能的发票文件
                        for inv_file in os.listdir(input_dir):
                            if inv_file.lower().endswith('.pdf'):
                                inv_base = inv_file.lower().replace('.pdf', '')
                                if base_name in inv_base or inv_base in base_name:
                                    invoice_file = inv_file
                                    break
                        
                        results.append({
                            '文件名': filename,
                            '实际支付金额': f"{amount:.2f}",
                            '发票文件': invoice_file
                        })
            
            # 显示处理结果统计
            if results:
                self.logger.info("\n支付金额提取统计:")
                self.logger.info(f"总计处理图片: {len(results)}张")
                self.logger.info(f"成功提��金额: {len(results)}个")
                
                # 显示提取的金额
                self.logger.info("\n提取的支付金额:")
                for result in results:
                    self.logger.info(f"文件: {result['文件名']}, "
                                   f"金额: {result['实际支付金额']}, "
                                   f"对应发票: {result.get('发票文件', 'N/A')}")
                
                # 合并所有支付截图为PDF
                merged_log_pdf = os.path.join(output_dir, f'merged_{datetime.now().strftime("%Y%m%d")}_log.pdf')
                self.merge_images_to_pdf(input_dir, merged_log_pdf)
            else:
                self.logger.warning("没有找到有效的支付记录")
            
            return results
            
        except Exception as e:
            self.logger.error(f"处理支付图片时出错: {str(e)}")
            traceback.print_exc()
            return []
            
    def merge_images_to_pdf(self, input_dir, output_pdf):
        """将支付截图合并为一个PDF文件"""
        try:
            images = []
            # 按文件名排序，确保合并顺序一致
            filenames = sorted([f for f in os.listdir(input_dir) 
                              if 'log' in f.lower() and f.lower().endswith(('.jpg', '.jpeg', '.png'))])
            
            for filename in filenames:
                image_path = os.path.join(input_dir, filename)
                img = Image.open(image_path)
                if img.mode == 'RGBA':
                    img = img.convert('RGB')
                images.append(img)
            
            if images:
                # 确保输出目录存在
                output_dir = os.path.dirname(output_pdf)
                if not os.path.exists(output_dir):
                    os.makedirs(output_dir)
                
                # 保存合并的PDF
                images[0].save(output_pdf, save_all=True, append_images=images[1:])
                self.logger.info(f"支付截图已合并到: {output_pdf}")
                self.logger.info(f"合并的图片数量: {len(images)}")
            else:
                self.logger.warning("没有找到支付截图可供合并")
        except Exception as e:
            self.logger.error(f"合并支付截图到PDF时出错: {str(e)}")
            traceback.print_exc()

def main():
    # 设置日志级别
    logging.basicConfig(level=logging.INFO,
                       format='%(asctime)s - %(levelname)s - %(message)s')
    
    # 创建测试实例
    tester = PaymentImageTester()
    
    # 指定输入目录
    input_dir = '.'  # 当前目录
    
    # 处理图片
    results = tester.process_payment_images(input_dir)
    
    # 显示处理结果统计
    if results:
        self.logger.info("\n支付金额提取统计:")
        self.logger.info(f"总计处理图片: {len(results)}张")
        self.logger.info(f"成功提取金额: {len(results)}个")
        
        # 显示提取的金额
        self.logger.info("\n提取的支付金额:")
        for result in results:
            self.logger.info(f"文件: {result['文件名']}, "
                           f"金额: {result['实际支付金额']}, "
                           f"对应发票: {result.get('发票文件', 'N/A')}")
        
        # 合并所有支付截图为PDF
        if not os.path.exists('output'):
            os.makedirs('output')
        merged_log_pdf = os.path.join('output', f'merged_{datetime.now().strftime("%Y%m%d")}_log.pdf')
        self.merge_images_to_pdf(input_dir, merged_log_pdf)
    else:
        self.logger.warning("没有找到有效的支付记录")

if __name__ == '__main__':
    main() 