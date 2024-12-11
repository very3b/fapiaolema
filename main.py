import os
import logging
from pdf_image_analyzer import DocumentAnalyzer
from test_image_payment import PaymentImageTester
import pandas as pd
from datetime import datetime
import sys
import re

def clean_filename_for_name(filename):
    """从文件名提取商品名称（去掉数字和log字符）"""
    # 移除扩展名
    name = os.path.splitext(filename)[0]
    # 移除数字
    name = re.sub(r'\d+', '', name)
    # 移除log字符（不区分大小写）
    name = re.sub(r'log', '', name, flags=re.IGNORECASE)
    # 移除特殊字符
    name = re.sub(r'[^\w\s\u4e00-\u9fff]', '', name)
    # 移除多余的空格
    name = ' '.join(name.split())
    return name.strip()

def setup_logging():
    """设置日志配置"""
    # 创建输出目录
    output_dir = 'output'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # 设置日志格式
    log_format = '%(asctime)s - %(levelname)s - %(message)s'
    
    # 配置控制台输出
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(logging.Formatter(log_format))
    
    # 配置文件输出
    log_file = os.path.join(output_dir, f'process_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log')
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setFormatter(logging.Formatter(log_format))
    
    # 配置根日志记录器
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    logger.addHandler(console_handler)
    logger.addHandler(file_handler)
    
    return logger

def print_system_info():
    """打印系统信息"""
    logger = logging.getLogger(__name__)
    logger.info("=" * 50)
    logger.info("发票处理工具启动")
    logger.info("=" * 50)
    logger.info(f"运行时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info(f"工作目录: {os.getcwd()}")
    logger.info(f"Python版本: {sys.version}")
    logger.info(f"系统平台: {sys.platform}")
    try:
        import pytesseract
        version = pytesseract.get_tesseract_version()
        logger.info(f"Tesseract版本: {version}")
    except Exception as e:
        logger.error(f"Tesseract检查失败: {str(e)}")
    logger.info("=" * 50)

def main():
    # 设置日志
    logger = setup_logging()
    
    # 打印系统信息
    print_system_info()
    
    try:
        # 指定输入目录
        input_dir = '.'  # 当前目录
        
        # 创建输出目录
        output_dir = 'output'
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # 处理PDF发票
        logger.info("\n开始处理PDF发票...")
        pdf_analyzer = DocumentAnalyzer(input_dir)
        pdf_analyzer.process_pdfs(input_dir)
        
        # 处理支付截图
        logger.info("\n开始处理支付截图...")
        payment_tester = PaymentImageTester()
        payment_results = payment_tester.process_payment_images(input_dir)
        
        # 合并结果
        logger.info("\n开始合并处理结果...")
        
        # 读取发票处理结果
        invoice_results_file = os.path.join(output_dir, 'invoice_results.csv')
        try:
            invoice_results = pd.read_csv(invoice_results_file, encoding='utf-8')
            logger.info(f"读取到 {len(invoice_results)} 条发票记录")
            
            # 显示发票数据的列名
            logger.info(f"发票数据列名: {invoice_results.columns.tolist()}")
            
            # 显示前几条记录的内容
            logger.info("\n发票数据示例:")
            logger.info(invoice_results.head().to_string())
            
        except FileNotFoundError:
            invoice_results = pd.DataFrame()
            logger.warning("未找到发票处理结果文件")
        except Exception as e:
            logger.error(f"读取发票文件时出错: {str(e)}")
            invoice_results = pd.DataFrame()
        
        # 转换支付记录为DataFrame
        if payment_results:
            payment_df = pd.DataFrame(payment_results)
            logger.info(f"处理了 {len(payment_df)} 条支付记录")
            
            # 显示支付数据的列名
            logger.info(f"支付数据列名: {payment_df.columns.tolist()}")
            
            # 显示前几条记录的内容
            logger.info("\n支付数据示例:")
            logger.info(payment_df.head().to_string())
            
        else:
            payment_df = pd.DataFrame()
            logger.warning("未找到支付记录")
        
        # 准备合并结果
        combined_results = []
        
        # 遍历所有记录
        processed_files = set()
        
        # 处理有发票的记录
        if not invoice_results.empty:
            for _, invoice in invoice_results.iterrows():
                # 从文件名提取商品名称
                filename = invoice.get('filename', '')
                product_name = clean_filename_for_name(filename)
                
                record = {
                    '名称': product_name or invoice.get('product_name', ''),  # 优先使用清理后的文件名
                    '品牌': 'NA',
                    '规格数量': 1,
                    '规格单位': '套',
                    '计量单位': '套',
                    '数量': 1,
                    '存放地点': '科技楼1907',
                    '供应商': invoice.get('supplier', ''),
                    '发票号码': invoice.get('invoice_number', ''),
                    '开票日期': invoice.get('invoice_date', ''),
                    '发票金额': invoice.get('price', ''),
                    '发票文件': invoice.get('filename', ''),
                    '实际支付金额': None,
                    '文件名': None,
                    '差额': None
                }
                
                # 查找对应的支付记录
                if not payment_df.empty:
                    payment_match = payment_df[payment_df['发票文件'] == invoice['filename']]
                    if not payment_match.empty:
                        record['实际支付金额'] = payment_match.iloc[0]['实际支付金额']
                        record['文件名'] = payment_match.iloc[0]['文件名']
                        
                        # 计算差额
                        try:
                            invoice_amount = float(record['发票金额'])
                            payment_amount = float(record['实际支付金额'])
                            record['差额'] = payment_amount - invoice_amount
                        except (ValueError, TypeError):
                            record['差额'] = None
                
                combined_results.append(record)
                processed_files.add(invoice.get('filename', ''))
        
        # 处理没有发票的支付记录
        if not payment_df.empty:
            for _, payment in payment_df.iterrows():
                if payment['发票文件'] not in processed_files:
                    # 从文件名提取商品名称
                    filename = payment.get('文件名', '')
                    product_name = clean_filename_for_name(filename)
                    
                    record = {
                        '名称': product_name,  # 使用清理后的文件名
                        '品牌': 'NA',
                        '规格数量': 1,
                        '规格单位': '套',
                        '计量单位': '套',
                        '数量': 1,
                        '存放地点': '科技楼1907',
                        '供应商': '',
                        '发票号码': '',
                        '开票日期': '',
                        '发票金额': '',
                        '发票文件': payment['发票文件'],
                        '实际支付金额': payment['实际支付金额'],
                        '文件名': payment['文件名'],
                        '差额': None
                    }
                    combined_results.append(record)
        
        # 保存合并结果
        if combined_results:
            df = pd.DataFrame(combined_results)
            
            # 重新排列列顺序
            columns_order = [
                '名称', '品牌', '规格数量', '规格单位', '计量单位', '数量', 
                '存放地点', '供应商', '发票号码', '开票日期', '发票金额',
                '实际支付金额', '差额', '发票文件', '文件名'
            ]
            
            # 确保所有列都存在
            for col in columns_order:
                if col not in df.columns:
                    df[col] = None
            
            # 按指定顺序排列列
            df = df[columns_order]
            
            output_file = os.path.join(output_dir, 'combined_results.csv')
            df.to_csv(output_file, index=False, encoding='utf-8')
            logger.info(f"\n合并结果已保存到: {output_file}")
            logger.info(f"总记录数: {len(df)}")
            
            # 显示合并后的数据示例
            logger.info("\n合并后的数据示例:")
            logger.info(df.head().to_string())
            
            # 显示统计信息
            logger.info("\n处理结果统计:")
            logger.info(f"发票记录数: {len(invoice_results) if not invoice_results.empty else 0}")
            logger.info(f"支付记录数: {len(payment_df) if not payment_df.empty else 0}")
            logger.info(f"成功匹配数: {df['实际支付金额'].notna().sum()}")
            
            # 检查金额差异
            if '发票金额' in df.columns and '实际支付金额' in df.columns:
                matched_records = df[df['发票金额'].notna() & df['实际支付金额'].notna()]
                if not matched_records.empty:
                    matched_records['误差百分比'] = abs(pd.to_numeric(matched_records['差额'], errors='coerce') / 
                                                  pd.to_numeric(matched_records['发票金额'], errors='coerce') * 100)
                    mismatches = matched_records[matched_records['误差百分比'] > 10]
                    
                    if not mismatches.empty:
                        logger.warning("\n发现金额不匹配的记录（误差>10%）:")
                        for _, row in mismatches.iterrows():
                            logger.warning(f"发票: {row['发票文件']}, "
                                         f"发票金额: {row['发票金额']}, "
                                         f"支付金额: {row['实际支付金额']}, "
                                         f"差额: {row['差额']:.2f}, "
                                         f"误差: {row['误差百分比']:.1f}%")
        else:
            logger.warning("没有找到任何处理结果")
            
        logger.info("\n处理完成!")
        logger.info("=" * 50)
        
    except Exception as e:
        logger.error(f"程序执行出错: {str(e)}")
        logger.error("详细错误信息:", exc_info=True)
        raise

if __name__ == '__main__':
    main() 