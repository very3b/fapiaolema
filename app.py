import os
import webview
from flask import Flask, render_template, request, jsonify
from datetime import datetime
from pdf_image_analyzer import DocumentAnalyzer
from test_image_payment import PaymentImageTester
import logging
import threading
from queue import Queue
import traceback

# 创建Flask应用
app = Flask(__name__)

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# 创建消息队列用于存储日志
log_queue = Queue()

class LogHandler(logging.Handler):
    def emit(self, record):
        log_entry = self.format(record)
        log_queue.put(log_entry)

# 添加自定义日志处理器
logger.addHandler(LogHandler())

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/select-folder')
def select_folder():
    try:
        folder_path = webview.windows[0].create_file_dialog(
            webview.FOLDER_DIALOG
        )
        if folder_path and len(folder_path) > 0:
            selected_path = folder_path[0]
            return jsonify({'status': 'success', 'path': selected_path})
        return jsonify({'status': 'cancelled'})
    except Exception as e:
        logger.error(f"选择文件夹时出错: {str(e)}")
        return jsonify({'status': 'error', 'message': str(e)})

@app.route('/process', methods=['POST'])
def process_files():
    try:
        folder_path = request.json.get('folder_path')
        if not folder_path:
            return jsonify({'status': 'error', 'message': '请选择文件夹'})
            
        # 清空日志队列
        while not log_queue.empty():
            log_queue.get()
            
        # 开始处理
        logger.info(f"开始处理文件夹: {folder_path}")
        
        # 创建处理器实例
        analyzer = DocumentAnalyzer(folder_path)
        tester = PaymentImageTester()
        
        # 处理PDF文件
        logger.info("开始处理PDF文件...")
        analyzer.analyze_documents()
        
        # 处理图片文件
        logger.info("开始处理图片文件...")
        tester.process_payment_images(folder_path)
        
        # 生成合并文件
        date_str = datetime.now().strftime("%Y%m%d")
        
        # 合并PDF
        merged_pdf = os.path.join(folder_path, f'merged_{date_str}.pdf')
        analyzer.merge_pdfs(merged_pdf)
        
        # 合并支付截图
        merged_log = os.path.join(folder_path, f'merged_{date_str}_log.pdf')
        analyzer.merge_images(merged_log)
        
        # 生成数据对应关系文件
        analyzer.save_combined_results()
        
        logger.info("处理完成!")
        return jsonify({'status': 'success'})
        
    except Exception as e:
        logger.error(f"处理文件时出错: {str(e)}")
        traceback.print_exc()
        return jsonify({'status': 'error', 'message': str(e)})

@app.route('/get-logs')
def get_logs():
    logs = []
    while not log_queue.empty():
        logs.append(log_queue.get())
    return jsonify({'logs': logs})

def run_server():
    app.run(port=5000)

def main():
    # 启动Flask服务器
    threading.Thread(target=run_server, daemon=True).start()
    
    # 创建webview窗口
    webview.create_window(
        '发票处理工具',
        'http://localhost:5000',
        width=800,
        height=600,
        resizable=True
    )
    webview.start()

if __name__ == '__main__':
    main() 