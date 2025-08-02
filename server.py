from flask import Flask, render_template, request, send_file, jsonify
from flask_cors import CORS
from print_with_time_delay import SplitAndPrint as SplitAndPrint
app = Flask(__name__)
CORS(app)
# 初始化 SplitAndPrint 类
sp = SplitAndPrint(1)

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/file/split', methods=['POST'])
def split_sheets():
    # 参数提取
    excel_file = request.files['excelFile']
    sheet_name = request.form.get('sheetName')
    delay = int(request.form.get('timeDelay', 0))
    voice_enabled = request.form.get('voice') == 'true'
    output_dir = request.form.get('outputPath')
    print(f"Received parameters: sheet_name={sheet_name}, delay={delay}, voice_enabled={voice_enabled}, output_dir={output_dir}")
    if not output_dir:
        output_dir = './output'
    if not excel_file or not sheet_name:
        return jsonify({'code': -1, 'error': '缺少必要参数'}), 400
    try:
        sp.split(excel_file, sheet_name, output_dir)
        if voice_enabled:
            # 这里可以添加语音提示的逻辑
            pass
        # 压缩输出目录下的所有文件
        zip_file_path = f"{output_dir}/result.zip"
        sp.zip_output_files(output_dir)
        # 返回压缩文件
        return send_file(zip_file_path, as_attachment=True, download_name='拆分结果.zip')
        return jsonify({'code': 0, 'message': '文件拆分成功'})
    except Exception as e:
        error_msg = str(e)
        print(f"Error occurred: {error_msg}")
        return jsonify({'code': 1, 'error': str(e)}), 500
    
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080, debug=True)
