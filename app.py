from flask import Flask, render_template, request, send_from_directory,session,make_response
from src.main import *
import os

upload_path = 'upload'  # 文件上传下载路径
app = Flask(__name__)
app.secret_key = 'xmindhellodfdf'

@app.route('/index',methods=['GET','POST'])
def index():
    session['tag'] = False
    if request.method == 'POST':
        res = {
            'error': '',
            'file_url': '',
            'sucess_msg': ''
        }
        # 删除upload下所有的文件(除__init__.py)
        del_files()
        # 获取xmind文件名
        xmind_file_obj = request.files['xmindFile']
        xmind_file_name = xmind_file_obj.filename
        # 判断xmind文件后缀是否正确
        if not xmind_file_name.endswith('.xmind'):
            res['error'] = '上传的 xmind 文件不正确！'
            return render_template('index.html', res = res)
        # 保存上传的文件
        save_path = os.path.join(upload_path, xmind_file_name)
        xmind_file_obj.save(save_path)
        # 生成excel文件名
        excel_file_name = xmind_file_name.rsplit(".", 1)[0] + '.xls'
        # 调用方法将xmind文件转换为excel文件
        get_xmind_content(save_path, os.path.join(upload_path, excel_file_name))
        res['file_url'] = os.path.join('/download/', excel_file_name)
        res['sucess_msg'] = xmind_file_name + ' 转换成功，点击下载用例！'
        return render_template('index.html', res = res)
    return render_template('index.html', res = {})

@app.route('/download/<filename>',methods=['GET'])
def download(filename):
    # 下载的文件路径
    excel_file_path = os.path.join(upload_path, filename)
    if request.method == "GET":
        if os.path.isfile(excel_file_path):
            return send_from_directory(upload_path, filename, as_attachment=True)

# 删除upload下所有的文件(除__init__.py)
def del_files():
    for file_name in os.listdir(upload_path):
        if file_name not in ['__init__.py']:
            os.remove(os.path.join(upload_path, file_name))

if __name__ == '__main__':
    app.run()


