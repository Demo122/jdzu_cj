import subprocess
import time

from flask import Flask, request, render_template, send_from_directory
from pyquery import PyQuery as py
import requests
import execjs
import xlwt
import pdfkit
import send_email

app = Flask(__name__)

# 构建一个session对象来保存cookie
sess = requests.Session()
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.149 Safari/537.36"
}


@app.route('/', methods=['GET', 'POST'])
def login():
    if request.method == "GET":
        getVerifyCode(sess)
        return render_template("index.html", val1=time.time())
    # 由于缓存的问题，图片覆盖了，但是前端页面刷新还是以前的旧图。
    # 需要_t参数，使每次请求的数据不一样，才能刷新图片
    else:
        try:
            username = request.form.get("username")
            password = request.form.get("password")
            verifycode = request.form.get("verifycode")
            file_format = request.form.get("file_format")
            msg = jdzuLogin(username, password, verifycode, file_format)
            # 将session里的cookies清除，使得下一位使用者可以登录
            sess.cookies.clear()
            if msg:
                return "<script>alert('{msg}');window.history.back(-1);</script>".format(msg=msg)
        except Exception:
            return "<script>alert('信息输入错误！');window.history.back(-1);</script>"
    if file_format == "excel":
        return send_from_directory(r"excel/", filename="{name}.xls".format(name=username), as_attachment=True)
    if file_format == "pdf":
        # print("返回pdf")
        return send_from_directory(r"pdf/", filename="{name}.pdf".format(name=username), as_attachment=True)


def getEncoded(userAccount, userPassword):
    # 读取js文件
    with open('static/encode.js', encoding='utf-8') as f:
        js = f.read()
    # 通过compile命令转成一个js对象
    docjs = execjs.compile(js)
    # 调用function方法
    account = docjs.call("encodeInp", userAccount)
    password = docjs.call("encodeInp", userPassword)
    encoded = account + "%%%" + password
    # print(encoded)
    return encoded


def getVerifyCode(sess):
    # 验证码请求url。获取验证码
    verifycode_url = "http://61.131.228.75:8080/jsxsd/verifycode.servlet"
    verifycode_data = sess.get(verifycode_url, headers=headers).content
    with open("static/verifycode.jpg", "wb") as f:
        f.write(verifycode_data)


def jdzuLogin(username, password, verifycode, file_format):
    # 获取js编码后的encoded
    encoded = getEncoded(username, password)
    # 准备表单数据
    data = {
        "userAccount": username,
        "userPassword": "",
        "RANDOMCODE": verifycode,
        "encoded": encoded
    }
    # 模拟登录
    r = py(sess.post('http://61.131.228.75:8080/jsxsd/xk/LoginToXk', data=data, headers=headers).text)
    # 如果验证码错误
    if r('#showMsg').text():
        return r('#showMsg').text()
    # 请求首页，获取用户名字
    rep = py(sess.get('http://61.131.228.75:8080/jsxsd/framework/xsMain.jsp', headers=headers).text)
    name = rep('#btn_gotoGrzx .glyphicon-class').text() + "的成绩单"
    # 请求成绩单数据
    response = py(sess.get('http://61.131.228.75:8080/jsxsd/kscj/cjcx_list', headers=headers).text)

    # 将成绩解析到列表组，方便生产excel
    scores = list()
    for item in response('tr').items():
        score = list()
        if item('th').items():
            # 表头
            for th in item('th').items():
                score.append(th.text())
        if item('td').items():
            # tbody
            for td in item('td').items():
                score.append(td.text())
        scores.append(score)

    # 根据用户选择的格式保存
    if file_format == "excel":
        # 保存到excel
        save_excel(scores, name, username)
    if file_format == "pdf":
        # 保存到pdf
        save_excel_for_pdf(scores, name, username)
        convert_to_pdf(username)
    # 发送邮件通知
    send_email.mail(name)


# 保存到excel
def save_excel(scores, name, filename):
    # 创建工作workbook
    workbook = xlwt.Workbook(encoding='utf-8')
    # 创建工作表worksheet,填入表名
    worksheet = workbook.add_sheet('score', cell_overwrite_ok=True)

    # 创建一个样式对象，初始化样式
    style = xlwt.XFStyle()
    al = xlwt.Alignment()
    al.horz = 0x02  # 设置水平居中
    al.vert = 0x01  # 设置垂直居中
    style.alignment = al

    # 在表中写入相应的数据
    # 行数
    row = len(scores)
    # 列数
    clo = len(scores[0])
    # 合并第一行 加标题
    worksheet.write_merge(0, 0, 0, 13, name, style)
    for i in range(1, row + 1):
        for j in range(clo):
            worksheet.write(i, j, scores[i - 1][j])
    # 保存表
    workbook.save('excel/{filename}.xls'.format(filename=filename))


# 为保存到pdf做准备，先保存为修改后的excel
def save_excel_for_pdf(scores, name, filename):
    # 创建工作workbook
    workbook = xlwt.Workbook(encoding='utf-8')
    # 创建工作表worksheet,填入表名
    worksheet = workbook.add_sheet('score', cell_overwrite_ok=True)
    # 设置单元格大小，让所有元素都能完整显示在pdf
    # xlwt中是行和列都是从0开始计算的
    first_col_0 = worksheet.col(0)
    first_col_1 = worksheet.col(1)
    first_col_2 = worksheet.col(2)
    first_col_3 = worksheet.col(3)
    first_col_4 = worksheet.col(4)
    first_col_5 = worksheet.col(5)
    first_col_6 = worksheet.col(6)
    first_col_7 = worksheet.col(7)
    # 设置宽度
    first_col_0.width = 256 * 4
    first_col_1.width = 256 * 29
    first_col_2.width = 256 * 8
    first_col_3.width = 256 * 6
    first_col_4.width = 256 * 6
    first_col_5.width = 256 * 6
    first_col_6.width = 256 * 9
    first_col_7.width = 256 * 13
    # 创建一个样式对象，初始化样式
    style = xlwt.XFStyle()
    al = xlwt.Alignment()
    al.horz = 0x02  # 设置水平居中
    al.vert = 0x01  # 设置垂直居中
    style.alignment = al

    # 处理成绩，删除不必要的列, 注意，每移除一个，剩下的序列就变了，故重新推算要删除的序列
    # subject = [1, 2, 5, 9, 10, 11]
    subject = [1, 1, 3, 6, 6, 6]
    # 行数
    row = len(scores)
    # 列数
    clo = len(scores[0])
    for i in range(row):
        for j in subject:
            scores[i].pop(j)

    # 在表中写入相应的数据
    # 合并第一行 加标题
    # 行数
    row = len(scores)
    # 列数
    clo = len(scores[0])
    worksheet.write_merge(0, 0, 0, 7, name, style)
    for i in range(1, row + 1):
        for j in range(clo):
            worksheet.write(i, j, scores[i - 1][j])
    # 保存到pdf文件夹下
    workbook.save('pdf/{filename}.xls'.format(filename=filename))


# 调用linux下的命令，用libreoffice生产pdf文件
def convert_to_pdf(username):
    str = "libreoffice6.3 --convert-to pdf:writer_pdf_Export ./pdf/{name}.xls --outdir ./pdf/".format(name=username)
    subprocess.call(str, shell=True)
    # print("保存到pdf.......")


if __name__ == "__main__":
    app.run(host='127.0.0.1', port=5050)
