from flask import Flask, request
import formatting2, datetime
app=Flask(__name__)

@app.route('/')
def index():
#    if 'date' in request.cookies:
#        entry_date = request.cookies.get('date')
#    else:
    entry_date = datetime.date.today()
    html = f'''<form action='/upload' method='POST'
    enctype = "multipart/form-data">
    <center><h1>This web helps tokenize your file atomatically.</h1><br>
    Select an excel/csv file to upload:
    <input type="file" name="file"><br><br>
    請輸入學生開學日期:
    <input type="date" name='day' value="{entry_date}"><br><br>
    <input type="submit"></center>
    </form>'''
    return html


@app.route('/upload', methods=['POST'])
def upload():
    f = request.files['file']
    entry_date = request.form['day']
#    response.set_cookie('date', entry_date)
#    print(entry_date)
    if f.filename != "":
        fn = f.filename
        f.save('static/' + fn)
        formatting2.main(entry_date, 'static/'+ fn)
#        fn = '學生基本資料上傳.xlsx'
        fn = '學生基本資料上傳.xlsx'
        return '<center>File <a href="static/{}">{}</a> is tokenized successfully</center>'.format(fn, fn)
    else:
        return 'Please select a file.'
