from flask import Flask, render_template, request, make_response
from calendar_utils import calendar_from_excel

app = Flask(__name__)


@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')


@app.route('/response/', methods=['GET', 'POST'])
def response():
    file=request.files['file']
    name=file.filename.rsplit('.', 2)[0]
    extension=file.filename.rsplit('.', 2)[1]
    if extension not in ['xls', 'xlsx']:
        return 'Invalid file type'
    cal=calendar_from_excel(file)
    print(str(cal))

    response = make_response(str(cal))
    cd = f'attachment; filename={name}.ics'
    response.headers['Content-Disposition'] = cd
    response.mimetype = 'text/plain'

    return response


if __name__ == '__main__':
    app.run(debug=True)
