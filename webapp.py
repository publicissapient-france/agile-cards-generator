__author__ = 'charles'

from flask import Flask
from flask import render_template
from flask import request
from flask import send_file
from cards_generator import generate_output_file
#from werkzeug import secure_filename


app = Flask(__name__)


@app.route('/', methods=['GET', 'POST'])
def home():
    if request.method == 'POST':
        file = request.files['file']
        output_file = generate_output_file(file)
        output_file.seek(0)
        return send_file(output_file, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         as_attachment=True, attachment_filename='output.xlsx')
    else:
        return render_template('home.html')


if __name__ == '__main__':
    app.run(debug=True)

