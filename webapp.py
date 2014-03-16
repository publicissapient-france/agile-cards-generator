__author__ = 'charles'

from flask import Flask
from flask import render_template
from flask import request
from flask import send_file
import logging
from cards_generator import generate_output_file
#from google.appengine.api.logservice import logservice
#from werkzeug import secure_filename


app = Flask(__name__)


@app.route('/', methods=['GET', 'POST'])
def home():
    logging.debug("Loading Home page")
    if request.method == 'POST':
        logging.info("Processing file ...")
        file = request.files['file']
        output_file = generate_output_file(file)
        output_file.seek(0)
        logging.info("Successfully processing file !")
        return send_file(output_file, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         as_attachment=True, attachment_filename='output.xlsx')
    else:
        return render_template('home.html')


@app.errorhandler(500)
def internal_error(error):
    logging.error("Error : %s" % error)
    return render_template('home.html', error=error)


if __name__ == '__main__':
    app.run(debug=True)

