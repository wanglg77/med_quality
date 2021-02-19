from define import *
from dept import *
from query import *
from assign import *
from dataprocess import *

@app.route('/', methods=['GET','POST'])
def index():
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)