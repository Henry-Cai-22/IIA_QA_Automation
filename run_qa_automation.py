from flask import Flask
from cred import *
from utils import *
from routes.routes import routes 

app = Flask(__name__)
app.secret_key = FLASK_SECRET_KEY
app.register_blueprint(routes)
    
if __name__ == '__main__':
    app.run(host='localhost', port=5000, debug=True)