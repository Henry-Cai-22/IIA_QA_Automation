from flask import Flask
from cred import *
from utils.utils import *
from routes.routes import routes 
import os

app = Flask(__name__)
app.secret_key = os.urandom(24).hex()  # Generate a random secret key
app.register_blueprint(routes)
    
if __name__ == '__main__':
    app.run(host='localhost', port=5000, debug=True)