import json
import jwt
from flask import Flask, request
import os
from dotenv import load_dotenv
config = load_dotenv(".env")

flask_app = Flask(__name__)

SECRET_KEY = os.environ['SECRET_KEY_JWT']


def check_token(something):
    def wrap():
        try:
            token_passed = request.headers['Authorization']
            if request.headers['Authorization'] != '' and request.headers['Authorization'] != None:
                try:
                    data = jwt.decode(token_passed,SECRET_KEY, algorithms=['HS256'])
                    return something()
                except jwt.exceptions.ExpiredSignatureError:
                    return_data = {
                        "message": "El token ha expirado"
                        }
                    return flask_app.response_class(response=json.dumps(return_data), mimetype='application/json'),401
                except:
                    return_data = {
                        "message": "Token invalido"
                    }
                    return flask_app.response_class(response=json.dumps(return_data), mimetype='application/json'),401
            else:
                return_data = {
                    "message" : "Proporciona token",
                }
                return flask_app.response_class(response=json.dumps(return_data), mimetype='application/json'),401
        except Exception as e:
            return_data = {
                "message" : "Ha ocurrido un error"
                }
            return flask_app.response_class(response=json.dumps(return_data), mimetype='application/json'),500

    return wrap