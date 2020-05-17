#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Author  : Jay
# @Time    : 2020/5/12

import json
import sys
from modules import poster
from flask import Flask, request

app = Flask(__name__)

@app.route('/', methods=['POST'])
def run():
    #print(request.json) # dict
    poster.show(request.json)
    poster.pptx(request.json)
    poster.write(request.json)
    poster.read()
    return 'Successfully!'

if __name__ == "__main__":
    app.run(debug=True, threaded=False)