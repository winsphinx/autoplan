#!/usr/bin/env python
# -- coding: utf-8 --

import base64
import getopt
import json
import sys


def dec(s):
    return str(base64.b64decode(bytes(s, 'utf-8')), 'utf-8')


def enc(s):
    return str(base64.b64encode(bytes(s, 'utf-8')), 'utf-8')


def add_user(user, pswd):
    data = read_data()
    data[user] = enc(pswd)
    write_data(data)


def del_user(user):
    data = read_data()
    del data[user]
    write_data(data)


def read_data():
    with open("password.json", 'r') as f:
        return json.load(f)


def write_data(data):
    with open("password.json", 'w') as f:
        json.dump(data, f)


tip = "\nUsage:\n\tTo add user:\tpswd -a username password\n\tTo delete user:\tpswd -d username\n"

try:
    opts, args = getopt.getopt(sys.argv[1:], "-h-a:-d:")

    if len(opts) == 0 and len(args) == 0:
        print(tip)
    else:
        for k, v in opts:
            if k == "-a":
                if len(args) == 1:
                    add_user(v, args[0])
                elif len(args) < 1:
                    print("password needed")
                else:
                    print("password too many")
            elif k == "-d":
                del_user(v)

except getopt.GetoptError:
    print(tip)
