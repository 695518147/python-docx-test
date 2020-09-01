#!/usr/bin/python
# -*- coding:UTF-8 -*-
'''
Author: zhangpeiyu
Date: 2020-09-01 21:51:45
LastEditTime: 2020-09-01 22:04:05
Description: 我不是诗人，所以，只能够把爱你写进程序，当作不可解的密码，作为我一个人知道的秘密。
'''
import json

def getJson(file_path):
    with open(file_path, 'r') as fd:
        jo = json.load(fd, encoding='utf-8')
    return jo

if '__main__' == __name__:
    print(getJson("api.json"))