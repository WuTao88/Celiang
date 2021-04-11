import os

dirs=open(input('文件夹数据:'), "r+",encoding='UTF-8')

for DIR in dirs:
    os.mkdir( DIR.strip(), 7777 );