import docx2txt
from flask import Flask, render_template, request
from zipfile import ZipFile
from PIL import Image
import os
import ftplib
import datetime
import configparser

app = Flask(__name__)
app.secret_key = 'dmitriy'
tmp = os.curdir + '/tmp'


def check_found_file(filename):
    """
    Проверка существования файла
    """
    try:
        file = open(filename)
    except IOError as e:
        return False
    else:
        with file:
            return True


def sendToFtp():
    cfg = configparser.ConfigParser()
    if check_found_file('config.ini'):
         cfg.read('config.ini')
    else
        print('Отсутствует файл настроек!')
    host = cfg['FTP']['host']
    user = cfg['FTP']['user']
    passwd = cfg['FTP']['passwd']
    dsk = ftplib.FTP(
        host=host,
        user=user,
        passwd=passwd
    )
    path = str(datetime.date.today())
    dsk.cwd('ukdemidov.ru/www/files/img/')
    try:
        dsk.mkd(path)
    except ftplib.error_perm:
        print('Папка уже существует.')
    dsk.cwd(path)
    for filename in os.listdir(tmp):
        if filename.upper().find('SMALL') != -1:
            with open(tmp + '/' + filename, 'rb') as frb:
                dsk.storbinary("STOR " + filename, frb)
            print('Uploaded - ' + filename)
    dsk.close()


def parse():
    i = []
    if check_found_file(tmp + '/2.docx'):
        doc = docx2txt.process(tmp + '/2.docx')
    else
        print('Отсутствует файл docx!!!')
    for it in doc.splitlines():
        if it == '':
            continue
        elif it[:4] == 'Фото':
            photolist = it[5:].split(', ')
            photo = []
            for il in photolist:
                photo.append(
                    '/files/img/' +
                    str(datetime.date.today()) +
                    '/SMALL' + il.strip() + '.JPG'
                )
            i.append({'photo': photo, 'size': len(photolist)})
        else:
            i.append({'paragraph': it.rstrip()})
    return i


def resize_pics():
    for f in os.listdir(tmp):
        if f.upper().find('JPG') != -1:
            with open(tmp + '/' + f, 'rb') as of:
                im = Image.open(of)
                prop = im.width / im.height
                if prop >= 1:
                    width = 600
                    height = int(width / prop)
                else:
                    width = 300
                    height = int(width / prop)
                res = im.resize((width, height), resample=Image.ANTIALIAS)
                res.save(tmp + '/SMALL' + f.upper())


def clean_tmp():
    if os.path.exists(tmp):
        for f in os.listdir(tmp):
            os.remove(tmp + '/' + f)
    else:
        os.mkdir(tmp)


@app.route('/')
def index():
    clean_tmp()
    return render_template('index.html')


@app.route('/upload', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        clean_tmp()
        down = request.files['down']
        down.save(tmp + '/' + down.filename)
        with open(tmp + '/1.zip', 'rb') as of:
            z = ZipFile(of)
            z.extractall(tmp)
        resize_pics()
        sendToFtp()
        return render_template('result.html', parsed=parse())

app.run(host='0.0.0.0', port=5000, debug=True)
