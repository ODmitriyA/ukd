'''UKD'''

from zipfile import ZipFile
import os
import ftplib
import datetime
import configparser
from PIL import Image
from flask import Flask, render_template, request
import docx2txt

APP = Flask(__name__)
APP.secret_key = 'dmitriy'
TMP = os.curdir + '/tmp'


def check_found_file(filename):
    """
    Проверка существования файла
    """
    try:
        file = open(filename)
    except IOError as err:
        print(err)
        return False
    else:
        with file:
            return True


def send_to_ftp():
    '''
    Отправить файлы по FTP
    '''
    cfg = configparser.ConfigParser()
    if check_found_file('config.ini'):
        cfg.read('config.ini')
    else:
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
    for filename in os.listdir(TMP):
        if filename.upper().find('SMALL') != -1:
            with open(TMP + '/' + filename, 'rb') as frb:
                dsk.storbinary("STOR " + filename, frb)
            print('Uploaded - ' + filename)
    dsk.close()


def parse():
    '''
    Парсинг docx
    '''
    i = []
    if check_found_file(TMP + '/2.docx'):
        doc = docx2txt.process(TMP + '/2.docx')
    else:
        print('Отсутствует файл docx!!!')
    for line in doc.splitlines():
        if line == '':
            continue
        elif line[:4] == 'Фото':
            photolist = line[5:].split(', ')
            photos = []
            for photo in photolist:
                photos.append(
                    '/files/img/' +
                    str(datetime.date.today()) +
                    '/SMALL' + photo.strip() + '.JPG'
                )
            i.append({'photo': photos, 'size': len(photolist)})
        else:
            i.append({'paragraph': line.rstrip()})
    return i


def resize_pics():
    '''
    Resize pics
    '''
    for file in os.listdir(TMP):
        if file.upper().find('JPG') != -1:
            with open(TMP + '/' + file, 'rb') as imagefile:
                image = Image.open(imagefile)
                prop = image.width / image.height
                if prop >= 1:
                    width = 600
                    height = int(width / prop)
                else:
                    width = 300
                    height = int(width / prop)
                res = image.resize((width, height), resample=Image.ANTIALIAS)
                res.save(TMP + '/SMALL' + file.upper())


def clean_tmp():
    '''
    Очистка папки 'tmp'
    '''
    if os.path.exists(TMP):
        for file in os.listdir(TMP):
            os.remove(TMP + '/' + file)
    else:
        os.mkdir(TMP)


@APP.route('/')
def index():
    '''
    Отображение главной страницы
    '''
    clean_tmp()
    return render_template('index.html')


@APP.route('/upload', methods=['GET', 'POST'])
def upload():
    '''
    /upload
    '''
    if request.method == 'POST':
        clean_tmp()
        down = request.files['down']
        down.save(TMP + '/' + down.filename)
        with open(TMP + '/1.zip', 'rb') as zipfile:
            archive = ZipFile(zipfile)
            archive.extractall(TMP)
        resize_pics()
        send_to_ftp()
        return render_template('result.html', parsed=parse())


if __name__ == '__main__':
    APP.run(host='0.0.0.0', port=5000)
