from datetime import datetime
import os
import requests


from utils import log


EXPORT_FOLDER = './export'
PHOTO_LOGO_URL = 'https://drive.google.com/file/d/1hApCr3FnpZedkaJnQkRA4GRoeKY-HZce/view?usp=sharing'


class FS():
    def __init__(self, is_new_folder=False):
        if not os.path.isdir(EXPORT_FOLDER):
            os.mkdir(EXPORT_FOLDER)
            log(f'The folder "{EXPORT_FOLDER}" is created', 'res')
        if is_new_folder:
            dt = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
            self.folder = os.path.join(EXPORT_FOLDER, 'export_' + dt)
            os.mkdir(self.folder)
            log(f'The folder "{self.folder}" is created', 'res')
        else:
            files = os.listdir(EXPORT_FOLDER)
            files = [f for f in files if f.startswith('export_')]
            files.sort()
            if len(files) == 0:
                log('Can not find the last "export_*" folder', 'err')
            self.folder = os.path.join(EXPORT_FOLDER, files[-1])
            log(f'The existing folder "{self.folder}" will be used', 'res')

        tmp_dir = self.get_path('tmp')
        if not os.path.isdir(tmp_dir):
            os.mkdir(tmp_dir)

        self.downloaded = []

    def get_path(self, file_name):
        return os.path.join(self.folder, file_name)

    def download_photo(self, name, url):
        if not name or not url:
            return
        file_path = self.get_path(f'tmp/{name}.jpg')
        if not file_path in self.downloaded:
            gdrive_download(url.split('/')[-2], file_path)
            self.downloaded.append(file_path)
            log(f'Photo of person "{name}" is downloaded', 'res')
        return file_path

    def download_photo_logo(self):
        file_path = self.get_path('tmp/cait.jpg')
        if not file_path in self.downloaded:
            gdrive_download(PHOTO_LOGO_URL.split('/')[-2], file_path)
            self.downloaded.append(file_path)
            log('Logo photo is downloaded', 'res')
        return file_path


def gdrive_download(uid, destination, chunk_size=32768):
    URL = "https://docs.google.com/uc?export=download"

    session = requests.Session()
    response = session.get(URL, params = { 'id' : uid }, stream = True)

    for key, value in response.cookies.items():
        if key.startswith('download_warning'):
            params = { 'id' : uid, 'confirm' : value }
            response = session.get(URL, params = params, stream = True)
            break

    with open(destination, 'wb') as f:
        for chunk in response.iter_content(chunk_size):
            if chunk:
                f.write(chunk)
