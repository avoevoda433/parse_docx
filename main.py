import zipfile
import glob

dir_path = r'z:\Обменная\16.Программисты\Воевода\Отчеты по СМР\РФ\**\*.docx'
res = glob.glob(dir_path, recursive=True)


def get_photos(path_to_docx):
    archive = zipfile.ZipFile(path_to_docx)
    for file in archive.filelist:
        if file.filename.startswith('word/media/'):
            archive.extract(file, path='\\'.join(path_to_docx.split('\\')[:-1])+"\\"+path_to_docx.split('\\')[-1][:-5])


for path in res:
    print(path)
    get_photos(path)

print(len(res))
