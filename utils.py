# Common utils useful across most of 500k automation tools
import os
import zipfile
from shutil import copyfile

def copy_unzip_docx(f_path):
  # Copy docx and change file extension to *.zip
  f_path_zip = f_path.split(".")[0] + ".zip"
  copyfile(f_path, f_path_zip)

  # unzip docx (now zip) file
  dir_path = f_path.rsplit("\\")[0]
  zip_ref = zipfile.ZipFile(f_path_zip, 'r')
  zip_ref.extractall(dir_path)
  zip_ref.close()
  return dir_path

def find_pic_in_docx(directory):
    # Finds first pic from docx and returns path to image.
    for dirName, subdirList, fileList in os.walk(directory):
        print('Found directory: %s' % dirName)
        for fname in fileList:
            print('\t%s' % fname)
            if fname.lower().endswith(('.png', '.jpg', '.jpeg')):
                return dirName + "\\" + fname
