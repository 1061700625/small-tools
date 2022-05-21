import os
import shutil
import zipfile
os.environ['UNRAR_LIB_PATH'] = r'./UnRAR64.dll'
from unrar import rarfile



def zip_uncompress(file_name: str, debugPrint=print):
    if not zipfile.is_zipfile(file_name):
        debugPrint('>> 不是有效的zip文件!')
        return
    re_decode = False
    store_path = os.path.dirname(file_name)
    with zipfile.ZipFile(file_name, 'r') as f:
        namelist = sorted(f.namelist())
        unzip_filepath = store_path + os.sep + namelist[0]
        try:
            unzip_filepath_gbk = unzip_filepath.encode('cp437').decode('gbk')
            re_decode = True
        except:
            unzip_filepath_gbk = unzip_filepath
        if os.path.exists(unzip_filepath_gbk):
            debugPrint('>> 文件夹已存在，请先删除!')
            return
        for fn in namelist:
            f.extract(fn, store_path)
            try:
                fn_gbk = fn.encode('cp437').decode('gbk')
                re_decode = True
            except:
                fn_gbk = fn
            fn = os.path.join(store_path, fn)
            fn_gbk = os.path.join(store_path, fn_gbk)
            os.rename(fn, fn_gbk)
    if os.path.exists(unzip_filepath) and re_decode and unzip_filepath.encode('utf-8').isalpha():
        debugPrint('删除临时文件')
        shutil.rmtree(unzip_filepath)
    os.startfile(unzip_filepath_gbk)
    debugPrint('>> 解压完成，已打开对应文件夹！')


def zip_compress(to_zip, save_zip_name=None, debugPrint=print):
    if not os.path.exists(to_zip):
        debugPrint('>> 文件夹不存在!')
        return None
    save_zip_name = save_zip_name or (to_zip + '.zip')
    # 1.先判断输出save_zip_name的上级是否存在(判断绝对目录)，否则创建目录
    # save_zip_name的上级目录
    save_zip_dir = os.path.split(os.path.abspath(save_zip_name))[0]
    if not os.path.exists(save_zip_dir):
        os.makedirs(save_zip_dir)
        debugPrint('>> 创建新目录%s' % save_zip_dir)
    f = zipfile.ZipFile(os.path.abspath(save_zip_name), 'w', zipfile.ZIP_DEFLATED)
    # 2.判断要被压缩的to_zip是否目录还是文件，是目录就遍历操作，是文件直接压缩
    # 如果不是目录,那就是文件
    if not os.path.isdir(os.path.abspath(to_zip)):
        # 判断文件是否存在
        if os.path.exists(os.path.abspath(to_zip)):
            f.write(to_zip)
            f.close()
            debugPrint('%s压缩为%s' % (to_zip, save_zip_name))
        else:
            debugPrint('%s文件不存在' % os.path.abspath(to_zip))
    else:
        # 判断目录是否存在，遍历目录
        if os.path.exists(os.path.abspath(to_zip)):
            zipList = []
            zipList.append(to_zip)
            # 遍历目录，加入列表
            for dir, subdirs, files in os.walk(to_zip):
                for fileItem in files:
                    zipList.append(os.path.join(dir, fileItem))
                for dirItem in subdirs:
                    zipList.append(os.path.join(dir, dirItem))
            # 读取列表压缩目录和文件
            for i in zipList:
                # replace是减少压缩文件的一层目录，即压缩文件不包括to_zip这个目录
                # os.path.dirname(to_zip)表示到上层目录
                # .encode('gbk').decode('cp437')
                f.write(i, i.replace(os.path.dirname(to_zip), ''))
            f.close()
            debugPrint('>> 压缩完成，路径是：' + save_zip_name)
        else:
            debugPrint('>> [%s]文件夹不存在' % os.path.abspath(to_zip))


def rar_uncompress(file_name: str, debugPrint=print):
    if not rarfile.is_rarfile(file_name):
        debugPrint('>> 不是有效的rar文件!')
        return False
    unzip_filepath = file_name.split('.')[0]
    if os.path.exists(unzip_filepath):
        debugPrint('>> 文件夹已存在!')
        # return
    else:
        os.mkdir(unzip_filepath)
    with rarfile.RarFile(file_name, 'r') as f:
        f.extractall(unzip_filepath)
    os.startfile(unzip_filepath)
    debugPrint('>> 解压完成，已打开对应文件夹！')
    return True


#rar_uncompress(r'C:\Users\SXF\Desktop\1.rar')
#zip_compress(r'C:\Users\SXF\Desktop\build\demo')
#zip_uncompress(r'C:\Users\SXF\Desktop\build\demo.zip')
#exit()

if __name__ == '__main__':
    while True:
        file = input('>> 请将文件拖入后按回车键: ')
        file = file.strip('"')
        if os.path.isdir(file):
            print('>> 是文件夹，即将进入压缩模式！')
            zip_compress(file)
        elif os.path.isfile(file):
            print('>> 是压缩包，即将进入解压模式！')
            if file.endswith('rar'):
                print('>> rar格式！')
                rar_uncompress(file)
            elif file.endswith('zip'):
                print('>> zip格式！')
                zip_uncompress(file)
        else:
            print('>> 输入非法！')
        print('\n')

    input('\n>> 任意键退出')
