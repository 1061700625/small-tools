import os
import requests


def excel2image(filePath, outputType='JPG'):
    fileName = filePath.split(os.sep)[-1]
    url = 'https://api.products.aspose.app/cells/conversion/api/ConversionApi/Convert?outputType={}'.format(outputType)
    files = {"1": (fileName, open(filePath, "rb"))}
    res = requests.post(url=url, files=files).json()
    if res['StatusCode'] == 200:
        print('>> 文件转换完成')
        return res
    else:
        print('>> 文件转换出错：' + str(res))
        return None

def downloadFile(fileName, folderName):
    url = 'https://api.products.aspose.app/cells/conversion/api/Download/{}?file={}'.format(folderName, fileName)
    res = requests.get(url)
    if res.status_code == 200:
        with open(fileName, 'wb+') as f:
            f.write(res.content)
        print('>> 下载文件完成')
    else:
        print('>> 下载文件出错：'+res.text)

if __name__ == '__main__':
    res = excel2image('a.docx')
    print(res)
    if res:
        downloadFile(res['FileName'], res['FolderName'])

