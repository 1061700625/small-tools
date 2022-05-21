import os
import requests


def excel2image(filePath, outputType='JPG', debugPrint=print):
    fileName = filePath.split(os.sep)[-1]
    url = 'https://api.products.aspose.app/cells/conversion/api/ConversionApi/Convert?outputType={}'.format(outputType)
    files = {"1": (fileName, open(filePath, "rb"))}
    debugPrint('>> 开始请求文件转换中...')
    res = requests.post(url=url, files=files).json()
    if res['StatusCode'] == 200:
        debugPrint('>> 文件转换完成')
        return res
    else:
        debugPrint('>> 文件转换出错：' + str(res))
        return None

def downloadFile(fileName, folderName, debugPrint=print):
    url = 'https://api.products.aspose.app/cells/conversion/api/Download/{}?file={}'.format(folderName, fileName)
    res = requests.get(url)
    if res.status_code == 200:
        with open(fileName, 'wb+') as f:
            f.write(res.content)
        debugPrint('>> 文件下载完成: '+fileName)
        return True
    else:
        debugPrint('>> 文件下载出错：'+res.text)
        return False

if __name__ == '__main__':
    res = excel2image('a.docx')
    print(res)
    if res:
        downloadFile(res['FileName'], res['FolderName'])

