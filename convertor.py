import os
import requests


class Convertor:
    def __init__(self):
        self.download_urls = {
            'cells': r'https://api.products.aspose.app/cells/conversion/api/Download/{}?file={}',
            'words': r'https://products.aspose.app/words/zh/conversion/Download?id=',
            'doc2image': r'https://softparade.net/download/file?id=',
        }
        pass

    def downloadFile(self, param_map, types='cells', debugPrint=print):
        fileName = param_map['fileName']
        storePath = param_map['storePath']
        if types == 'cells':
            url = self.download_urls[types].format(param_map['folderName'], fileName)
        elif types == 'words':
            url = self.download_urls[types] + param_map['id']
        elif types == 'doc2image':
            url = self.download_urls[types] + param_map['id']
            fileName = fileName.split('.')[0] + '.zip'
        else:
            return False
        debugPrint('>> 开始请求文件下载')
        res = requests.get(url)
        if res.status_code == 200:
            with open(os.path.join(storePath, fileName), 'wb') as f:
                f.write(res.content)
            debugPrint('>> 文件下载完成: ' + fileName)
            return True
        else:
            debugPrint('>> 文件下载出错：' + res.text)
            return False

    def excel2image(self, filePath, outputType='JPG', debugPrint=print):
        fileName = filePath.split(os.sep)[-1]
        storePath = os.path.dirname(filePath)
        url = 'https://api.products.aspose.app/cells/conversion/api/ConversionApi/Convert?outputType={}'.format(
            outputType)
        files = {"1": (fileName, open(filePath, "rb"))}
        debugPrint('>> 开始请求文件转换中...')
        res = requests.post(url=url, files=files).json()
        if res['StatusCode'] == 200:
            debugPrint('>> 文件转换完成')
            param_map = {'fileName': res['FileName'], 'folderName': res['FolderName'], 'storePath': storePath}
            self.downloadFile(param_map, types='cells', debugPrint=debugPrint)
            return res
        else:
            debugPrint('>> 文件转换出错：' + str(res))
            return None


    def doc2docx(self, filePath, debugPrint=print):
        return self.word2image(filePath, outputType='DOCX', debugPrint=debugPrint)

    def word2image(self, filePath, outputType='JPG', debugPrint=print):
        """
        参考：https://products.aspose.app/words/zh/conversion/doc-to-docx
        注意：中文docx 转 image 容易出现格式错误
        """
        url = r'https://products.aspose.app/words/zh/conversion/api/convert?outputType={}&useOcr=false&locale=zh'.format(outputType)
        fileName = filePath.split(os.sep)[-1]
        storePath = os.path.dirname(filePath)
        files = {"1": (fileName, open(filePath, "rb"))}
        debugPrint('>> 开始请求文件转换中...')
        res = requests.post(url=url, files=files).json()
        if res['id']:
            debugPrint('>> 文件转换完成')
            param_map = {'fileName': res['id'].split('/')[-1], 'id': res['id'], 'storePath': storePath}
            self.downloadFile(param_map, types='words', debugPrint=debugPrint)
            return res
        else:
            debugPrint('>> 文件转换出错：' + str(res))
            return None


    def doc2image(self, filePath, outputType='JPG', debugPrint=print):
        """
        https://word2jpg.com/cn
        """
        fileName = filePath.split(os.sep)[-1]
        storePath = os.path.dirname(filePath)
        def download(id):
            url = r'https://softparade.net/convert'
            data = 'file={}&from=doc&to={}'.format(id, outputType.lower())
            headers = {
                'Host': 'softparade.net',
                'Origin': 'https://word2jpg.com',
                'Referer': 'https://softparade.net/',
                'Content-Type': r'application/x-www-form-urlencoded',
                'User-Agent': r'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.64 Safari/537.36 Edg/101.0.1210.53',
            }
            res = requests.post(url=url, headers=headers, data=data)
            res = res.json()
            param_map = {'fileName': fileName, 'id': res['id'], 'storePath': storePath}
            self.downloadFile(param_map, types='doc2image', debugPrint=debugPrint)


        url = r'https://softparade.net/uploadfile/'
        files = {"file": (fileName, open(filePath, "rb"))}
        debugPrint('>> 开始请求文件转换中...')
        res = requests.post(url=url, files=files).json()
        if res['status'] == 'ok':
            debugPrint('>> 文件转换完成')
            res = download(res['id'])
            return res
        else:
            debugPrint('>> 文件转换出错：' + str(res))
            return None




rootPath = os.path.dirname(__file__)
if __name__ == '__main__':
    # res = Convertor.excel2image('a.xls')
    # print(res)
    res = Convertor().doc2image('./test/a.doc')
    print(res)
