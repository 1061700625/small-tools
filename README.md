# small-tools
使用pyqt5和Python制作的小工具集，包括压缩、解压、excel转图片等  

- 压缩与解压部分的博客:https://blog.csdn.net/sxf1061700625/article/details/123632095    
- excel转图片部分的博客：https://xfxuezhang.blog.csdn.net/article/details/124898696   

## 文件简介
- `GUI.py`和`GUI.ui`：绘制界面相关
- `convert.bat`：使用pyuic5将ui转为py
- `打包.bat`和`main.spec`：使用pyinstaller进行打包为exe
- `UnRAR64.dll`和`compress.py`：压缩解压相关
- `convertor.py`：excel/word转图片相关
- `main.py`：主入口

## 功能支持
- 文件夹压缩
- zip/rar文件解压
- excel转图片
- doc/docx转图片


## 效果图
![1](https://user-images.githubusercontent.com/31002981/169663001-60e740b0-c297-4c9c-9afe-d311305b8537.png)
![2](https://user-images.githubusercontent.com/31002981/169663019-c6406963-9b1c-4d6d-aa9d-ac92459185a1.png)
![3](https://user-images.githubusercontent.com/31002981/169663026-df8fa1ad-c1b7-47f1-9106-b7a512fa6914.png)

## TODO
- 使用QThread执行耗时任务，如Excel转图片时的post请求
- 使用progressBar
- debug窗口实时滚动
- 添加更多其他小工具

