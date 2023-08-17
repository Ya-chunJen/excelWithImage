## 功能
在问卷星，如果问卷中有文件上传题，下载到的Excel文件中，只有图片的URL，并不能直接显示图片。使用本工具，可以将「显示图片URL」转为「直接显示图片」。

## 使用方法
1、运行「Excel中URL转图片.exe」（可直接在Releases这个exe可执行文件。，也可以使用后面的方法，自己打包成一个exe可执行文件。）

2、将要处理的Excel拖到打开的命令窗口。（目前只能处理xls文件。）

![图片](https://helpimage.paperol.cn/20230817091553.png)

3、输入要转化列的序号，支持多个用空格分开即可。（可以提前查询要转化的列序号）
![图片](https://helpimage.paperol.cn/20230817091632.png)

4、导出的新文件，存放路径和原文件一致，文件名会以 out+原文件名的命名方式。

## 效果
![图片](https://helpimage.paperol.cn/20230817091947.png)

## Window平台打包为exe可执行文件
1、导出依赖包：pip freeze > requirements.txt

2、安装打包工具：pip install pyinstaller

3、在本项目目录下，执行打包命令：pyinstaller --onefile  main.py --name Excel中URL转图片

4、打包完成后，在本项目目录下的dist文件夹，寻找exe可执行文件
