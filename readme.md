## 功能
可以将Excel中URL的图片，直接转化正常的图片呈现。

## Window平台打包为exe可执行文件
1、导出依赖包：pip freeze > requirements.txt
2、安装打包工具：pip install pyinstaller
3、在本项目目录下，执行打包命令：pyinstaller --onefile  main.py --name Excel中URL转图片
