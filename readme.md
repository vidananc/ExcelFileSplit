# 拆分给定Excel表格并控制打印系统
## 已经实现的功能
1. 前端页面选择`excel`文件，后端按学校（考点）进行拆分，压缩为zip文件，返回
2. 可以通过`python`脚本调用打印软件`bartender`，通过命令行执行，可以将指定的excel文件按照模板文件`btw`进行打印
3. 前端页面可以选择excel文件，指定对应的工作表，设置打印间隔、是否需要语音播报等参数。
## 目前的问题
1. 通过`python`脚本直接调用`bartender`打印软件打印excel软件时，必须手动打开一次excel文件，不然打印软件无法识别格式，这个问题还未解决，如果解决了，基本就完成了所有需求
## 环境配置
1. python 环境
### 执行命令
```
pip install pandas
pip install xlwt
pip install flask flask-cors
```
## 启动服务
到项目的根目录，确保已经配置`python`环境变量，执行命令`python server.py`
## 访问服务
打开浏览器输入：`http://localhost:8080`即可。
