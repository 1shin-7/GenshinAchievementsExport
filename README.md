# genshin_achievements

#### 介绍
基于图像处理和字符识别，将原神中的成就导出为excel 并与全成就列表比对，从而知道自己哪些成就未完成。

#### 怎么运行
https://www.bilibili.com/video/BV1v44y1L7Et?share_source=copy_web

对于后续的新版本(我也许会更新吧)，只需要
1、在已知栏目.txt 里添加成就栏目名字
2、获取最新的成就列表。有专门的网站https://docs.qq.com/sheet/DS01hbnZwZm5KVnBB?tab=d7oz1q
3、打开原神（1920x1080独占全屏，电脑分辨率同），进入成就并进入具体栏目，管理员身份运行main.py（稍等一会儿，按'r'键启动），导出自己的成就列表，并与最新的成就列表进行比对，生成compare_ans.xlsx。(目前运行时间较长，约10min,把语音播报禁用会快几分钟。）

就可以直接查看自己哪些成就未做。


#### 后续改进点
1、翻页和ocr是最费时间的过程，暂未找到更快的翻页方法。
2、多线程或许可以节省时间。一个线程获取图片，另一个线程用来识别。
3、让参数可以设置

#### 目前问题
想打包成exe，但是paddleocr用pyinstaller打包一直出问题
