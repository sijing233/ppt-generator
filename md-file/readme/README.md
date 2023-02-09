
---
Title: 自动生成的PPT
Author: 司镜233
Date: 2023.02.09
---

# 效果展示

## 效果展示

> 自动化生成的PPT在/out-print目录中

- 自动化生成的PPT.pptx 是本篇内容自动化生成的PPT
- 依照/ppt-model目录中的default模板
- 关注微信公众号"司镜233"，可以免费获取3个模板

## 介绍视频

> 介绍和学习使用的视频，已经放在了b站

- 链接:
- 记得一键三连哦~~~

# 现有功能

## 现有功能
> 你是否还在因为制作PPT而耗费时间？代码自动生成帮助您一站式解决

- 根据母版，自动化生成首页
- 自动化生成尾页
- 自动化生成目录
- 根据文字内容，生成具体的PPT页面
- 可以插入图片，一张PPT内不要超过两张
- 生成后，根据自己的需要可以微调
- 暂不支持动画

# 环境准备

## 环境要求

> python、及一些安装包

![img_7.png](E:\project\ppt-generator\md-file\readme\assets\img_7.png)

- 安装python3+
- 可以使用pycharm
- 安装python-pptx包

## 1、安装python

> 如果已安装可以忽略

![img_8.png](E:\project\ppt-generator\md-file\readme\assets\img_8.png)

- 打开官网python.org
- 点击灰色的按钮下载
- 下载后，按正常安装软件的程序安装
- 记得勾选Path，不然还要自己配置环境变量


## 2、用pip安装需要的库

> 安装需要的库

- pip install python-pptx

# 基本的格式

## 1、在指定的文件夹内，书写文档内容

> 可以在任何地方，书写md文件，但要记得路径

![img_1.png](E:\project\ppt-generator\md-file\readme\assets\img_1.png)

- 起任何名字，与生成的PPT标题无关
- 最好起英文的，因为你的环境有可能不支持中文
- 生成的ppt文件，在out-print中
- 可以建一个指定的文件夹，专门用于写文档

## 2、写PPT的题目

> 用两行---夹住的中间部分

![img.png](E:\project\ppt-generator\md-file\readme\assets\img_3.png)

- Title: 后面写PPT的首页标题
- Author: 后面写你的名字，或是其他副标题
- Data: 后面可以写日期，作为尾页的日期


## 3、写PPT的目录

> 用#，加上一个空格，在后面书写的是一级标题，会自动生成目录

![img_2.png](E:\project\ppt-generator\md-file\readme\assets\img_2.png)

- 标题和标题，不要重复，否则会覆盖
- 一级标题，会自动生成PPT的目录

## 4、写每一页的标题

> 用##，加上一个空格，在后面书写的是二级标题，会自动生成为每一页的标题

![img.png](E:\project\ppt-generator\md-file\readme\assets\img.png)

- 在每个一级标题下，一定要有二级标题
- 有多少个二级标题，就会生成多少页的PPT
- 二级标题之间，不要重复，否则会覆盖

## 5、写每一页的内容

> 每页内容，可以有几种组成的部分

![img_4.png](E:\project\ppt-generator\md-file\readme\assets\img_4.png)

- ‘>’+空格：一行文字介绍，呈现在页标题的下方
- 图片![图片的名字](图片的目录地址)
- 图片地址要写绝对路径
- 一张图片会居中，两张图片会平分
- 列表用'-’+空格，后面书写该陈列点的内容
- 可以一张图片+一部分陈列点的方式


# 使用指南

## 1、选择合适的模板

> 在ppt-model中，选择自己喜欢的模板

- 有一个默认的模板，在ppt-model/default-model.pptx
- 其他更多的模板，可以关注微信公众号"司镜233"，付费获得

## 2、修改自己的文件路径

> 在代码中，修改自己文件的路径

![img_5.png](E:\project\ppt-generator\md-file\readme\assets\img_5.png)

- 在run-file/default-ppt-generate.py中修改
- 如果更换模板，修改模板路径file_path
- 修改输出的文件名 out_file_path
- 修改你写的文件名 md_file_path


## 3、运行生成PPT

> 运行代码，生成PPT

![img_6.png](E:\project\ppt-generator\md-file\readme\assets\img_6.png)

- 运行模板相应的代码wutong-pptx-generate.py
- 重新生成，需要关掉已经生成的ppt文件


# 付费定制化

## 获取更多的模板

> 在微信公众号内回复"PPT模板"，即可获得更多的模板

- 现在共有10个模板
- 定价￥50
- 关注微信公众号"司镜233", 回复”PPT模板“
- 付费后，会给您发送特定的PPT模板和模板代码

## 模板定制化

> 可以根据你自己的喜好和要求，定制化一套模板和相应的代码

- 定制带有特定个的logo、风格的模板，作为长期使用
- 付费标价￥500/次，如果比较复杂看情况而定
- 关注微信公众号"司镜233", 留下你的要求

