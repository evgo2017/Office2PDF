# Office2PDF

作者：evgo（evgo2017.com）

[exe 下载地址](<https://github.com/evgo2017/Office2PDF/releases> )

## 待增加功能

> 近期在完善，暂未更新 exe 文件

- [x] 增加界面
- [x] 手动选择文件夹，无需复制粘贴
- [x] 选择 word, excel, ppt 类型，默认全部
- [x] 支持子文件夹，默认转化
    - [X] 支持原文件夹结构或平铺

![office2pdf_v2.0运行示例](https://evgo-public.oss-cn-shanghai.aliyuncs.com/repo/office2pdf/office2pdf_v2.png)

## 一、程序功能

`Office` 文件（word、excel、ppt）批量转为 `PDF` 文件。功能较完善，已用半年多，很满意。

提供 `py` 源码与生成的 `exe` 。

输出细节：

- [x] Word 有内容
- [x] Word 无内容
- [x] Excle 有内容 -（多工作表）
- [x] PPT 有内容 - （多页）

- [x] Excle 无内容 - 报错跳过
- [x] PPT 无内容 - 报错跳过
- [x] 每种格式转换只打开**一个**进程
- [x] 转换完成，关闭进程， **gc** 收集

>  Excel 文件会根据内部的工作表数量生成对应数量的 `PDF` 文件。

## 二、运行示例

需要转换的 `Office` 文件，若与程序在**同一级**目录下，**直接回车**即可转换；否则输入 `Office` 文件所在文件夹的**绝对路径**。

> 仅当级目录，不包含子目录。

![py运行示例](assets/example.png)

（程序运行示例图）

## 三、运行要求

- [x] 已安装 Office（推荐 2007 版本以上）

> 主要是利用 [Microsoft Save as PDF](<https://www.microsoft.com/zh-cn/download/details.aspx?id=7> ) 插件，较新的版本都自带了。

## 四、多种格式

### Office2PDF.py

- [x] 需 `Python` 环境
- [x] 已安装引入的包

### Office2PDF.exe

下载地址：[Office2PDF.exe](<https://github.com/evgo2017/Office2PDF/releases> )

兼容性暂时无法测试，若出问题，可选择 `py` 文件或联系我。

### Office2PDF.java

用 `Java` 语言实现功能的源码。不推荐。

只是实现了基础功能，不够完善。有一定的对比学习意义。因为 `Java` 安装运行较为麻烦，分享不够便利，于是换成 `Python` 语言实现。

## 五、Test

在 `test ` 文件夹内，是用于测试的各格式文档。

## 六、最后

若有更新信息，会在此文档进行说明。



原先在自己 `SomeTools` 项目内，了解到 `release` 后就想着独立出来去发布 `exe` ，之前是与源码在一起的。

写出程序的同时写了一篇文章：[office 转 pdf 技巧及软件](<https://mp.weixin.qq.com/s?__biz=MzIwMjk2MTQ1MQ==&mid=2247484268&idx=1&sn=80bf791cae04e836b25525e3039fa3ff&chksm=96d7e428a1a06d3eb0ba59c98b5f772ca621792cda53abef70218d94ac1239d2c2fb71a8b539#rd> )，有兴趣可以读读。

## 七、更新记录

| 时间       | 内容                                                 | 备注 |
| ---------- | ---------------------------------------------------- | ---- |
| 2020.04.30 | 看尝试 c# 重新实现，并且写出界面。加入遍历子文件夹。 |      |
|            |                                                      |      |
|            |                                                      |      |

