# Office2PDF

![office2pdf_v2.0](https://evgo-public.oss-cn-shanghai.aliyuncs.com/repo/office2pdf/office2pdf_v2.png)

## 一、下载使用

### 1. Office2PDF.exe

下载地址：[蓝奏云](https://evgo2017.lanzous.com/s/office2pdf) 密码: evgo2017

下载地址：[Github Release](https://github.com/evgo2017/Office2PDF/releases)

> 1）兼容性未过多测试，若使用有问题可联系我或使用 py 文件。
>
> 2）v2 的 exe 不是单文件，需要解压压缩包（13.5MB），在里面打开 Office2PDF.exe 使用。原因是采用 pyinstaller 进行打包，设置了 -w（Windows 下去掉命令框）和 -F（打包为单文件），就会有 Windows Defender 报错和打开很慢的问题。

### 2. Office2PDF.py

- [x] 需 `Python3` 环境
- [x] 已安装引入的包

## 二、详细说明

### 1. 基本功能


- [x] 将 Office（Word、Excel、PPT ）批量转为 PDF 文件
  - [x] 支持选择**类型**，默认全部类型
- [x] 支持 GUI 界面
- [x] 支持选择**文件夹**，无需复制粘贴
- [x] 支持选择转换**子文件夹**：默认转化
  - [x] 支持结构或平铺：目标文件夹内文件可依照来源文件夹结构放置，也可全平铺无结构
- [x] 内存管理
  - [x] 每种格式转换只打开**一个**进程
  - [x] 转换完成后关闭进程，内存回收 **gc** 收集

### 2. 转换细节

- [x] Word
  - [x] 有内容
  - [x] 无内容
- [x] Excle
  - [x] 有内容：生成该内部工作表个数的 PDF 文件
  - [x] 无内容：生成的 PDF 不可正确打开
- [x] PPT
  - [x] 有内容：多页
  - [x] 无内容：提示错误跳过转换

### 3. 运行要求

- [x] 已安装 Office
- [x] [Microsoft Save as PDF](<https://www.microsoft.com/zh-cn/download/details.aspx?id=7> ) 加载项

> 建议 Office 2007 及以上，自带 Microsoft Save as PDF 加载项，基本上安装了 Office 就可以直接使用了

## 三、最后

若有更新信息，会在此文档进行说明。

如果有问题可以在 [Issues](https://github.com/evgo2017/Office2PDF/issues) 留言，或者[联系我](https://evgo2017.com/link)。

写出第一版程序的同时写了一篇文章：[office 转 pdf 技巧及软件](<https://mp.weixin.qq.com/s?__biz=MzIwMjk2MTQ1MQ==&mid=2247484268&idx=1&sn=80bf791cae04e836b25525e3039fa3ff&chksm=96d7e428a1a06d3eb0ba59c98b5f772ca621792cda53abef70218d94ac1239d2c2fb71a8b539#rd> )，有兴趣可以读读。

## 四、更新记录

| 时间       | 内容                                                         | 备注 |
| ---------- | ------------------------------------------------------------ | ---- |
| 2020.08.26 | v2，加入 GUI，支持选择类型、子文件夹等功能                   |      |
| 2019.05.13 | 将此项目从自己的 `SomeTools` 项目独立出来，通过 `release` 发布 `exe` |      |
| 2018.11.02 | v1，功能基本实现                                             |      |
