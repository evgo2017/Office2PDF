# Office2PDF

![office2pdf_v2.0](https://github.com/evgo2017/Office2PDF/raw/master/assets/office2pdf_v2.png)

## 一、下载使用

下载地址 1：[网盘](http://evgo2017.ysepan.com/)（项目密码: evgo2017）（两小时内有下载量的限制，故设置密码）

下载地址 2：[Github Release](https://github.com/evgo2017/Office2PDF/releases)

## 二、详细说明

### 1. 基本功能

- [x] GUI 界面
- [x] 转换类型（Word、Excel、PPT）
  - [x] 转换特定类型
  - [x] 转换全部类型
- [x] 子文件夹
  - [x] 是否转换
  - [x] 转换后的目标文件夹结构：既可保持来源文件夹**结构**，也可全**平铺**
- [x] 内存回收

### 2. 转换细节

|            | Word | Excel                                | PPT               |
| ---------- | ---- | ------------------------------------ | ----------------- |
| 文档有内容 | ✅    | ✅（若有多个 Sheet，则生成多个文件）  | ✅（多页）         |
| 文档无内容 | ✅    | ❌（会跳过，不会产生对应的 PDF 文件） | ❌（提示错误跳过转换） |

### 3. 运行要求

#### Office 2007 及以后版本

- [x] 已安装 Office

#### Office 2007 以前版本

- [x] 已安装 Office
- [x] Microsoft Save as PDF  加载项

> 建议 Office 2007 及以上，自带 Microsoft Save as PDF  加载项

## 三、最后

若更新信息，会在此文档进行最新说明。

若有问题请在 [Issues](https://github.com/evgo2017/Office2PDF/issues) 留言，或者[联系我](https://evgo2017.com/about)。

> 含界面的 exe 需要解压压缩包（13.5MB左右）后，在里面打开 Office2PDF.exe 使用。
>
> 原因是采用 pyinstaller 进行打包时，若设置了 -w（Windows 下去掉命令框）和 -F（打包为单文件），就会有**打开很慢**和 Windows Defender 报错的问题。

## 四、版本更新记录

| 时间       | 内容                                                         | 相关文章 |
| ---------- | ------------------------------------------------------------ | ---- |
| 2020.08.26 | v2，加入 GUI，支持选择类型、子文件夹等功能                   | [Office2PDF 批量转 PDF（第二版）](https://mp.weixin.qq.com/s/VxHxvUUqK2tn0PKNQkXTsQ)     |
| 2019.05.13 | 将此项目从自己的 `SomeTools` 项目独立出来，通过 `release` 发布 `exe` |      |
| 2018.11.02 | v1，功能基本实现                                             | [office 转 pdf 技巧及软件](https://mp.weixin.qq.com/s/jZvVXgqcMOIxkKVzJXYEZA)      |
