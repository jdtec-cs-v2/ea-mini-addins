# Enterprise Architect  Mini Add-Ins


这个项目基于Enterprise Architect Add-Ins Model与Automation Interface开发出来的小工具，包含：批量、配置化图片导出；模型内容统计.


## Overview
由于EA仅支持开发语言VB、C#、Delphi，从而本项目采用支持面向对象开发的C#语言。其中：
* 基于Windows Form的MiniAddinsFacade工程实现EA对接门面.
* 基于WPF技术的MiniAddins工程实现EA类似画面的定制、MVVM、实现功能.

![comp](/uploads/9b9494798d7f8c148ce157b3e3936b67/comp.png)

## Develope Language and Tool
* C#
* .NET Framework 4，Visual Studio 2017

## Function Description 
### 用例图

![ea_addins](/uploads/b2d0abfd9e40e27cfcc0cc3f740c0e34/ea_addins.png)

### 概要说明
* 插件安装配置
   - 详细安装方法

* 建模工程师视角
  1. 模型Diagram图片导出(export image、save screen display value)
        ![ucase1](/uploads/150f21345435eddf92c85d20c26e1aa3/ucase1.png)

  2. 模型制作内容统计(statistics workload)
        ![ucase2](/uploads/774dc4b8509b57f59102e0371b92ff38/ucase2.png)

  3. 实时变更监控(export image with modeless)
        ![ucase3](/uploads/a8f16bdcfb044e1d7a7dd84ea46ddfe5/ucase3.png)
        
        - 实现机制是订阅EA Model Object的Watcher对象；
        - 在EA中打开另外模型文件时，Add-Ins上的表示内容也将同步刷新；
        - 变更日志反馈到画面底部的状态栏中，可通过点击【Copy】、【Copy All】拷出当次变更点以及全所变更点;
        - **注意：** EA对Diagram模型中的联接线、位置的变更不会产生变更事件，故在此画面中监控不到此类变更

  4. 保存画面上设定值(save screen display value) <br/>
     - 点击【Confirm】按钮，保存模型的图片导出状况、导出文件名、路径。再次打开模型时，如已导出的Diagram有改变，则显示“changed”，并且呈现变更背景色。

* 其它
    - 画面支持中英双语切换;
    - 支持可通过Diagram名或Package名检索模型图，并且可直接按回车执行检索;


## Supported
* Enterprise Architect 15.2.1554
* Windows 10 企业版

## Install
* 安装方法
    1. 进入Add-Ins-Install目录，双击“RegistryEntry_MiniAddins.reg”文件（注册EA插件）
    
    2. 把dll子目录中的MiniAddinsFacade等DLL文件拷贝到Enterprise Architect的安装目录，例如：C:\Program Files (x86)\Sparx Systems\EA Trial
    
    3. 打开Enterprise Architect，检查Add-Ins是否安装成功

        - Add-Ins是否Loaded状态
        ![ea-addins-install1](/uploads/e9d545ef2fce7fd46c49134c803e6b36/ea-addins-install1.png)
        
        - 菜单是否可用
        ![ea-addins-install2](/uploads/dc736ade2a0fa9dd8a5f4a83d9b99015/ea-addins-install2.png)
