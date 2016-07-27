# Silly stories Word 外接程序示例：加载文件和使用内容控件

此 Word 外接程序将向你展示如何：

1. 从服务中加载一系列 docx 文件并将使用作为选项的文件名填充下拉列表框控件。
2. 从服务中加载一个 docx 文件并将其插入到 Word 文档中。
3. 加载内容控件集合并基于内容控件创建输入框。
4. 基于输入框中的值更新内容控件集合的文本值。
5. 使用 Office UI 结构创建无缝的 Word 用户体验。

> 注意：运行此示例的一个副作用是“偷笑”。

## 先决条件

若要使用 Silly stories Word 外接程序示例，必须符合以下条件。

* [node.js](https://nodejs.org)，可提供 docx 文件。
* [npm](https://www.npmjs.com/)，可安装依赖项。
* JQuery，用于 Office UI 结构[下拉列表](dev.office.com/fabric/components/dropdown)组件。
* Word 2016 或任何支持 Word Javascript API 的客户端。此示例会执行要求检查以查看是否正在受支持的主机中运行。

## 启动 Web 应用程序

1. 通过在命令行的项目根目录中运行 ```npm install``` 来安装项目依赖项和节点的程序包管理器 (npm)。
2. 通过在项目根目录中运行 ```node server.js``` 启动开发服务器。外接程序将在 127.0.0.1:8080 运行。

### 在 Word for Mac 2016 上配置并运行

1. 在 Users/Library/Containers/com.microsoft.word/Data/Documents/ 中创建一个名称为“wef”的文件夹
2. 将清单放入 wef 文件夹 (Users/Library/Containers/com.microsoft.word/Data/Documents/wef) 中
3. 打开 Mac 上的 Word 2016，并单击“插入”选项卡 >“我外接程序”下拉列表。您应该看到下拉列表中列出了该外接程序。选择该外接程序，然后将加载该外接程序。

### 在 Word for Windows 2016 上配置并运行

1. 创建网络共享，或[将文件夹共享到网络](https://technet.microsoft.com/zh-cn/library/cc770880.aspx)，并将 [word-add-in-sillystories.xml](word-add-in-sillystories.xml) 清单文件放入该文件夹中。此时，你已部署了外接程序。现在，你需要让 Word 知道在哪里可以找到该外接程序。
2. 启动 Word，然后打开一个文档。
3. 选择**文件**选项卡，然后选择**选项**。
4. 选择**信任中心**，然后选择**信任中心设置**按钮。
5. 选择**受信任的外接程序目录**。
6. 在“**目录 URL**”字段中，输入包含 word-add-in-sillystories.xml 的文件夹共享的网络路径，然后选择“**添加目录**”。
7. 选中**显示在菜单中**复选框，然后单击**确定**。
8. 随后会出现一条消息，告知您下次启动 Office 时将应用您的设置。关闭并重新启动 Word。 

现在，你即可在 Word 中运行。 

1. 打开一个 Word 文档。 
2. 在 Word 2016 中的**插入**选项卡上，选择**我的外接程序**。 
3. 选择**共享文件夹**选项卡。
4. 选择“**Silly stories 外接程序**”，然后选择“**插入**”。
5. 外接程序将加载在任务窗格中。参见图 1 查看其在加载时的外观。
6. 选择一个故事以在 Word 文档中输入样本文字。

**图 1. 加载到 Word 中的 Silly stories 外接程序**

![加载了 Silly stories 外接程序的 Word 应用程序的图片](../readme-images/sillystoriesUI.PNG)

## 问题和意见

我们乐意倾听你对 Silly stories Word 外接程序示例的相关反馈。你可以在该存储库中的“[问题](https://github.com/OfficeDev/Word-Add-in-SIllyStories/issues)”部分将问题和建议发送给我们。

与外接程序开发相关的问题一般应发布到 [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API)。确保你的问题或意见使用了 [office-js]、[word] 和 [API] 标记。

## 了解详细信息

以下是更多的资源，可帮助你创建基于 Word Javascript API 的外接程序：

* [将 Office 外接程序旁加载到 iPad 和 Mac 上](dev.office.com/docs/add-ins/testing/sideload-an-office-add-in-on-ipad-and-mac)
* [Office 外接程序平台概述](https://msdn.microsoft.com/zh-cn/library/office/jj220082.aspx)
* [Word 外接程序](https://github.com/OfficeDev/office-js-docs/blob/master/word/word-add-ins.md)
* [Word 外接程序编程概述](https://github.com/OfficeDev/office-js-docs/blob/master/word/word-add-ins-programming-guide.md)
* [适用于 Word 的代码段资源管理器](http://officesnippetexplorer.azurewebsites.net/#/snippets/word)
* [Word 外接程序 JavaScript API 参考](https://github.com/OfficeDev/office-js-docs/tree/master/word/word-add-ins-javascript-reference)

## 版权
版权所有 (c) 2015 Microsoft。保留所有权利。
