# [ARCHIVED] Silly stories Word add-in sample: load files and use content controls

**Note:** This repo is archived and no longer actively maintained. Security vulnerabilities may exist in the project, or its dependencies. If you plan to reuse or run any code from this repo, be sure to perform appropriate security checks on the code or dependencies first. Do not use this project as the starting point of a production Office Add-in. Always start your production code by using the Office/SharePoint development workload in Visual Studio, or the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office), and follow security best practices as you develop the add-in. 

This Word add-in will show you how to:

1. Load a list of docx files from a service and populate a drop down box control with the file names as options.
2. Load a docx file from the service and insert it into the Word document.
3. Load the content control collection and create input boxes based on the content controls.
4. Update the text value of the content control collection based on the values in the input boxes.
5. Use Office UI Fabric to create a seamless Word user experience.

> Note: A giggle is a side effect of running this sample.

## Prerequisites

To use the Silly stories Word add-in sample, the following are required.

* [node.js](https://nodejs.org) to serve up the docx files.
* [npm](https://www.npmjs.com/) to install the dependencies.
* JQuery, for the Office UI Fabric [dropdown](dev.office.com/fabric/components/dropdown) component.
* Word 2016, or any client that supports the Word Javascript API. This sample does a requirement check to see if it is running in a supported host.

## Start the web application

1. Install project dependencies with Node's package manager (npm) by running ```npm install``` in the project's root directory on the command line.
2. Start the development server by running ```node server.js``` in the project's root directory. The add-in will be running at 127.0.0.1:8080.

### Configure and run on Word for Mac 2016

1. Create a folder called “wef” in Users/Library/Containers/com.microsoft.word/Data/Documents/
2. Put the manifest in the wef folder (Users/Library/Containers/com.microsoft.word/Data/Documents/wef)
3. Open Word 2016 on the Mac and click on the Insert tab > My Add-ins drop down. You should see the add-in listed in the drop down. Select it and it will load the add-in.

### Configure and run on Word for Windows 2016

1. Create a network share, or [share a folder to the network](https://technet.microsoft.com/en-us/library/cc770880.aspx) and place the [word-add-in-sillystories.xml](word-add-in-sillystories.xml) manifest file in it. You've deployed your add-in at this point. Now you need to let Word know where to find the add-in.
2. Launch Word and open a document.
3. Choose the **File** tab, and then choose **Options**.
4. Choose **Trust Center**, and then choose the **Trust Center Settings** button.
5. Choose **Trusted Add-ins Catalogs**.
6. In the **Catalog Url** box, enter the network path to the folder share that contains word-add-in-sillystories.xml and then choose **Add Catalog**.
7. Select the **Show in Menu** check box, and then choose **OK**.
8. A message is displayed to inform you that your settings will be applied the next time you start Office. Close and restart Word. 

Now you are ready to run it in Word. 

1. Open a Word document. 
2. On the **Insert** tab in Word 2016, choose **My Add-ins**. 
3. Select the **Shared folder** tab.
4. Choose **Silly stories add-in**, and then select **Insert**.
5. A new group called **Word Add-in** will appear on the **Home** tab. The group has a button named **Silly Stories**. (These are not seen in the screen shot.) Click this button to open the add-in task pane.
6. Select a story,  to have boilerplate text entered into the Word document.

__Figure 1. The Silly stories add-in loaded in Word__

![Picture of the Word application with the Silly stories add-in loaded](./readme-images/sillystoriesUI.PNG)

## Questions and comments

We'd love to get your feedback about the Silly stories Word add-in sample. You can send your questions and suggestions to us in the [Issues](https://github.com/OfficeDev/Word-Add-in-SIllyStories/issues) section of this repository.

Questions about add-in development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Make sure that your questions or comments are tagged with [office-js], [word], and [API].

## Learn more

Here are more resources to help you create Word Javascript API based add-ins:

* [Sideload an Office Add-in on iPad and Mac](http://dev.office.com/docs/add-ins/testing/sideload-an-office-add-in-on-ipad-and-mac)
* [Office Add-ins platform overview](https://msdn.microsoft.com/EN-US/library/office/jj220082.aspx)
* [Word add-ins](https://github.com/OfficeDev/office-js-docs/blob/master/word/word-add-ins.md)
* [Word add-ins programming overview](https://github.com/OfficeDev/office-js-docs/blob/master/word/word-add-ins-programming-guide.md)
* [Snippet Explorer for Word](http://officesnippetexplorer.azurewebsites.net/#/snippets/word)
* [Word add-ins JavaScript API Reference](https://github.com/OfficeDev/office-js-docs/tree/master/word/word-add-ins-javascript-reference)

## Copyright
Copyright (c) 2015 Microsoft. All rights reserved.


This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
