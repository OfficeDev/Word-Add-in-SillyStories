/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
(function() {
    "use strict";

    // The initialize function is run each time the page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
              
            // Use this to check whether the new API is supported in the Word client.
            if (Office.context.requirements.isSetSupported("WordApi", "1.1")) {

                console.log('This code is using Word 2016 or greater.');

                // Load the story names from the service into the UI. 
                getStoryNames();
                
                // Setup the event handlers for UI.
                $('#selectStory').change(selectStoryHandler);
                $('#makeSillyStory').click(makeSillyHandler);
            }
            else {
                // Just letting you know that this code will not work with your version of Word.
                console.log('This add-in requires Word 2016 or greater.');
            }    
        });
    };

    /*********************/
    /* Word JS functions */
    /*********************/
    
    // Using the Word JS API. Clears the current document and adds a base64 docx file to the document.
    function displayContents(myBase64) {
        Word.run(function (context) {

            // Create a proxy object for the document.
            var thisDocument = context.document;
            
            // Queue a command to clear the body contents. 
            thisDocument.body.clear();
            
            // Create a proxy object for the default selection. 
            var mySelection = thisDocument.getSelection();
            
            // Queue a command to insert the file into the current document.
            mySelection.insertFileFromBase64(myBase64, "replace");

            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    // Now we want to get all of the content controls.
                    getAllContentControls();
                });
        })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });            
    }

    // Using the Word JS API. Gets all of the content controls that are in the loaded document. 
    function getAllContentControls() {
        Word.run(function (context) {

            // Create a proxy object for the document.
            var thisDocument = context.document;

            // Create a proxy object for the content control collection in the current document.
            var contentControls = thisDocument.contentControls;

            // Queue a command to load the tag and title properties of the content controls.
            contentControls.load('tag, title');

            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync(contentControls).then(function () {

                // Remove duplicate content controls and get back an array of objects
                // that represent unique fields in the add-in. For example, there may be 
                // many content controls titled "name" in the document, but we want all
                // content controls with the same name represented
                // by a single field in the add-in.
                var uniqueFields = removeDuplicateContentControls(contentControls);
                
                // Create HTML input fields based on the content controls.
                createFields(uniqueFields);
            })
        })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });  
    }

    // Using the Word JS API. Set the values from the INPUT elements into the associated
    // content controls to make the story silly. 
    function makeSillyHandler() {

        // We will only act on the INPUT elements. We assume that all INPUT
        // elements are mapped to a content control.
        var entryFields = document.getElementById("entryFields").querySelectorAll("input");

        // I don't like this. I'm loading the content controls again.
        Word.run(function (context) {

            // Create a proxy object for the document.
            var thisDocument = context.document;

            // Create a proxy object for the content control collection in the document.
            var contentControls = thisDocument.contentControls;

            // Queue a command to load the content controls collection with the tag and title properties.
            contentControls.load('tag, title');

            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {

                    var i, j;
                
                    // For each input element, we want to map it to the associated
                    // content control.
                    for (i = 0; i < entryFields.length; i++) {
                        for (j = 0; j < contentControls.items.length; j++) {

                            // Matching content control tag with the tag set as the id on each input element.
                            // Set the content text to the text value of the INPUT element.
                            if (contentControls.items[j].tag === entryFields[i].id) {
                                contentControls.items[j].insertText(entryFields[i].value.trim(), Word.InsertLocation.replace)
                            }
                        }
                    }
                })
                // Synchronize the document state by executing the queued commands, 
                // and return a promise to indicate task completion.
                .then(context.sync);
        })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });    
    }
    
    // Handles the story selection in the add-in UI. Results in a call to the service to get
    // a docx file that contains a silly story. The silly story gets displayed in the Word UI.
    function selectStoryHandler() {
        
        // Form the URL to get the docx file. Need to remove the host information by slicing
        // off the host information beginning at ?_host_Info.
        var fileName = this.value;
        var currentUrl = location.href.slice(0, location.href.indexOf('?'));
        var getFileUrl = currentUrl + 'getfile?filename=' + fileName;

        // Call the helper to get the selected file then insert the file into Word.
        httpGetAsync(getFileUrl, function(response) {
            displayContents(response);
        });
    }

    /********************/
    /* UI functions     */
    /********************/
    
    // Gets a list of story file names from the service and then create entries in a drop down list box.
    // A user can choose a story once the drop down list box is populated.
    function getStoryNames() {

        // Form the URL to get the docx file list. Need to remove the host information by slicing
        // off the host information beginning at ?_host_Info.
        var currentUrl = location.href.slice(0, location.href.indexOf('?'));
        var getFileNamesUrl = currentUrl + 'getfilenames';
    
        // Call the helper to get the file list and then create the drop down 
        // list box from the results.
        httpGetAsync(getFileNamesUrl, function(rawResponse) {

            // Helper that processes the response so that we have an array of file names.
            var response = processResponse(rawResponse);
   
            // Get a handle on the empty drop down list box.
            var select = document.getElementById("selectStory");

            // Populate the drop down list box.       
            for(var i = 0; i < response.length; i++) {
                var opt = response[i];
                var el = document.createElement("option");
                el.textContent = opt;
                el.value = opt;
                select.appendChild(el);
            };
            
            // Initialize stylized fabric UI for dropdown. This happens after we
            // populate the dropdown since the drop down values are copied by fabric.
            $(".ms-Dropdown").Dropdown();
        });
    }    
    
    // Creates HTML input fields based on the content control title and tag properties.
    // This method assumes that all content controls should be turned into HTML input fields
    // and that the duplicate have already been removed.
    function createFields(uniqueFields) {
        
        // Get the DIV where we will add out INPUT controls.
        var entryFields = document.getElementById("entryFields");

        // Clear the contents in case it has already been populated with INPUT controls.
        while (entryFields.hasChildNodes()) {
            entryFields.removeChild(entryFields.lastChild);
        }

        // Create a unique INPUT element for each unique content control tag.
        for (var i = 0; i < uniqueFields.length; i++) {
            entryFields.appendChild(document.createTextNode(uniqueFields[i].title + ': '));
            var input = document.createElement("input");
            input.type = "text";
            input.id = uniqueFields[i].tag;
            entryFields.appendChild(input);
            entryFields.appendChild(document.createElement("br"));
        }
    }
    
    /********************/
    /* Helper functions */
    /********************/
    
    // Helper that deduplicates the set of content controls. Content controls are
    // considered duplicates if they share the same tag.
    function removeDuplicateContentControls(contentControls) {
    
        var i,
            len = contentControls.items.length,
            uniqueFields= [],
            currentContentControl = {};
        
        for (i = 0; i < len; i++) {
            currentContentControl[contentControls.items[i].tag] = contentControls.items[i].title;
        }
        
        for (i in currentContentControl) {
            
            var obj = {
                tag: i,
                title: currentContentControl[i]
            };
            
            uniqueFields.push(obj);
        }
        
        return uniqueFields;
    }
    
    // Helper for calls to the service. 
    function httpGetAsync(theUrl, callback)
    {
        var request = new XMLHttpRequest();
        request.open("GET", theUrl, true);
        request.onreadystatechange = function() { 
            if (request.readyState === 4 && request.status === 200)
                callback(request.responseText);
        }
        request.send(null);
    }
    
    // Helper that processes file names into an array. This is because the service returns
    // the file names as ["filename1.docx","filename2.docx","filename3.docx"].
    function processResponse(rawResponse) {
        
        // Remove quotes.
        rawResponse = rawResponse.replace(/"/g, "");
        
        // Remove opening brackets.
        rawResponse = rawResponse.replace("[", "");
        
        // Remove closing brackets.
        rawResponse = rawResponse.replace("]", "");
        
        // Return an array of file names.
        return rawResponse.split(',');
    }
})();