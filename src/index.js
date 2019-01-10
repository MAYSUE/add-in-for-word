/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
(function () {
    "use strict";

    // The initialize function is run each time the page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // Use this to check whether the new API is supported in the Word client.
            if (Office.context.requirements.isSetSupported("WordApi", "1.1")) {

                console.log('This code is using Word 2016 or greater.');
                $('#test').click(gettingimagedata);
                $('#create').click(show_signature_pad);
                $('#save').click(insert);
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
    /*Insert  Image function into document*/

    function show_signature_pad() {
        document.getElementById('signature-pad').removeAttribute("style");
    }

    function gettingimagedata() {
    Word.run(function (context) {

        // Create a proxy object for the paragraphs collection.
        var paragraphs = context.document.body.paragraphs;

        // Queue a commmand to load the text property for all of the paragraphs.
        context.load(paragraphs, 'text');

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {

            // Queue a a set of commands to get the HTML of the first paragraph.
            var html = paragraphs.items[0].getHtml();

            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Paragraph HTML: ' + html.value);
                var paragraphs = context.document.body;
                paragraphs.insertParagraph("" + html.value, 'End');
            });
        });
    })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });
}
 function insert() {
        var dataURL = signaturePad.toDataURL("image/png");
        dataURL = dataURL.replace('data:image/png;base64,', '');
        Office.context.document.setSelectedDataAsync(dataURL, {
            coercionType: Office.CoercionType.Image,
            imageLeft: 50,
            imageTop: 50,
            imageWidth: 100,
            imageHeight: 80
        },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.log("Action failed with error: " + asyncResult.error.message);
                }
                else {
                    hide_signature_pad();
                }
            });
    }
  function hide_signature_pad() {
        var canvas = document.getElementById('canvas');
        var ctx = canvas.getContext('2d');
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        document.getElementById('signature-pad').setAttribute("style", "visibility:hidden;");
    }

  })();