/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  return Word.run(async function (context) {
      // Get the current selection
      var range = context.document.getSelection();

      // Load the text property
      context.load(range, 'text');

      // Synchronize the document state
      await context.sync();

      // Get the text from the range
      var selectedText = range.text;

      // Prepare the request body with selectedText
      const data = selectedText
      const requestBody = JSON.stringify(data);
      console.log(requestBody);
      const requestHeaders = new Headers({
        "Content-Type": "application/json",
      });
      const url = 'http://localhost:7071/api/httpWordAddIn'

      try {
          const response = await fetch(url, {
              method: "POST",
              body: requestBody,
              headers: requestHeaders
          });
          console.log(response);

          if (!response.ok) {
              console.error('Error: ' + response.text());
              throw new Error("Request failed with status code " + response.status);
          }

          const textResponse = await response.text();
          const textToInsert = "\n##############################################\n" + textResponse + "\n##############################################\n";
          console.log(textToInsert);
          // Assuming the response contains the text to be inserted
          //const resultText = jsonResponse.resultText; // Adjust according to actual response structure

          // Write the result text to the document
          range.insertText(textToInsert, Word.InsertLocation.end);
          await context.sync();

      } catch (error) {
          console.error('Error: ' + JSON.stringify(error));
      }
  });
}