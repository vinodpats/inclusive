/*
 * Copyright (c) Accenture Technology. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 * Developer: Vinod Patil (Date: 6 June 2022)
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
  console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
}

// Assign event handlers and other initialization logic.
document.getElementById("insert-paragraph").onclick = insertParagraph;
  }
});


async function insertParagraph() {
  await Word.run(async (context) => {


    const nonInclusiveList = ['master', 'Slave', 'Blacklist', 'Whitelist', 'Native', 'Non-native', 'Non native','man hours', 'man-hours', 'man days', 'man-days','sanity check', 'insanity check', 'dummy variable', 'stonith', 'kill', 'one throat to check'];
    const inclusiveList= [' primary', ' secondary', ' deny list', ' allow list', ' original', ' non-original', ' non original', ' work hours',' work-hours', ' work days', ' work-days', ' confidence check', ' confidence check', ' indicator variable', ' hardware redundancy', ' discontinue', ' single point of contact'];

    for (let i = 0; i < nonInclusiveList.length; i++) {

    let results = context.document.body.search(nonInclusiveList[i]);
    results.load("length");
    
    context.load(results, 'font');
    await context.sync();

    // Let's traverse the search results... and highlight...
    for (let j = 0; j < results.items.length; j++) {
      
      if(results.items[j].font.strikeThrough != true){
          results.items[j].font.highlightColor = "yellow";
          results.items[j].font.strikeThrough = true;
          
          results.items[j].insertText(inclusiveList[i], Word.InsertLocation.after);
      } //end if
     
    } // for loop of results

  } //main for loop
  
    

      await context.sync();
  })
  .catch(function (error) {
    
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}