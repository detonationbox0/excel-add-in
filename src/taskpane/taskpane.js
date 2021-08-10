/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {

    // Attach Events
    // If the table does not yet exist, add event to init click
    Excel.run(function (context) {

      var sheet = context.workbook.worksheets.getActiveWorksheet();
      var mergeTable = sheet.tables.getItem("DatamergeTable");

      mergeTable.load("name");

      return context.sync().then(function() {
        // Table exists.... Disable the initialize button
        $("#init").addClass("disabled-button").find("span").text("Table Initialized...");
        // Show the split options
        $("#split-options").slideDown();
        
      });

    }).catch(function (err) {
        // Table doesn't exist, add event to Initialize button
        console.log(err)
        document.getElementById("init").onclick = init;
    });
    
    document.getElementById("count").onclick = count;

  }
  console.log("READY!");
});


/**
 * -------------------------------------------------------------------------------------------------- |
 * Convert the list to a Table
 * -------------------------------------------------------------------------------------------------- |
 */

export async function init() {
  //#region
  console.log("Converting the Used Range of Active Sheet to a Table");
  Excel.run(function (context) {
    /**
     * Insert your Excel code here
     */
    const range = context.workbook.getSelectedRange();
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedData = sheet.getUsedRange();

    // Send to Sink
    usedData.load("values");
    usedData.load("address")
   

    return context.sync()
      .then(function () {
          var dataTable = sheet.tables.add(usedData.address, true /*hasHeaders*/);
          dataTable.name = "DatamergeTable";

          // Update the initialize button
          $("#init").addClass("disabled-button").find("span").text("Table Initialized...");
          // Remove event
          document.getElementById("init").onclick = function() {
            return false;
          }
          // Show the split options
          $("#split-options").slideDown();

      });
    // console.log(`The range address was ${range.address}.`);

  }).catch(function(err){
    console.log(err);
    showError("The table has already been initialized.");

  });



  //#endregion
}

/**
 * -------------------------------------------------------------------------------------------------- |
 * Add Count Column
 * -------------------------------------------------------------------------------------------------- |
 */
function addCountColumn () {

  Excel.run(function (context) {

    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var mergeTable = sheet.tables.getItem("DatamergeTable");
    mergeTable.columns.add(0 /*index*/, null /*values*/, "count" /*name*/)
    // return context.sync().then(function() {
    //   console.log(mergeTable.values);
    // });
    return context.sync();

  }).catch(function(err){
    console.log(err)
    showError(err);
  });

}


export async function count() {
  //#region
  console.log("Add Count column to table");
  addCountColumn();
  
  //#endregion
}


/**
 * -------------------------------------------------------------------------------------------------- |
 * Update / Show the Message Log
 * -------------------------------------------------------------------------------------------------- |
 */
function showError(theError) {
  //#region
  $("#log-slide").fadeIn(); // Show the error log area
  $("#log-message").text(theError);
  //#endregion
}

/**
 * -------------------------------------------------------------------------------------------------- |
 * Hide the Message Log
 * -------------------------------------------------------------------------------------------------- |
 */

$("#log-dismiss").on("click", function() {
  //#region
  $("#log-slide").fadeOut(function() {
    // Done fading
    $("#log-message").empty();
  });
  //#endregion
})

/**
 * -------------------------------------------------------------------------------------------------- |
 * Split Button
 * -------------------------------------------------------------------------------------------------- |
 */

$("#split").on("click", function() {
  var alt = $("#split-alt").is(":checked"); // Will we alternate?
  var up = $("#split-up").val();

  /**
   * Perform the following actions, in order:
   * Add count column, add alternate column if alt is checked
   * Split the document based on up and row count
   */
  sort(alt);
  // split();

})

/**
 * -------------------------------------------------------------------------------------------------- |
 * The Fancy Sort Function
 * Adds the count column
 * Optionally - Creates an alternate column with modulo function and sorts by that.
 * -------------------------------------------------------------------------------------------------- |
 */

function sort(alt) {

  Excel.run(function (context) {

    // Grab the table
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var mergeTable = sheet.tables.getItem("DatamergeTable");
    
    // Insert the count column
    var countCol = mergeTable.columns.add(0 /*index*/, null /*values*/, "count" /*name*/);
    // Add 1 to the first cell

    // If alternate is selected, add alternate column
    if (alt) {
      mergeTable.columns.add(0 /*index*/, null /*values*/, "alternate" /*name*/);
    }
    
    var tableData = mergeTable.getRange().load("values");
    // countRange.load("values");

    /**
     * The Sink
     */
    // var rowCount = mergeTable.rows.getCount()
    // countCol.load("values");

    return context.sync().then(function () {
      // Grab row count
      console.log(tableData.values)
      tableData.values[1][0] = "1";
      return context.sync();

    });

  }).catch(function(err){
    console.log(err)
    showError(err);
  });

}