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
    // document.getElementById("sideload-msg").style.display = "none";
    // document.getElementById("app-body").style.display = "flex";
    document.getElementById("init").onclick = init;
    document.getElementById("insert").onclick = insert;
  }
  console.log("READY!");
});


/**
 * -------------------------------------------------------------------------------------------------- |
 * Initialize the list (Convert to Table)
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
          dataTable.name = "Datamerge Table";
          dataTable.rows.add(usedData.values);
          // console.log(usedData.values)
      });
    // console.log(`The range address was ${range.address}.`);

  }).catch(function(err){
    console.log("Error caught... displaying log");
    showError("This document has already been initialized...");
  });
  //#endregion
}

export async function insert() {
  console.log("RUN!");
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      var sheet = context.workbook.worksheets.getActiveWorksheet();
      // sheet.getRange("C:C").insert('right'); 
      var expensesTable = sheet.tables.add("A1:D1", true /*hasHeaders*/);
      expensesTable.name = "ExpensesTable";
      expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

      expensesTable.rows.add(null /*add rows to the end of the table*/, [
        ["1/1/2017", "The Phone Company", "Communications", "$120"],
        ["1/2/2017", "Northwind Electric Cars", "Transportation", "$142"],
        ["1/5/2017", "Best For You Organics Company", "Groceries", "$27"],
        ["1/10/2017", "Coho Vineyard", "Restaurant", "$33"],
        ["1/11/2017", "Bellows College", "Education", "$350"],
        ["1/15/2017", "Trey Research", "Other", "$135"],
        ["1/15/2017", "Best For You Organics Company", "Groceries", "$97"],
      ]);
      if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
      }

      sheet.activate();

      await context.sync();
    });
  } catch (error) {
    $("#log-error").text("Table already exists...")
    $("#log").show();
    console.error(error);
  }
}

$("#hide").on("click", function() {
  $("#log").hide();
})

$("#del").on("click", function() {

  Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    return context.sync()
        .then(function () {
            expensesTable.delete();
            $("#log").hide();
        });
  }).catch(function(){
    console.log("error")
  });
})


/**
 * Show Error Message
 */
function showError(theError) {
  //#region
  console.log("Inside showError...")
  $("#log-slide").fadeIn(); // Show the error log area
  $("#log-message").text(theError);
  //#endregion
}

/**
 * EVENT - Close the error message
 */
$("#log-dismiss").on("click", function() {
  $("#log-slide").fadeOut(function() {
    // Done fading
    $("#log-message").empty();
  });

})