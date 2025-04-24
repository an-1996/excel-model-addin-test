/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office, OfficeExtension */

/**
 * Initializes the add-in. Attaches event handlers.
 */
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Assign function to button click
    document.getElementById("calculateGpmButton").onclick = runGpmCalculation;

    // Clear status on load
    updateStatus("");
  }
});

/**
 * Handles the button click event to calculate Gross Profit Margin.
 */
async function runGpmCalculation() {
  try {
    // Get the period input value from the text box
    const periodInput = document.getElementById("periodInput").value;
    if (!periodInput || periodInput.trim() === "") {
       updateStatus("Error: Please enter a period.", true);
       return;
    }
    const trimmedPeriod = periodInput.trim();

    updateStatus(`Calculating GPM for period ${trimmedPeriod}...`);

    // Run the Excel calculation logic
    await Excel.run(async (context) => {
      // --- 1. Get Data ---
      // Get the specific worksheet named "DataSheet"
      const sheet = context.workbook.worksheets.getItem("DataSheet");
      // Get the specific table named "InputData" within that sheet
      const table = sheet.tables.getItem("InputData");

      // Load the values from the data body range of the table
      const dataRange = table.getDataBodyRange();
      dataRange.load("values");

      // Synchronize with Excel to fetch the loaded values
      await context.sync();

      // --- 2. Process Data ---
      // Store the retrieved values in a variable
      const tableValues = dataRange.values;
      // Initialize variables to hold aggregated values
      let revenue = 0;
      let cogs = 0;
      let foundRevenue = false;
      let foundCogs = false;

      // Loop through each row in the retrieved table data
      // Assuming columns: Account (index 0), Period (index 1), Amount (index 2)
      for (let i = 0; i < tableValues.length; i++) {
        const row = tableValues[i];
        // Extract data, trim strings, and handle potential nulls/undefined
        const currentPeriod = row[1] ? String(row[1]).trim() : ""; // Period in column 1
        const currentAccount = row[0] ? String(row[0]).trim().toLowerCase() : ""; // Account in column 0
        // Ensure amount is treated as a number, default to 0 if not numeric
        const currentAmount = typeof row[2] === 'number' ? row[2] : 0; // Amount in column 2

        // Check if the current row's period matches the user's input period
        if (currentPeriod === trimmedPeriod) {
          // Check the account type and aggregate the amount accordingly
          switch (currentAccount) {
            case "revenue":
              revenue += currentAmount;
              foundRevenue = true; // Mark that revenue data was found
              break;
            case "cogs":
              cogs += currentAmount;
              foundCogs = true; // Mark that COGS data was found
              break;
          }
        }
      }

      // --- 3. Validate and Calculate ---
      // Check if essential data points were found for the calculation
      if (!foundRevenue) {
        updateStatus(`Error: No 'Revenue' found for period ${trimmedPeriod}.`, true);
        return; // Stop execution if revenue is missing
      }
      if (!foundCogs) {
        updateStatus(`Error: No 'COGS' found for period ${trimmedPeriod}.`, true);
        return; // Stop execution if COGS is missing
      }
      // Check for non-positive revenue to prevent division by zero or meaningless margin
      if (revenue <= 0) {
         updateStatus(`Error: Revenue (${revenue}) is zero or negative for period ${trimmedPeriod}. Cannot calculate margin.`, true);
         return; // Stop execution if revenue is not positive
      }

      // Perform the Gross Profit and Margin calculations
      const grossProfit = revenue - cogs;
      const grossProfitMargin = grossProfit / revenue;

      // --- 4. Prepare and Write Output ---
      // Define the name for the new output worksheet
      const outputSheetName = `GPM_${trimmedPeriod}`;

      // Check if an output sheet with the same name already exists
      const existingSheet = context.workbook.worksheets.getItemOrNullObject(outputSheetName);
      // Load a property (like 'name') to check its existence after the sync
      existingSheet.load("name");
      await context.sync(); // Sync to check if the sheet exists

      // If the sheet exists (is not a null object), delete it
      if (!existingSheet.isNullObject) {
          console.log(`Deleting existing sheet: ${outputSheetName}`);
          existingSheet.delete();
          // Sync again immediately to ensure deletion completes before adding the new sheet
          await context.sync();
      }

      // Add the new output sheet
      const outputSheet = context.workbook.worksheets.add(outputSheetName);

      // Prepare the data array to be written to the output sheet.
      // **Ensure this array matches the dimensions of the target range (A1:B7 -> 7 rows, 2 columns)**
      const outputData = [
        ["Period:", trimmedPeriod],         // Row 1
        [null, null],                     // Row 2 (blank)
        ["Metric", "Value"],              // Row 3
        ["Revenue", revenue],             // Row 4
        ["COGS", cogs],                   // Row 5
        ["Gross Profit", grossProfit],     // Row 6
        ["Gross Profit Margin", grossProfitMargin] // Row 7
      ];

      // Get the target range on the output sheet
      const outputRange = outputSheet.getRange("A1:B7");
      // Write the prepared data array to the range
      outputRange.values = outputData;

      // Apply formatting to the output range
      outputSheet.getRange("B4:B6").numberFormat = "$#,##0.00"; // Currency format
      outputSheet.getRange("B7").numberFormat = "0.00%"; // Percentage format
      
      // Apply bold formatting to header rows separately
      outputSheet.getRange("A1:B1").format.font.bold = true; 
      outputSheet.getRange("A3:B3").format.font.bold = true; 

      // Autofit columns A and B for better readability using the outputRange object
      outputRange.format.autofitColumns(); 

      // Make the newly created output sheet the active sheet
      outputSheet.activate();

      // Sync all queued changes (add sheet, write data, format, activate) to Excel
      await context.sync();

      // Update the status message in the task pane to indicate completion
      updateStatus(`Calculation complete. Results on sheet '${outputSheetName}'.`);

    }); // End of Excel.run
  } catch (error) {
    // Catch any errors that occur during the process
    console.error("Error: " + error); // Log the full error to the console
    // Check if it's an OfficeExtension specific error for more details
    if (error instanceof OfficeExtension.Error) {
      console.error("Debug info: " + JSON.stringify(error.debugInfo));
      // Display a user-friendly error message in the task pane
      updateStatus(`Error: ${error.message} (Debug: ${error.debugInfo.errorLocation || 'N/A'})`, true);
    } else {
      // Display a generic error message for other types of errors
       updateStatus(`Error: ${error.message}`, true);
    }
  }
}

/**
 * Helper function to update the status message in the task pane.
 * @param message The message to display.
 * @param isError Optional. If true, applies error styling.
 */
function updateStatus(message, isError = false) {
   const statusElement = document.getElementById("status");
   // Set the text content of the status element
   statusElement.innerText = message;
   // Apply red color for errors, default color otherwise
   statusElement.style.color = isError ? "red" : "#555";
}
