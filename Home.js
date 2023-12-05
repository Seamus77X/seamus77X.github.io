function helloWorld() {
    console.log('hello')
}

(function () {
    "use strict";

    // Declaration of global variables for later use
    let messageBanner;
    let dialog
    let accessToken;  // used to store user's access token
    let LessonsTable  // used to stored lessons learned data in memory

    // Constants for client ID, redirect URL, and resource domain for authentication
    const clientId = "be63874f-f40e-433a-9f35-46afa1aef385"
    const redirectUrl = "https://seamus77x.github.io/index.html"
    const resourceDomain = "https://gsis-pmo-australia-sensei-dev.crm6.dynamics.com/"

    // Initialization function that runs each time a new page is loaded.
    Office.initialize = function (reason) {
        $(function () {

            Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
            Office.context.document.settings.saveAsync();
            Office.addin.setStartupBehavior(Office.StartupBehavior.load);

            try {
                // Notification mechanism initialization and hiding it initially
                let element = document.querySelector('.MessageBanner');
                messageBanner = new components.MessageBanner(element);
                messageBanner.hideBanner();

                // Fallback logic for versions of Excel older than 2016
                if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                    throw new Error("Sorry, this add-in only works with newer versions of Excel.")
                }

                // add external js
                //$('#myScriptX').attr('src', 'Test.js')
                //$.getScript('Test.js', function () {
                //    externalFun()
                //})

                // UI text setting for buttons and descriptions
                $('#button1-text').text("Download");
                $("#button1").attr("title", "Load Data to Excel")
                $('#button1').on("click", loadSampleData);

                $('#button2-text').text("Button 2");

                // Authentication and access token retrieval logic
                if (typeof accessToken === 'undefined') {
                    // Constructing authentication URL
                    let authUrl = "https://login.microsoftonline.com/common/oauth2/authorize" +
                        "?client_id=" + clientId +
                        "&response_type=token" +
                        "&redirect_uri=" + redirectUrl +
                        "&response_mode=fragment" +
                        "&resource=" + resourceDomain;

                    // Displaying authentication dialog
                    Office.context.ui.displayDialogAsync(authUrl, { height: 30, width: 30, requireHTTPS: true },
                        function (result) {
                            if (result.status === Office.AsyncResultStatus.Failed) {
                                // If the dialog fails to open, throw an error
                                throw new Error("Failed to open dialog: " + result.error.message);
                            }
                            dialog = result.value;
                            dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                        }
                    );
                }
            } catch (error) {
                errorHandler(error.message)
            }
        });
    }

    Office.actions.associate("buttonFunction", function (event) {
        console.log('Hey, you just pressed a button in Excel ribbon. Test')
        console.log(accessToken)
        event.completed();
    })

    // Process message (access token) received from the dialog
    function processMessage(arg) {
        try {
            // Check if the message is present
            if (!arg.message) {
                throw new Error("No message received from the dialog.");
            }

            // Parse the JSON message received from the dialog
            const response = JSON.parse(arg.message);

            // Check the status of the response
            if (response.Status === "Success") {
                // store the token in memory for later use
                accessToken = response.AccessToken
                console.log("Access Token Received")
            } else if (response.Status === "Error") {
                // Handle the error scenario
                errorHandler(response.Message || "An error occurred.");
            } else {
                // Handle unexpected status
                errorHandler("Unexpected response status.");
            }

        } catch (error) {
            // Handle any errors that occur during processing
            errorHandler(error.message);
        } finally {
            // Close the dialog, regardless of whether an error occurred
            if (dialog) {
                dialog.close();
            }
        }
    }

    // Function to load sample data
    async function loadSampleData() {
        await loadData(`${resourceDomain}api/data/v9.1/sensei_lessonslearned`
            , 'Sheet1', 'A1', 'sensei_lessonslearned')

        registerTableChangeEvent('sensei_lessonslearned')
    }
    //sc_integrationrecentgranulartransactions
    //sensei_financialtransaction
    //sensei_financialtransactions?$select=sc_kbrkey,sc_vendorname,sensei_value,sc_docdate,sensei_financialtransactionid&$top=50000

    // Function to retrieve data from Dynamics 365
    async function loadData(resourceUrl, defaultSheet, defaultTpLeftRng, tableName) {
        try {
            const DataArr = await Read_D365(resourceUrl);

            // report an error and interupt if failed to read data from Dataverse
            if (!DataArr || DataArr.length === 0) {
                throw new Error("No data retrieved or data array is empty");
            }

            // paste data into Excel worksheet 
            await Excel.run(async (ctx) => {
                const ThisWorkbook = ctx.workbook;
                const Worksheets = ThisWorkbook.worksheets;
                ctx.application.calculationMode = Excel.CalculationMode.manual;
                Worksheets.load("items/tables/items/name");

                await ctx.sync();

                let tableFound = false;
                let table;
                let oldRangeAddress;
                let sheet

                if (typeof tableName !== 'undefined') {

                    // Attempt to find the existing table.
                    for (sheet of Worksheets.items) {
                        const tables = sheet.tables;

                        // Check if the table exists in the current sheet
                        table = tables.items.find(t => t.name === tableName);

                        // if the table found, delete the existing data
                        if (table) {
                            tableFound = true;
                            // Clear the data body range.
                            const dataBodyRange = table.getDataBodyRange();
                            dataBodyRange.load("address");
                            //dataBodyRange.clear();
                            await ctx.sync();
                            // Load the address of the range for new data insertion.
                            oldRangeAddress = dataBodyRange.address.split('!')[1];
                            break;
                        }
                    }

                    if (tableFound) {
                        // keep first data row for customised function on LHS and RHS
                        const oldAddressWithouutFirstRow = oldRangeAddress.replace(/\d+/, parseInt(oldRangeAddress.match(/\d+/)[0], 10) + 1)
                        sheet.getRange(oldAddressWithouutFirstRow).clear()
                        // Situation 1: Insert new data into the cleared data body range.
                        const startCell = oldRangeAddress.split(":")[0]
                        const endCell = oldRangeAddress.replace(/\d+$/, parseInt(oldRangeAddress.match(/\d+/)[0], 10) + DataArr.length - 2).split(":")[1]
                        const range = sheet.getRange(`${startCell}:${endCell}`);
                        DataArr.shift()
                        range.values = DataArr;

                        // include header row when resize
                        const startCellWithHeader = oldRangeAddress.replace(/\d+/, parseInt(oldRangeAddress.match(/\d+/)[0], 10) - 1).split(":")[0]
                        const WholeTabkeRange = sheet.getRange(`${startCellWithHeader}:${endCell}`)
                        table.resize(WholeTabkeRange)

                        range.format.autofitColumns();
                        range.format.autofitRows();
                    } else {
                        // Situation 2: If the table doesn't exist, create a new one.
                        let tgtSheet = Worksheets.getItem(defaultSheet);
                        let endCellCol = columnNumberToName(columnNameToNumber(defaultTpLeftRng.replace(/\d+$/, "")) - 1 + DataArr[0].length)
                        let endCellRow = parseInt(defaultTpLeftRng.match(/\d+$/)[0], 10) + DataArr.length - 1
                        const rangeAddress = defaultTpLeftRng + ":" + endCellCol + endCellRow;
                        const range = tgtSheet.getRange(rangeAddress);
                        range.values = DataArr;
                        const newTable = tgtSheet.tables.add(rangeAddress, true /* hasHeaders */);
                        newTable.name = tableName;

                        newTable.getRange().format.autofitColumns();
                        newTable.getRange().format.autofitRows();
                    }

                } else {
                    // Situation 3: paste the data in sheet directly, no table format
                    let tgtSheet = Worksheets.getItem(defaultSheet);
                    let endCellCol = columnNumberToName(columnNameToNumber(defaultTpLeftRng.replace(/\d+$/, "")) - 1 + DataArr[0].length)
                    let endCellRow = parseInt(defaultTpLeftRng.match(/\d+$/)[0], 10) + DataArr.length - 1
                    const rangeAddress = defaultTpLeftRng + ":" + endCellCol + endCellRow;
                    const range = tgtSheet.getRange(rangeAddress);
                    range.values = DataArr;

                    range.format.autofitColumns();
                    range.format.autofitRows();
                }

                await ctx.sync();
            })  // end of pasting data
        } catch (error) {
            errorHandler(error.message)
        } finally {
            await Excel.run(async (ctx) => {
                ctx.application.calculationMode = Excel.CalculationMode.automatic;
                await ctx.sync()
            })
        }
    }
    async function updateData() {
        Update_D365('sensei_lessonslearned', '0f0db491-3421-ee11-9966-000d3a798402', { 'sc_additionalcommentsnotes': 'Update Test' })
        //Create_D365('sensei_lessonslearned', { 'sensei_name': 'Add Test', 'sc_additionalcommentsnotes': 'ADD test from Web Add-In' })
        //Delete_D365('sensei_lessonslearned','f38edda5-8d8d-ee11-be35-6045bd3db52a')
    }

    // Function to create data in Dynamics 365
    async function Create_D365(entityLogicalName, addedData) {
        const url = `${resourceDomain}api/data/v9.1/${entityLogicalName}`;

        try {
            const response = await fetch(url, {
                method: 'POST',
                headers: {
                    'OData-MaxVersion': '4.0',
                    'OData-Version': '4.0',
                    'Accept': 'application/json',
                    'Content-Type': 'application/json; charset=utf-8',
                    'Authorization': `Bearer ${accessToken}`,
                    'Prefer': 'return=representation'
                },
                body: JSON.stringify(addedData)
            });

            if (!response.ok) {
                const errorData = await response.json();
                throw new Error(`Server responded with status ${response.status}: ${errorData.error?.message}`);
            }

            const responseData = await response.json();
            console.log("Record added successfully. New record ID:");
            //console.log(JSON.stringify(responseData))
            return responseData
        } catch (error) {
            if (error.name === 'TypeError') {
                // Handle network errors (e.g., no internet connection)
                errorHandler("Network error: " + error.message);
            } else {
                // Handle other types of errors (e.g., server responded with error code)
                errorHandler("Error encountered when adding new records in Dataverse:" + error.message);
            }
        }
    }
    // Function to read data in Dynamics 365
    async function Read_D365(url) {
        let totalRecords = 0;
        let finalArr = [];
        let startTime = new Date().getTime();

        try {
            do {
                let response = await fetch(url, {
                    method: 'GET',
                    headers: {
                        'OData-MaxVersion': '4.0',
                        'OData-Version': '4.0',
                        'Accept': 'application/json',
                        'Content-Type': 'application/json; charset=utf-8',
                        'Authorization': `Bearer ${accessToken}`,
                    }
                });

                if (!response.ok) {
                    const errorData = await response.json();
                    throw new Error(`Server responded with status ${response.status}: ${errorData.error?.message}`);
                }

                let jsonObj = await response.json();
                let headers = [];
                let tempArr_5k = [];

                if (jsonObj["value"] && jsonObj["value"].length > 0) {
                    for (let fieldName in jsonObj["value"][0]) {
                        if (typeof jsonObj["value"][0][fieldName] === "object" && jsonObj["value"][0][fieldName] != null) {
                            for (let relatedField in jsonObj["value"][0][fieldName]) {
                                let expandedFieldName = `${fieldName} / ${relatedField}`;
                                headers.push(expandedFieldName);
                            }
                        } else {
                            headers.push(fieldName);
                        }
                    }

                    tempArr_5k = [headers];

                    jsonObj["value"].forEach((row) => {
                        let itemWithRelatedFields = {};

                        for (let cell in row) {
                            if (typeof row[cell] === "object" && row[cell] !== null) {
                                for (let field in row[cell]) {
                                    let relatedFieldName = `${cell} / ${field}`;
                                    itemWithRelatedFields[relatedFieldName] = row[cell][field];
                                }
                            } else {
                                itemWithRelatedFields[cell] = row[cell];
                            }
                        }

                        let tempValRow = headers.map((header) => {
                            return itemWithRelatedFields[header] || null;
                        });

                        tempArr_5k.push(tempValRow);
                    });

                    if (totalRecords >= 1) {

                        let tempArr = [];
                        let headerRow = tempArr_5k[0];

                        for (let row of tempArr_5k) {
                            let tempValRow = [];
                            for (let fieldName of finalArr[0]) {
                                let trueColNo = headerRow.indexOf(fieldName);
                                tempValRow.push(row[trueColNo] || null);
                            }
                            tempArr.push(tempValRow);
                        }

                        tempArr.splice(0, 1);
                        finalArr = finalArr.concat(tempArr);
                    } else {
                        finalArr = finalArr.concat(tempArr_5k);
                    }
                }

                if (jsonObj["@odata.nextLink"]) {
                    url = jsonObj["@odata.nextLink"];
                } else {
                    url = null; // No more pages to retrieve
                }

                totalRecords += 1;
                console.log('HTTP Status Code: ' + response.status + ' - Page: ' + totalRecords);

            } while (url != null);

            // Update Excel with the collected data
            if (finalArr.length > 0) {
                let finishTime = new Date().getTime();
                console.log(`${(finishTime - startTime) / 1000} s used to download ${finalArr.length} records with ${finalArr[0].length} cols.`);

                return finalArr
            } else {
                throw new Error("No data downloaded");
            }


        } catch (error) {
            if (error.name === 'TypeError') {
                // Handle network errors (e.g., no internet connection)
                errorHandler("Network error: " + error.message);
            } else {
                // Handle other types of errors (e.g., server responded with error code)
                errorHandler("Error encountered when retrieving records from Dataverse:" + error.message);
            }
        }
    }
    // Function to update data in Dynamics 365
    async function Update_D365(entityLogicalName, recordId, updatedData) {
        const url = `${resourceDomain}api/data/v9.1/${entityLogicalName}(${recordId})`;

        try {
            const response = await fetch(url, {
                method: 'PATCH',
                headers: {
                    'OData-MaxVersion': '4.0',
                    'OData-Version': '4.0',
                    'Accept': 'application/json',
                    'Content-Type': 'application/json; charset=utf-8',
                    'Authorization': `Bearer ${accessToken}`,
                    //'Prefer': 'return=representation'
                },
                body: JSON.stringify(updatedData)
            });

            if (!response.ok) {
                // If the server responded with a non-OK status, handle the error
                const errorData = await response.json();
                throw new Error(`Server responded with status ${response.status}: ${errorData.error?.message}`);
            }

            console.log(`Record updated successfully. Updated record ID: [${recordId}]`);
        } catch (error) {
            if (error.name === 'TypeError') {
                // Handle network errors (e.g., no internet connection)
                errorHandler("Network error: " + error.message);
            } else {
                // Handle other types of errors (e.g., server responded with error code)
                errorHandler("Error encountered when updating records in Dataverse" + error.message);
            }
        }
    }
    // Function to delete data in Dynamics 365
    async function Delete_D365(entityLogicalName, recordId) {
        const url = `${resourceDomain}api/data/v9.1/${entityLogicalName}(${recordId})`;

        try {
            const response = await fetch(url, {
                method: 'DELETE',
                headers: {
                    'OData-MaxVersion': '4.0',
                    'OData-Version': '4.0',
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json; charset=utf-8'
                }
            });

            if (!response.ok) {
                const errorData = await response.json();
                throw new Error(`Server responded with status ${response.status}: ${errorData.error?.message}`);
            }

            console.log(`Record with ID [${recordId}] deleted successfully.`);
        } catch (error) {
            if (error.name === 'TypeError') {
                // Handle network errors (e.g., no internet connection)
                errorHandler("Network error: " + error.message);
            } else {
                // Handle other types of errors (e.g., server responded with error code)
                errorHandler("Error encountered when deleting new records in Dataverse:" + error.message);
            }
        }
    }

    // Progress bar update function
    function updateProgressBar(progress) {
        let elem = document.getElementById("myProgressBar");
        elem.style.width = progress + '%';
        //elem.innerHTML = progress + '%';
    }
    //// Example: Update the progress bar every second
    //let progress = 0;
    //let interval = setInterval(function () {
    //    progress += 10; // Increment progress
    //    updateProgressBar(progress);

    //    if (progress >= 100) clearInterval(interval); // Clear interval at 100%
    //}, 1000);


    // Utility function to convert column number to name
    function columnNumberToName(columnNumber) {
        let columnName = "";
        while (columnNumber > 0) {
            let remainder = (columnNumber - 1) % 26;
            columnName = String.fromCharCode(65 + remainder) + columnName;
            columnNumber = Math.floor((columnNumber - 1) / 26);
        }
        return columnName;
    }
    // Utility function to convert column name to number
    function columnNameToNumber(columnName) {
        let columnNumber = 0;
        for (let i = 0; i < columnName.length; i++) {
            columnNumber *= 26;
            columnNumber += columnName.charCodeAt(i) - 64;
        }
        return columnNumber;
    }

    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.error("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }

    async function registerTableChangeEvent(tableName) {

        console.log(`I am tracking the changes in ${tableName}`)






        //Excel.run(function (context) {
        //    var sheet = context.workbook.worksheets.getActiveWorksheet();

        //    var table = sheet.tables.getItem("sensei_lessonslearned");

        //    var headerRange = table.getHeaderRowRange();
        //    headerRange.load("values, cellCount, id");

        //    return context.sync()
        //        .then(function () {
        //            for (var i = 0; i < headerRange.values[0].length; i++) {
        //                console.log("Header cell value: " + headerRange.values[0][i] + headerRange.clientId);
        //            }
        //        });
        //}).catch(function (error) {
        //    console.error("Error: " + error);
        //    if (error instanceof OfficeExtension.Error) {
        //        console.error("Debug info: " + JSON.stringify(error.debugInfo));
        //    }
        //});









        let ThisWorkbook;
        let Worksheets;
        let tableFound = false;

        Excel.run(function (ctx) {
            try {
                ThisWorkbook = ctx.workbook;
                Worksheets = ThisWorkbook.worksheets;
                Worksheets.load("items/tables/items/name");
                return ctx.sync().then(() => {
                    for (let sheet of Worksheets.items) {
                        const tables = sheet.tables;
                        // Check if the 'Test' table exists in the current sheet
                        let table = tables.items.find(t => t.name === tableName);

                        if (table) {
                            // if the table found, then listen to the change in the table
                            table.onChanged.add(handleTableChange);
                            tableFound = true;
                            break;
                        }
                    }

                    if (!tableFound) {
                        // if the table not found, then raise an error
                        throw new Error(`[${tableName}] table is not found in Excel`);
                    }
                }).then(ctx.sync);
            } catch (error) {
                // Error handling for issues within the Excel.run block
                errorHandler("Error in registerTableChangeEvent: " + error.message);
            }
        }).catch(function (error) {
            // Error handling for issues related to Excel.run itself
            errorHandler("Error in Excel.run: " + error.message);
        });
    }

    // hanle table change.    tip: get after value from Excel if multiple range changes
    function handleTableChange(eventArgs) {

        switch (eventArgs.changeType) {
            case 'RangeEdited':
                console.log(`Range [${eventArgs.address}] was just updated.`)
                break;
            case "RowInserted":
                console.log(`Row [${eventArgs.address}] was just inserted.`)
                break;
            case "RowDeleted":
                console.log(`Row [${eventArgs.address}] was just deleted.`)
                break;
            case "ColumnInserted":
                console.log(`Column [${eventArgs.address}] was just inserted.`)
                break;
            case "ColumnDeleted":
                console.log(`Column [${eventArgs.address}] was just deleted.`)
                break;
            case "CellInserted":
                console.log(`Cell [${eventArgs.address}] was just inserted.`)
                break;
            case "CellDeleted":
                console.log(`Cell [${eventArgs.address}] was just deleted.`)
                break;
            default:
                console.log(`Unknown action.`)
                break;
        }

    }







})();

