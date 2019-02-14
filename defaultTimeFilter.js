'use strict';

// Wrap everything in an anonymous function to avoid polluting the global namespace
(function() {
    const fiscalYearStartMonth = '4';
    const fiscalYearStartMonthKey = 'fiscalYearStartMonth';
    let timefilters = [];
    let dataSources = {};
    // Use the jQuery document ready signal to know when everything has been initialized
    $(document).ready(function() {
        if(sessionStorage.extensionRun){return;}
        // Add your startup code here
        tableau.extensions.initializeAsync({ 'configure': configure }).then(function() {
            // Get the dashboard name from the tableau namespace and set it as our title
            const dashboardName = tableau.extensions.dashboardContent.dashboard.name;
            // $('#choose_sheet_title').text(dashboardName);

            // alert(dashboardName);
            fetchDataSources();

            fetchFilters();

            if (typeof(Storage) !== "undefined") {
                sessionStorage.extensionRun = 1;
            } else {
                // Sorry! No Web Storage support..
            }
        });
    });

    function fetchDataSources() {
        let dataSourceFetchPromises = [];
        const dashboard = tableau.extensions.dashboardContent.dashboard;

        dashboard.worksheets.forEach(function(worksheet) {
            dataSourceFetchPromises.push(worksheet.getDataSourcesAsync());
        });

        Promise.all(dataSourceFetchPromises).then(function(fetchResults) {
            fetchResults.forEach(function(dataSourcesForWorksheet) {
                dataSourcesForWorksheet.forEach(function(dataSource) {
                    if (!dataSources[dataSource.id]) {
                        dataSources[dataSource.id] = dataSource;
                    }
                });
            });
        });
    }

    function getFiscalStartMonthSetting() {
        let startMonth = tableau.extensions.settings.get(fiscalYearStartMonthKey);
        return startMonth;
    }

    function configure() {
        // This uses the window.location.origin property to retrieve the scheme, hostname, and 
        // port where the parent extension is currently running, so this string doesn't have
        // to be updated if the extension is deployed to a new location.
        const popupUrl = `${window.location.origin}/tbExtension/defaultTimeFilterDialog.html`;

        let startMonth = getFiscalStartMonthSetting();
        if (startMonth == undefined) {
            let minMonth;
            let targetRow;

            for (let i in dataSources) {
                dataSources[i].refreshAsync().then(function() {
                    dataSources[i].getUnderlyingDataAsync({
                        // columnsToInclude: ["Month Of Year", "Fiscal Quarter", "Fiscal Month Of Year"],
                        maxRows: 1000
                    }).then(dataTable => {
                        // console.log(dataTable);
                        let columnA = dataTable.columns.find(column => column.fieldName === "Month Of Year");
                        let columnB = dataTable.columns.find(column => column.fieldName === "Fiscal Quarter");
                        let columnC = dataTable.columns.find(column => column.fieldName === "Fiscal Month Of Year");
                        targetRow = dataTable.data.find(row => row[columnB.index].value == "Q1" && row[columnC.index].value == "1");
                        minMonth = targetRow[columnA.index].value;
                        tableau.extensions.settings.set(fiscalYearStartMonthKey, minMonth);
                        displayDialog(popupUrl, minMonth);
                    }).catch((error) => {
                        tableau.extensions.settings.set(fiscalYearStartMonthKey, fiscalYearStartMonth);
                        displayDialog(popupUrl, fiscalYearStartMonth);
                    });
                });
            }
        } else {
            displayDialog(popupUrl, startMonth);
        }
    }

    function displayDialog(popupUrl, startMonth) {
        /**
         * This is the API call that actually displays the popup extension to the user.  The
         * popup is always a modal dialog.  The only required parameter is the URL of the popup,  
         * which must be the same domain, port, and scheme as the parent extension.
         * 
         * The developer can optionally control the initial size of the extension by passing in 
         * an object with height and width properties.  The developer can also pass a string as the
         * 'initial' payload to the popup extension.  This payload is made available immediately to 
         * the popup extension.  In this example, the value '5' is passed, which will serve as the
         * default interval of refresh.
         */
        tableau.extensions.ui.displayDialogAsync(popupUrl, startMonth, { height: 500, width: 500 }).then((closePayload) => {
            // The promise is resolved when the dialog has been expectedly closed, meaning that
            // the popup extension has called tableau.extensions.ui.closeDialog.
            // $('#inactive').hide();
            // $('#active').show();

            refreshFiscalFilter(closePayload);

        }).catch((error) => {
            // One expected error condition is when the popup is closed by the user (meaning the user
            // clicks the 'X' in the top right of the dialog).  This can be checked for like so:
            switch (error.errorCode) {
                case tableau.ErrorCodes.DialogClosedByUser:
                    console.log("Dialog was closed by user");
                    break;
                default:
                    console.error(error.message);
            }
        });
    }

    function refreshFiscalFilter(startMonth) {
        // The close payload is returned from the popup extension via the closeDialog method.
        let lastCompleteFiscalQuarterMoment = getLastCompleteFiscalQuarter(startMonth);

        // let lastCompleteFiscalQuarter = '';
        // if (lastCompleteFiscalQuarterMoment.nextYear == null) {
        //     lastCompleteFiscalQuarter = lastCompleteFiscalQuarter.concat(lastCompleteFiscalQuarterMoment.year, '/Q', lastCompleteFiscalQuarterMoment.quarter);
        // } else {
        //     lastCompleteFiscalQuarter = lastCompleteFiscalQuarter.concat(lastCompleteFiscalQuarterMoment.nextYear, '/Q', lastCompleteFiscalQuarterMoment.quarter);
        // }
        let lastCompleteFiscalYear;
        let lastCompleteFiscalQuarter;
        if (lastCompleteFiscalQuarterMoment.nextYear == null) {
            lastCompleteFiscalYear = lastCompleteFiscalQuarterMoment.year;
            lastCompleteFiscalQuarter = 'Q'.concat(lastCompleteFiscalQuarterMoment.quarter);
        } else {
            lastCompleteFiscalYear = lastCompleteFiscalQuarterMoment.nextYear;
            lastCompleteFiscalQuarter = 'Q'.concat(lastCompleteFiscalQuarterMoment.quarter);
        }

        applyDefaultTimeFilter(lastCompleteFiscalYear, lastCompleteFiscalQuarter);
    }

    function getLastCompleteFiscalQuarter(startMonth) {
        return moment().subtract(3, 'months').fquarter(startMonth);
    }

    function fetchFilters() {
        // While performing async task, show loading message to user.
        $('#loading').addClass('show');

        // Whenever we restore the filters table, remove all save handling functions,
        // since we add them back later in this function.
        // unregisterHandlerFunctions.forEach(function (unregisterHandlerFunction) {
        //   unregisterHandlerFunction();
        // });

        // Since filter info is attached to the worksheet, we will perform
        // one async call per worksheet to get every filter used in this
        // dashboard.  This demonstrates the use of Promise.all to combine
        // promises together and wait for each of them to resolve.
        let filterFetchPromises = [];

        // List of all filters in a dashboard.
        let dashboardfilters = [];

        // To get filter info, first get the dashboard.
        const dashboard = tableau.extensions.dashboardContent.dashboard;

        // Then loop through each worksheet and get its filters, save promise for later.
        dashboard.worksheets.forEach(function(worksheet) {
            filterFetchPromises.push(worksheet.getFiltersAsync());

            //   // Add filter event to each worksheet.  AddEventListener returns a function that will
            //   // remove the event listener when called.
            //   let unregisterHandlerFunction = worksheet.addEventListener(tableau.TableauEventType.FilterChanged, filterChangedHandler);
            //   unregisterHandlerFunctions.push(unregisterHandlerFunction);
        });

        // Now, we call every filter fetch promise, and wait for all the results
        // to finish before displaying the results to the user.
        Promise.all(filterFetchPromises).then(function(fetchResults) {
            fetchResults.forEach(function(filtersForWorksheet) {
                // filtersForWorksheet.forEach(function (filter) {
                //   dashboardfilters.push(filter);
                // });
                filtersForWorksheet.filter(getYearQuarterFilters).forEach(function(filter) {
                    timefilters.push(filter);
                });
            });

            let startMonth = getFiscalStartMonthSetting();
            if (startMonth != undefined) {
                refreshFiscalFilter(startMonth);
            } else {
                alert('Please validate FiscalYearStartMonth in configuration dialog.');
            }
        });
    }

    function getYearQuarterFilters(filter) {
        let patt = /(Fiscal Year.*|Fiscal Quarter.*)/g;
        // let patt = /Fiscal Year Quarter/g;
        return patt.test(filter.fieldName)
    }

    function applyDefaultTimeFilter(defaultFiscalYear, defaultFiscalQuarter) {
        let defaultYearArray = [];
        let defaultQuarterArray = [];
        defaultYearArray.push(defaultFiscalYear);
        defaultQuarterArray.push(defaultFiscalQuarter);
        const dashboard = tableau.extensions.dashboardContent.dashboard;

        timefilters.forEach(function(filter) {
            const worksheetName = filter.worksheetName;
            const targetWorksheet = dashboard.worksheets.find(function(worksheet) {
                return worksheet.name == worksheetName;
            });

            // targetWorksheet.applyFilterAsync(filter.fieldName, defaultArray, tableau.FilterUpdateType.Replace);

            if (filter.fieldName.startsWith('Fiscal Year')) {
                targetWorksheet.applyFilterAsync(filter.fieldName, defaultYearArray, tableau.FilterUpdateType.Replace);
            }

            if (filter.fieldName.startsWith('Fiscal Quarter')) {
                targetWorksheet.applyFilterAsync(filter.fieldName, defaultQuarterArray, tableau.FilterUpdateType.Replace);
            }
        });
    }

    // Constructs UI that displays all the dataSources in this dashboard
    // given a mapping from dataSourceId to dataSource objects.
    function buildFiltersTable(filters) {
        // Clear the table first.
        $('#filtersTable > tbody tr').remove();
        const filtersTable = $('#filtersTable > tbody')[0];

        filters.forEach(function(filter) {
            let newRow = filtersTable.insertRow(filtersTable.rows.length);
            let nameCell = newRow.insertCell(0);
            let worksheetCell = newRow.insertCell(1);
            let typeCell = newRow.insertCell(2);
            let valuesCell = newRow.insertCell(3);

            const valueStr = getFilterValues(filter);

            nameCell.innerHTML = filter.fieldName;
            worksheetCell.innerHTML = filter.worksheetName;
            typeCell.innerHTML = filter.filterType;
            valuesCell.innerHTML = valueStr;
        });

        updateUIState(Object.keys(filters).length > 0);
    }

    // This helper updates the UI depending on whether or not there are filters
    // that exist in the dashboard.  Accepts a boolean.
    function updateUIState(filtersExist) {
        $('#loading').addClass('hidden');
        if (filtersExist) {
            $('#filtersTable').removeClass('hidden').addClass('show');
            $('#noFiltersWarning').removeClass('show').addClass('hidden');
        } else {
            $('#noFiltersWarning').removeClass('hidden').addClass('show');
            $('#filtersTable').removeClass('show').addClass('hidden');
        }
    }

    // This returns a string representation of the values a filter is set to.
    // Depending on the type of filter, this string will take a different form.
    function getFilterValues(filter) {
        let filterValues = '';

        switch (filter.filterType) {
            case 'categorical':
                filter.appliedValues.forEach(function(value) {
                    filterValues += value.formattedValue + ', ';
                });
                break;
            case 'range':
                // A range filter can have a min and/or a max.
                if (filter.minValue) {
                    filterValues += 'min: ' + filter.minValue.formattedValue + ', ';
                }

                if (filter.maxValue) {
                    filterValues += 'min: ' + filter.maxValue.formattedValue + ', ';
                }
                break;
            case 'relative-date':
                filterValues += 'Period: ' + filter.periodType + ', ';
                filterValues += 'RangeN: ' + filter.rangeN + ', ';
                filterValues += 'Range Type: ' + filter.rangeType + ', ';
                break;
            default:
        }

        // Cut off the trailing ", "
        return filterValues.slice(0, -2);
    }

    function loadSelectedMarks(worksheetName) {
        // For now, just pop up an alert saying that we've selected a sheet
        alert(`Loading selected marks for ${worksheetName}`);
    }
})();