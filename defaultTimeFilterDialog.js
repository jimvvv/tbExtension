'use strict';

// Wrap everything in an anonymous function to avoid polluting the global namespace
(function() {
    /**
     * This extension collects the IDs of each datasource the user is interested in
     * and stores this information in settings when the popup is closed.
     */
    const datasourcesSettingsKey = 'selectedDatasources';
    const fiscalYearStartMonthKey = 'fiscalYearStartMonth';

    let selectedDatasources = [];

    $(document).ready(function() {
        // The only difference between an extension in a dashboard and an extension 
        // running in a popup is that the popup extension must use the method
        // initializeDialogAsync instead of initializeAsync for initialization.
        // This has no affect on the development of the extension but is used internally.
        tableau.extensions.initializeDialogAsync().then(function(openPayload) {
            $('#startMonth').val(openPayload);
            $('#closeButton').click(closeDialog);

            // let dashboard = tableau.extensions.dashboardContent.dashboard;
            // let visibleDatasources = [];
            // selectedDatasources = parseSettingsForActiveDataSources();

            // // Loop through datasources in this sheet and create a checkbox UI 
            // // element for each one.  The existing settings are used to 
            // // determine whether a datasource is checked by default or not.
            // dashboard.worksheets.forEach(function (worksheet) {
            //   worksheet.getDataSourcesAsync().then(function (datasources) {
            //     datasources.forEach(function (datasource) {
            //       let isActive = (selectedDatasources.indexOf(datasource.id) >= 0);

            //       if (visibleDatasources.indexOf(datasource.id) < 0) {
            //         addDataSourceItemToUI(datasource, isActive);
            //         visibleDatasources.push(datasource.id);
            //       }
            //     });
            //   });
            // });
        });
    });

    /**
     * Helper that parses the settings from the settings namesapce and 
     * returns a list of IDs of the datasources that were previously
     * selected by the user.
     */
    function parseSettingsForActiveDataSources() {
        let activeDatasourceIdList = [];
        let settings = tableau.extensions.settings.getAll();
        if (settings.selectedDatasources) {
            activeDatasourceIdList = JSON.parse(settings.selectedDatasources);
        }

        return activeDatasourceIdList;
    }

    /**
     * Helper that updates the internal storage of datasource IDs
     * any time a datasource checkbox item is toggled.
     */
    function updateDatasourceList(id) {
        let idIndex = selectedDatasources.indexOf(id);
        if (idIndex < 0) {
            selectedDatasources.push(id);
        } else {
            selectedDatasources.splice(idIndex, 1);
        }
    }    

    /**
     * Stores the selected datasource IDs in the extension settings,
     * closes the dialog, and sends a payload back to the parent. 
     */
    function closeDialog() {
        let currentSettings = $('#startMonth').val()
        tableau.extensions.settings.set(fiscalYearStartMonthKey, currentSettings);

        tableau.extensions.settings.saveAsync().then((newSavedSettings) => {
            tableau.extensions.ui.closeDialog(currentSettings);
        });
    }
})();