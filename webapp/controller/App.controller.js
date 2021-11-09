sap.ui.define(
  ["sap/ui/core/mvc/Controller",
"com/sap/excel/utils/excelUtils"],
  /**
   * @param {typeof sap.ui.core.mvc.Controller} Controller
   */
  function (Controller, excelUtils) {
    "use strict";

    return Controller.extend("com.sap.excel.controller.App", {
      onInit: function () {
        this._mainModel = this.getOwnerComponent().getModel();
      },

      onExcelDownload: async function (oEvent) {
        // Remove any existing DOM artifacts...
        let domHyperlink = document.getElementById("domHyperlink");
        if (domHyperlink) domHyperlink.remove();

        // Set busy indicator to true
        this.getView().setBusy(true);

        // Create DOM table
        // Populate and format DOM table with data
        let table = await excelUtils.createDOMTable(this);
        
        // Write DOM table to Excel
        excelUtils.writeToExcel(table, "northwind")

        // Set busy indicator to false
        this.getView().setBusy(false);

      },
    });
  }
);
