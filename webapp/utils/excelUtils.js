sap.ui.define([], function () {

	"use strict";

	return {

        _createTable: function (id) {
			let table = document.createElement("table");
			table.id = id;
			table.style.cssText = "font-family: sans-serif; font-size: 0.9em; border: 1px solid rgb(174, 215, 255);";

			return table;
		},

        createDOMTable: function (controller) {
            let table = this._createTable("domTable");
            return table;
        },

        writeToExcel: function (table, filename) {

			let dataType = "application/vnd.ms-excel";
			let tableHTML = table.outerHTML.replace(/ /g, "%20");

			// Create download link element
			let downloadLink = document.createElement("a");
			downloadLink.id = "domHyperlink";

			document.body.appendChild(downloadLink);

			if (navigator.msSaveOrOpenBlob) {
				let blob = new Blob(["\ufeff", tableHTML], {
					type: dataType
				});
				navigator.msSaveOrOpenBlob(blob, filename);
			} else {
				// Create a link to the file
				downloadLink.href = "data:" + dataType + ", " + tableHTML;

				// Setting the file name
				downloadLink.download = filename;

				//triggering the function
				downloadLink.click();
			}
		},

    };
});