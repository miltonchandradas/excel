sap.ui.define([], function () {

	"use strict";

	return {

        _createTable: function (id) {
			let table = document.createElement("table");
			table.id = id;
			table.style.cssText = "font-family: sans-serif; font-size: 0.9em; border: 1px solid rgb(174, 215, 255);";

			return table;
        },
        
        _addColumnHeaders: function (table, columnHeaders) {

			let tr = document.createElement("tr");
			tr.style.color = "#fff";
			tr.style.textAlign = "left";
			tr.style.fontWeight = "bold";

			// Add the column headers
			columnHeaders.forEach((header) => {
				let th = document.createElement("th");
				th.innerHTML = header;
				th.style.backgroundColor = "rgb(40, 110, 180)";
				th.style.padding = "12px 15px";
				th.style.height = "40px";
				th.style.border = "1px solid rgb(174, 215, 255)";
				tr.appendChild(th);
			});

			table.appendChild(tr);

			return tr;
		},

        createDOMTable: function (controller) {
            let table = this._createTable("domTable");
            let columnHeaders = [
                "Product Name",
                "Supplier ID",
                "Category ID",
                "Quantity Per Unit",
                "Unit Price",
                "Units In Stock",
                "Discontinued"

            ];
            let tr = this._addColumnHeaders(table, columnHeaders);

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