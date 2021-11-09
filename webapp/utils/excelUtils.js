sap.ui.define(["com/sap/excel/utils/odataUtils"], function (odataUtils) {
  "use strict";

  return {
    _createTable: function (id) {
      let table = document.createElement("table");
      table.id = id;
      table.style.cssText =
        "font-family: sans-serif; font-size: 0.9em; border: 1px solid rgb(174, 215, 255);";

      return table;
    },

    _createHeaderSection: function () {
      let headerTbody = document.createElement("tbody");
      headerTbody.style.border = "1px solid black";
      let headerTr = headerTbody.insertRow();
      let headerCell0 = headerTr.insertCell();
      headerCell0.innerHTML = `<b>OData Service:</b> Northwind OData Service`;
      headerCell0.colSpan = 2;
      headerTr = headerTbody.insertRow();
      let headerCell1 = headerTr.insertCell();
      headerCell1.innerHTML = `Excel Download with formatting in SAPUI5`;
      headerCell1.colSpan = 2;
      headerTr = headerTbody.insertRow();
      let headerCell2 = headerTr.insertCell();

      return headerTbody;
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

    _setCellStyle: function (cell) {
      cell.style.padding = "12px 15px";
      cell.style.border = "1px solid rgb(174, 215, 255)";

      cell.style.textAlign = "center";
      cell.style.verticalAlign = "middle";
    },

    _insertRowValues: function (fieldNames, cells, index, result) {
      fieldNames.forEach((fieldName) => {
        switch (fieldName) {
          case "ProductID": {
            let cellValue = `${result.ProductID}<br>`;
            cells[0].innerHTML = cellValue;

            if (index % 2 === 0) {
              cells.forEach((cell) => (cell.style.backgroundColor = "#f0f0f0"));
            }

            break;
          }
          case "ProductName": {
            let cellValue = `<b>${result.ProductName}</b><br>`;
            cells[1].innerHTML = cellValue;
            break;
          }
          case "SupplierID": {
            let cellValue = `${result.SupplierID}<br>`;
            cells[2].innerHTML = cellValue;
            break;
          }
          case "CategoryID": {
            let cellValue = `${result.CategoryID}<br>`;
            cells[3].innerHTML = cellValue;
            break;
          }
          case "QuantityPerUnit": {
            let cellValue = `${result.QuantityPerUnit}<br>`;
            cells[4].innerHTML = cellValue;
            break;
          }
          case "UnitPrice": {
            let cellValue = `${result.UnitPrice}<br>`;
            cells[5].innerHTML = cellValue;
            break;
          }
          case "UnitsInStock": {
            let cellValue = `${result.UnitsInStock}<br>`;
            cells[6].innerHTML = cellValue;
            break;
          }
          case "Discontinued": {
            let cellValue = `${
              result.Discontinued ? "Discontinued" : "Available"
            }<br>`;
            if (result.Discontinued) cells[7].style.backgroundColor = "yellow";
            cells[7].innerHTML = cellValue;
            break;
          }
          default:
            break;
        }
      });
    },

    createDOMTable: async function (controller) {
      let table = this._createTable("domTable");

      // Add the header section on top...
      let headerSection = this._createHeaderSection();
      table.appendChild(headerSection);

      let columnHeaders = [
        "Product ID",
        "Product Name",
        "Supplier ID",
        "Category ID",
        "Quantity Per Unit",
        "Unit Price",
        "Units In Stock",
        "Discontinued",
      ];
      let tr = this._addColumnHeaders(table, columnHeaders);

      // Add JSON data as rows to the table
      let tbody = document.createElement("tbody");

      // Retrieve data from backend...
      let data = await odataUtils.readFromBackend(
        "Products",
        controller._mainModel
      );

      if (data && data.results && data.results.length > 0) {
        let fieldNames = Object.keys(data.results[0]);

        data.results.forEach((result, index) => {
          tr = tbody.insertRow();

          let cells = [];
          columnHeaders.forEach((columnHeader, i) => {
            cells[i] = tr.insertCell();
            this._setCellStyle(cells[i]);
          });

          this._insertRowValues(fieldNames, cells, index, result);
        });
      }

      table.appendChild(tbody);

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
          type: dataType,
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
