sap.ui.define(
  [
    "com/sap/excel/utils/odataUtils",
    "com/sap/excel/utils/dataManipulationUtils",
  ],
  function (odataUtils, dataManipulationUtils) {
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

      _insertRowValues: function (fieldNames, cells, index, product) {
        fieldNames.forEach((fieldName) => {
          switch (fieldName) {
            case "ProductID": {
              let cellValue = `<b>${product.ProductID}</b><br>${product.ProductName}<br>`;
              cells[0].innerHTML = cellValue;

              if (index % 2 === 0) {
                cells.forEach(
                  (cell) => (cell.style.backgroundColor = "#f0f0f0")
                );
              }

              break;
            }
            
            case "SupplierID": {
              let cellValue = `${product.CompanyName}<br>`;
              cells[1].innerHTML = cellValue;
              break;
            }
            case "QuantityPerUnit": {
              let cellValue = `${product.QuantityPerUnit}<br>`;
              cells[2].innerHTML = cellValue;
              break;
            }
            case "UnitPrice": {
              let cellValue = `${product.UnitPrice}<br>`;
              cells[3].innerHTML = cellValue;
              break;
            }
            case "UnitsInStock": {
              let cellValue = `${product.UnitsInStock}<br>`;
              cells[4].innerHTML = cellValue;
              break;
            }
            case "Discontinued": {
              let cellValue = `${
                product.Discontinued ? "Discontinued" : "Available"
              }<br>`;
              if (product.Discontinued)
                cells[5].style.backgroundColor = "yellow";
              cells[5].innerHTML = cellValue;
              break;
            }
            case "ReorderLevel": {
                // For our current example, we will use this field to show an image since there is no field that contains the image...
                // In the real world, you will have a field that contains the image...
                let cellValue = this._getProuctImage();
                cells[6].innerHTML = cellValue;
                break;
            }
            default:
              break;
          }
        });
      },

      _getExpandedData: async function (controller) {
        let expand = "Supplier,Category";
        let data = await odataUtils.readExpandFromBackend(
          "Products",
          expand,
          controller._mainModel
        );

        if (data && data.results && data.results.length > 1) {
            data.results.forEach(result => {
                result.CategoryName = result.Category.CategoryName ? result.Category.CategoryName : result.CategoryID;
                result.CompanyName = result.Supplier.CompanyName ? result.Supplier.CompanyName : result.SupplierID;
            });
        }

        if (data.results) return data.results;

        return null;
      },

      _insertCategory: function (key, table) {
			let groupTbody = document.createElement("tbody");
			let groupTr = groupTbody.insertRow();
			groupTr.style.textAlign = "left";
			let groupCell = groupTr.insertCell();
			groupCell.colSpan = 7;
			groupCell.style.backgroundColor = "rgb(201, 238, 242)";
			groupCell.style.height = "30px";
			groupCell.style.verticalAlign = "middle";

			groupCell.innerHTML =
				`<b>Category: </b>${key}`;
			table.appendChild(groupTbody);
		},

        _getProuctImage: function() {
            return `<img src="https://via.placeholder.com/150/92c952" width="40" height="30" alt=""/><br>`;
        },

      createDOMTable: async function (controller) {
        let table = this._createTable("domTable");

        // Add the header section on top...
        let headerSection = this._createHeaderSection();
        table.appendChild(headerSection);

        let columnHeaders = [
          "Product ID",
          "Supplier Company",
          "Quantity Per Unit",
          "Unit Price",
          "Units In Stock",
          "Discontinued",
          "Product Image",
        ];
        let tr = this._addColumnHeaders(table, columnHeaders);

        // Retrieve data from backend...
        let products = await this._getExpandedData(controller);
        if (!products) return table;

          let fieldNames = Object.keys(products[0]);
          dataManipulationUtils.sortByCategory(products);
          let productsGroupedByCategory = dataManipulationUtils.groupByCategory(products);

          productsGroupedByCategory.forEach((productsInSingleCategory, key) => {

            this._insertCategory(key, table);

            // Add JSON data as rows to the table
            let tbody = document.createElement("tbody");

            productsInSingleCategory.forEach((product, index) => {
                tr = tbody.insertRow();

                let cells = [];
                columnHeaders.forEach((columnHeader, i) => {
                cells[i] = tr.insertCell();
                this._setCellStyle(cells[i]);
                });

                this._insertRowValues(fieldNames, cells, index, product);
            });

            table.appendChild(tbody);

          });

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
  }
);
