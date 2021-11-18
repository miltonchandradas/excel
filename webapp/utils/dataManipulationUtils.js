sap.ui.define([], function () {
	"use strict";

	return {

        sortByCategory: function (products) {
            products.sort((a, b) => (a.CategoryName > b.CategoryName) ? 1 : -1);
		},

		groupByCategory: function (products) {
            let productsGroupedByCategory = products.reduce((accumulator, current) => {
				if (!accumulator.has(current.CategoryName))
					accumulator.set(current.CategoryName, []);

				accumulator.get(current.CategoryName).push(current);
				return accumulator;
			}, new Map());

			return productsGroupedByCategory;
        }
	};

});