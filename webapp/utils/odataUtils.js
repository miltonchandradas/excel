sap.ui.define([], function () {
	"use strict";

	return {

		readFromBackend: function (entity, mainModel) {

			return new Promise(function (resolve, reject) {
				mainModel.read("/" + entity, {
					success: function (data) {
						resolve(data);
					},
					error: function (err) {
						reject(err);
					}
				});
			});
        }, 
        
        readExpandFromBackend: function (entity, expand, mainModel) {

			return new Promise(function (resolve, reject) {
				mainModel.read("/" + entity, {
					urlParameters: {
						"$expand": expand
					},
					success: function (data) {
						resolve(data);
					},
					error: function (err) {
						reject(err);
					}
				});
			});
		}
	};

});