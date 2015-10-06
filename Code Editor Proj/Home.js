/* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
    See full license at the bottom of this file. */

(function () {
	"use strict";

	// The initialize function must be run each time a new page is loaded
	Office.initialize = function (reason) {
		$(document).ready(function () {
		    app.initialize();

		    // If not using Excel 2016, return
		    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
		        app.showNotification("Need Office 2016 or greater", "Sorry, this add-in only works with newer versions of Excel.");
		        return;
		    }

			$('#load-data-and-create-chart').click(loadDataAndCreateChart);
		});
	};

	// Load some sample data into the worksheet and then create a chart
	function loadDataAndCreateChart() {
		// Run a batch operation against the Excel object model
		Excel.run(function (ctx) {

			// Create a proxy object for the active worksheet
			var sheet = ctx.workbook.worksheets.getActiveWorksheet();

			//Queue commands to set the report title in the worksheet
			sheet.getRange("A1").values = "Quarterly Sales Report";
			sheet.getRange("A1").format.font.name = "Century";
			sheet.getRange("A1").format.font.size = 26;

			//Create an array containing sample data
			var values = [["Product", "Qtr1", "Qtr2", "Qtr3", "Qtr4"],
						  ["Frames", 5000, 7000, 6544, 4377],
						  ["Saddles", 400, 323, 276, 651],
						  ["Brake levers", 12000, 8766, 8456, 9812],
						  ["Chains", 1550, 1088, 692, 853],
						  ["Mirrors", 225, 600, 923, 544],
						  ["Spokes", 6005, 7634, 4589, 8765]];

			//Queue a command to write the sample data to the specified range
			//in the worksheet and bold the header row
			var range = sheet.getRange("A2:E8");
			range.values = values;
			sheet.getRange("A2:E2").format.font.bold = true;

			//Queue a command to add a new chart
			var chart = sheet.charts.add("ColumnClustered", range, "auto");

			//Queue commands to set the properties and format the chart
			chart.setPosition("G1", "L10");
			chart.title.text = "Quarterly sales chart";
			chart.legend.position = "right"
			chart.legend.format.fill.setSolidColor("white");
			chart.dataLabels.format.font.size = 15;
			chart.dataLabels.format.font.color = "black";
			var points = chart.series.getItemAt(0).points;
			points.getItemAt(0).format.fill.setSolidColor("pink");
			points.getItemAt(1).format.fill.setSolidColor('indigo');

			//Run the queued-up commands, and return a promise to indicate task completion
			return ctx.sync();
		})
		  .then(function () {
			  app.showNotification("Success");
			  console.log("Success!");
		  })
		.catch(function (error) {
		    // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
		    app.showNotification("Error: " + error);
		    console.log("Error: " + error);
		    if (error instanceof OfficeExtension.Error) {
		        console.log("Debug info: " + JSON.stringify(error.debugInfo));
		    }
		});
	}
})();

/*
Excel-Add-in-JS-QuarterlySalesReportManifest, https://github.com/OfficeDev/Excel-Add-in-JS-QuarterlySalesReport

Copyright (c) Microsoft Corporation

All rights reserved. 

MIT License:

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and 
associated documentation files (the "Software"), to deal in the Software without restriction, including 
without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell 
copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the 
following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial 
portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT 
LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT 
SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN 
ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE 
USE OR OTHER DEALINGS IN THE SOFTWARE. 
*/