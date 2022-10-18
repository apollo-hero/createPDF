// 'use strict';

var PDFDocument = require('pdfkit');
var SVGtoPDF = require('svg-to-pdfkit');
var Excel = require('exceljs');

var fs = require('fs');
var readExcel = require('./readExcel.js');
const { createConverter } = require('convert-svg-to-png');
var util = require('util');
var readdir = util.promisify(fs.readdir);

var TOP_Y = cmToPt(2);       // top of workarea
// var BOTTOM_Y = cmToPt(28.5); // bottom of workarea  21cm * 29.7cm 21*15
var BOTTOM_Y = cmToPt(13.8);
// var CENTER_X = cmToPt(10.5);
var CENTER_X = cmToPt(10.5);
// var DRAW_WIDTH = cmToPt(18.6);
var DRAW_WIDTH = cmToPt(18.6);
// var PAGE_WIDTH = cmToPt(21);
var PAGE_WIDTH = cmToPt(21);
var MARGIN_X = cmToPt(1);
var MARGIN_Y = cmToPt(1);
var FONTSIZE = 12;

var cPage = 1;              // current page
var cur_X = CENTER_X;       // current X
var cur_Y = TOP_Y;          // current Y 
var newpage = false;
var n_lines = 0;
var page_number = 0;
var cur_X1;
var cur_Y1;
var page_numbers = [];

// Preprocess the read data
function preProcess(readdata) {
	var data = [];
	var tmpID = "";

	readdata.forEach(function (record) {
		let svgdata = {};
		svgdata.qty = record.qty;
		svgdata.fName = record.fname;
		svgdata.isPng = record.isPNG;
		svgdata.isText = record.isText;
		svgdata.pngPath = record.pngPath;
		svgdata.width = record.width;
		svgdata.height = record.height;
		svgdata.data = record.data;
		if (tmpID != record.id) {
			var new_record = {};
			new_record.id = record.id;
			// new_record.svg = [{qty: record.qty, fName: record.fName, isPng: record.isPng, width: record.width, height: record.height, data: record.data}];
			new_record.svg = [svgdata];
			data.push(new_record);
			tmpID = record.id;

		}
		else {

			var record = data[data.length - 1];
			record.svg.push(svgdata);
		}
	})

	return data;
}

var lines = []
var lineData = [];

// Make line data for drawing
function makeLineData(id, nodes, space) {
	if (nodes.length == 0) {
		if (lineData.length != 0)
			lines.push(lineData);
		return;
	}
	var node = nodes[0];
	if (isDrawable(node, space)) { 	// If all images are drawable in extra space
		// Append this item to the line data
		node.id = id;
		lineData.push(node);

		// Remove this node from nodes
		nodes.shift();

		// Check if extra space
		var restWidth = calRest(node, space);
		if (restWidth > 0) {			// space is 0px over.
			makeLineData(id, nodes, restWidth);
		}
		else {
			lines.push(lineData);
			lineData = [];
			makeLineData(id, nodes, DRAW_WIDTH);
		}
	}
	else {
		// Calculate qty that can be drawn   ex: 8
		var qty = calQty(node, space);
		if (qty == 0) {
			lines.push(lineData);
			lineData = [];

			makeLineData(id, nodes, DRAW_WIDTH);
		}
		else {
			// Add this node to lines
			item = { id: id, qty: qty, fName: node.fName, isPng: node.isPng, pngPath: node.pngPath, width: node.width, height: node.height, data: node.data };
			lineData.push(item);
			lines.push(lineData);
			lineData = [];

			// Decrease node's qty   ex: 11 -8 = 3
			nodes[0].qty -= qty;

			// Call me with changed nodes
			makeLineData(id, nodes, DRAW_WIDTH);
		}

	}
}



function createPDF(excelFile, imgPath, output, csvFile) {

	// Write data from excel to PDF file
	function writeData(readdata) {

		// Init doc

		var doc = new PDFDocument({
			autoFirstPage: false,
			size: 'A5',             // 21cm * 29.7cm
			layout: 'landscape',     //'landscape'
			bufferPages: true
		});


		console.log("Converting...");

		// Preprocess the readdata so that it is sorted according to ID
		var firstData = preProcess(readdata, imgPath);
		// console.log(JSON.stringify(firstData, null, 2));
		// make lines data 
		for (var i = 0; i < firstData.length; i++) {
			// For test purpose, prepare line data for the first ID
			var nodes = firstData[i].svg;
			var id = firstData[i].id;
			makeLineData(id, nodes, DRAW_WIDTH);
			lineData = [];

		}
		// draw items
		cur_X = CENTER_X;
		cur_Y = TOP_Y;
		var next_id = "";
		n_lines = lines.length;

		cur_X1 = cur_X;
		cur_Y1 = cur_Y;

		for (i = 0; i < lines.length; i++) {

			items = lines[i];
			item_id = items[0].id;

			next_items = lines[i + 1];
			if (next_items != null)
				next_id = next_items[0].id;
			else
				next_id = "";
			pageCheck(items, cur_X, cur_Y, doc);
			drawItem(doc, cur_X, cur_Y, items, imgPath, output, i + 1);
			if (item_id != next_id) {
				//pageCheck("", cur_X, cur_Y, doc);
				writeID(doc, item_id, cur_X, cur_Y);
				page_numbers[item_id] = page_number;
				page_number = 1;
				doc.addPage();
				//  red line of workarea
				doc.rect(cmToPt(1), cmToPt(1), doc.page.width - cmToPt(2), doc.page.height - cmToPt(2))
					.fillOpacity(1.0)
					.fillAndStroke("white", "red");

				// pagenumber botton center of page
				doc.fillColor("black")
					.fontSize(14)
					.font('Times-Roman')
					.text(cPage, 0, cmToPt(14), {
						width: doc.page.width,
						height: 14,
						align: 'center',

					});

				// add id
				doc.fillColor("black")
					.fontSize(FONTSIZE)
					.font('Times-Roman')
					.text('#' + next_id + "-" + page_number, 0, cmToPt(0.5), {
						width: doc.page.width,
						height: FONTSIZE + 0.1,
						align: 'center'
					});

				largeHeight1 = FONTSIZE + 0.1 + MARGIN_Y;
				if (largeHeight1 >= cmToPt(26.5))	// over top 2cm, bottom 28.7
				{
					cur_Y = (cmToPt(29.7) - largeHeight) / 2;
				}
				else
					cur_Y = TOP_Y;
				cPage++;
				newpage = true;
			}
		}

		// Write pdf doc.

		// doc.pipe(fs.createWriteStream(output))
		// 	.on('finish', function () {
		// 		console.log('Done.');
		// 	});
		// doc.end();

		//deleteFolderRecursive(imgPath + '/temp');
		//deleteFolderRecursive(imgPath + '/strings');

		cPage = 1;              // current page
		cur_X = CENTER_X;       // current X
		cur_Y = TOP_Y;          // current Y 
		newpage = false;
		n_lines = 0;
		page_number = 0;
		lines = []
		lineData = [];

		writeData1(readdata);

	}

	function writeData1(readdata) {

		// Init doc

		var newDoc = new PDFDocument({
			autoFirstPage: false,
			size: 'A5',             // 15cm * 21cm
			layout: 'landscape',     //'landscape'
			bufferPages: true
		});


		// Preprocess the readdata so that it is sorted according to ID
		var firstData = preProcess(readdata, imgPath);
		// console.log(JSON.stringify(firstData, null, 2));
		// make lines data 
		for (var i = 0; i < firstData.length; i++) {
			// For test purpose, prepare line data for the first ID
			var nodes = firstData[i].svg;
			var id = firstData[i].id;
			makeLineData(id, nodes, DRAW_WIDTH);
			lineData = [];

		}
		// draw items
		cur_X = CENTER_X;
		cur_Y = TOP_Y;
		var next_id = "";
		n_lines = lines.length;

		cur_X1 = cur_X;
		cur_Y1 = cur_Y;

		for (i = 0; i < lines.length; i++) {

			items = lines[i];
			item_id = items[0].id;

			next_items = lines[i + 1];
			if (next_items != null)
				next_id = next_items[0].id;
			else
				next_id = "";
			pageCheck(items, cur_X, cur_Y, newDoc);
			drawItem(newDoc, cur_X, cur_Y, items, imgPath, output, i + 1);
			if (item_id != next_id) {
				//pageCheck("", cur_X, cur_Y, doc);
				writeID(newDoc, item_id, cur_X, cur_Y);
				if (next_id != "") {
					page_number = 1;
					newDoc.addPage();
					//  red line of workarea
					newDoc.rect(cmToPt(1), cmToPt(1), newDoc.page.width - cmToPt(2), newDoc.page.height - cmToPt(2))
						.fillOpacity(1.0)
						.fillAndStroke("white", "red");

					// pagenumber botton center of page
					newDoc.fillColor("black")
						.fontSize(14)
						.font('Times-Roman')
						.text(cPage, 0, cmToPt(14), {
							width: newDoc.page.width,
							height: 14,
							align: 'center',

						});

					// add id
					newDoc.fillColor("black")
						.fontSize(FONTSIZE)
						.font('Times-Roman')
						.text('#' + next_id + " - [" + page_number + "/" + page_numbers[next_id] + "]", 0, cmToPt(0.5), {
							width: newDoc.page.width,
							height: FONTSIZE + 0.1,
							align: 'center'
						});

					largeHeight1 = FONTSIZE + 0.1 + MARGIN_Y;
					if (largeHeight1 >= cmToPt(26.5))	// over top 2cm, bottom 28.7
					{
						cur_Y = (cmToPt(29.7) - largeHeight) / 2;
					}
					else
						cur_Y = TOP_Y;
					cPage++;
					newpage = true;
				}
			}
		}

		// Write pdf doc.

		newDoc.pipe(fs.createWriteStream(output))
			.on('finish', function () {
				console.log('Done.');
			});
		newDoc.end();

		//deleteFolderRecursive(imgPath + '/temp');
		//deleteFolderRecursive(imgPath + '/strings');

	}
	readExcel(excelFile, imgPath, writeData, csvFile);
}

function isDrawable(node, space) {
	var lineWidth = node.width * node.qty + MARGIN_X * (node.qty - 1);
	if (lineWidth <= space)
		return true;
	else
		return false;
}

function calRest(node, space) {
	var lineWidth = node.width * node.qty + MARGIN_X * (node.qty - 1);
	var rest = space - lineWidth - cmToPt(2);		//insert item edge = 2cm

	return rest;
}

function calQty(node, space) {	//DRAW_WIDTH = 481.8905px	497.647px   
	var lineWidth = node.width;
	num = 0;

	while (lineWidth <= space) {
		lineWidth += node.width + MARGIN_X;
		num++;
	}
	return num;
}

//  page add check & Red frame draw & page number add
function pageCheck(items, x, y, doc) {

	var largeHeight = 0;
	if (items != "") {	//item draw space

		items.forEach(function (item) {
			if (largeHeight < item.height)
				largeHeight = item.height;
		})
	}
	else {					// id draw space
		largeHeight = FONTSIZE + 0.1 + MARGIN_Y;
		// y = y + largeHeight;
	}


	if (y + largeHeight >= BOTTOM_Y || cPage == 1) {     // if redline over when put current item , addpage
		page_number++;
		doc.addPage();
		//  red line of workarea
		doc.rect(cmToPt(1), cmToPt(1), doc.page.width - cmToPt(2), doc.page.height - cmToPt(2))
			.fillOpacity(1.0)
			.fillAndStroke("white", "red");

		// pagenumber botton center of page
		doc.fillColor("black")
			.fontSize(14)
			.font('Times-Roman')
			.text(cPage, 0, cmToPt(14), {
				width: doc.page.width,
				height: 14,
				align: 'center',

			});

		// add 
		var s;
		if (page_numbers[items[0].id]) {
			s = "/" + page_numbers[items[0].id] + "]";
		} else {
			s = "";
		}
		doc.fillColor("black")
			.fontSize(FONTSIZE)
			.font('Times-Roman')
			.text('#' + items[0].id + " - [" + page_number + s, 0, cmToPt(0.5), {
				width: doc.page.width,
				height: FONTSIZE + 0.1,
				align: 'center'
			});

		if (largeHeight >= cmToPt(26.5))	// over top 2cm, bottom 28.7
		{
			cur_Y = (cmToPt(29.7) - largeHeight) / 2;
		}
		else
			cur_Y = TOP_Y;
		cPage++;
		newpage = true;
	}

	if (items == "")
		cur_Y = cur_Y - cmToPt(0.5);
}

function positionX(items, x) {
	var width = 0;
	var col = 0;

	items.forEach(function (item) {
		for (var i = 0; i < item.qty; i++) {
			width += item.width + MARGIN_X;
		}
	})

	return (PAGE_WIDTH - width + MARGIN_X) / 2;
}

function positionY(items, y) {
	var largeHeight = 0;

	items.forEach(function (item) {
		if (largeHeight < item.height)
			largeHeight = item.height;
	})
	// console.log(y.toFixed(4), '\t', largeHeight.toFixed(4), '\t=', (y + largeHeight + MARGIN_Y).toFixed(4));
	return y + largeHeight + MARGIN_Y;
}

// png or svg draw
function drawItem(doc, x, y, records, svgPath, output, line_num) {
	x = positionX(records, x);
	for (var i = 0; i < records.length; i++) {
		for (j = 0; j < records[i].qty; j++) {

			if (records[i].isPng == true)
				pngTopdf(doc, records[i].pngPath, x, y, records[i].width, records[i].height, output, line_num);
			else
				svgTopdf(doc, records[i].data, x, y, records[i].width, records[i].height, svgPath, output, line_num);
			x = x + records[i].width + MARGIN_X;
		}

	}

	cur_X = CENTER_X;
	cur_Y = positionY(records, y);
}


// png file to pdf
function pngTopdf(doc, pngPath, x, y, w, h, output, line_num) {

	doc.image(pngPath, x, y, {
		fit: [w, h],
		align: 'center',
		valign: 'center'
	});
}

// svg file to pdf 
function svgTopdf(doc, record, x, y, w, h, svgPath, output, line_num, ratio = "none") {
	SVGtoPDF(doc, record, x, y, {
		preserveAspectRatio: ratio, //"xMinYMin meet",
		width: w,
		height: h,
		align: 'center',
		valign: 'center'
	});
}

// id & dash line  draw
function writeID(doc, id, x, y) {
	doc.fill("red")
		.moveTo(-5, y)
		.lineTo(doc.page.width + 5, y)
		.dash(11, { space: 11 })
		.stroke();


	//dash line draw
	// y = y + FONTSIZE + MARGIN_Y;
	// doc.fillColor("black")
	// .fontSize(FONTSIZE)
	// .font('Times-Roman')
	// .text('#' + id, 5, cmToPt(14), {
	// 	width: doc.page.width,
	// 	height: FONTSIZE + 0.1,
	// 	align: 'right'
	// });
	cur_Y = y + MARGIN_Y;
}

//scale function
function cmToPt(cm) {
	return cm * 28.3465;
}

function ptToCM(pt) {
	return pt / 28.3465;
}

function deleteFolderRecursive(path) {
	var files = [];
	if (fs.existsSync(path)) {
		files = fs.readdirSync(path);
		files.forEach(function (file, index) {
			var curPath = path + "/" + file;
			if (fs.lstatSync(curPath).isDirectory()) { // recurse
				deleteFolderRecursive(curPath);
			} else { // delete file
				fs.unlinkSync(curPath);
			}
		});
		fs.rmdirSync(path);
	}

};

async function convertSvgFiles(excelFile, svgPath, outPath, csvFile) {

	const converter = createConverter();
	var svgFilePath = "";
	try {
		const filePaths = await readdir(svgPath);
		if (!fs.existsSync(svgPath + '/temp'))
			fs.mkdirSync(svgPath + '/temp');
		for (const filePath of filePaths) {
			if (filePath.includes('.svg')) {
				svgFilePath = filePath;
				fileName = filePath.substring(0, filePath.length - 4);
				await converter.convertFile(svgPath + '/' + filePath, { outputFilePath: svgPath + '/temp/' + fileName + '.png' });
			}
			// throw filePath;
		}
	}
	catch (err) {
		console.log(svgFilePath, "converting error", err);
	}
	finally {
		await converter.destroy();
		createPDF(excelFile, svgPath, outPath, csvFile);
	}
}

// Call test function
// convertSvgFiles('./Input/input.xlsx', './SVGs', './Output/input.pdf', './Output/input.csv');
// createPDF('Input/A4Input.xlsx', './SVGs', 'Output/A4Input.pdf', 'Output/A4Input.csv');
createPDF('Input/input.xlsx', './SVGs', 'Output/output-new.pdf', 'Output/output.csv');



