var fs = require('fs');
var Excel = require('exceljs');
var sanitize = require('sanitize-filename');
var DOMParser = require('xmldom').DOMParser;
const TextToSVG = require('text-to-svg');
const opentype = require('opentype.js');

const { convertFile } = require('convert-svg-to-png');

var fonts = [''];

function loadFonts(fontsPath) {
	if (fs.existsSync(fontsPath)) {
		var files = fs.readdirSync(fontsPath, { withFileTypes: true });

		files.forEach(file => {
			if (file.isFile())
				fonts.push(file.name);
		})
	}

}


function getImgData(fname, svgFolder, isString) {
	var parser = new DOMParser();
	var imgFile = svgFolder + '/' + ((isString) ? 'strings/' : '') + fname + '.svg';
	if (!fs.existsSync(imgFile)) {
		return false;
	}

	var contents = fs.readFileSync(imgFile, 'utf8');
	var doc = parser.parseFromString(contents, "image/svg+xml");

	var svg = doc.getElementsByTagName("svg");
	var img = doc.getElementsByTagName("image");

	var imgData = {};
	imgData.width = parseFloat(svg[0].getAttribute('width'));
	imgData.height = parseFloat(svg[0].getAttribute('height'));

	if (img.toString() != "") {
		return false;
	}
	else {

		imgData.isPNG = false;
		imgData.pngPath = '';
		imgData.svgData = contents;
	}

	return imgData;
}

function loadFont(font, optionsByFont) {

	if (font >= fonts.length) {
		return TextToSVG.loadSync();
	} else {
		return TextToSVG.loadSync('./fonts/' + optionsByFont[font].font);
	}
}

function loadFont2(font) { //

	if (font >= fonts.length) {
		return opentype.loadSync('./fonts/' + fonts[1]);
	} else {
		return opentype.loadSync('./fonts/' + fonts[font]);
	}
}


function createSVGFromText(svgPath, title, color, size, font) {
    
	// font3 section 
	var font3_1 = title.slice(-1).match(/[b|d|h|k|l]/)?0.42:0;
	var font3_2 = title.slice(-2,-1).match(/[b|d|h|k|l]/)?0.33:0;
	var font3_3 = title.slice(-1).match(/[L|Z]/)?0.2:0;
	var font3_4 = (title.slice(-3,-2).match(/[b|d|h|k|l]/) || title.slice(-2,-1).match(/[L|Z]/) || title.slice(-1).match(/[g|j|q|t|y|V|T]/))?0.11:0;
	var font3_5 = (title.slice(-4,-3).match(/[b|d|h|k|l]/) || title.slice(-1).match(/[S]/))?0.05:0;
	var width3_1 = title.slice(0,1).match(/[f|g|j|q|y|z|J]/)?0.1:0;
	var font3=0;
	if(font3_1){font3=font3_1}else if(font3_2){font3=font3_2}else if(font3_3){font3=font3_3}else if(font3_4){font3=font3_4}else{font3=font3_5};

	//font4 section
	var hei4_down1 = title.match(/[|g|]/i)?2.1:1;
	var hei4_down2 = title.match(/[|j|p|s|y|]/i)?1.5:1;
	var hei4_down = (hei4_down1==2.1)?hei4_down1:hei4_down2;
	var hei4_up1 = title.match(/[|t|]/i)?0.5:0;
	var hei4_up2 = title.match(/[|l|]/i)?0.2:0;
	var hei4_up = (hei4_up1)?hei4_up1:hei4_up2;
	var hei4 = hei4_up + hei4_down;

	var wid4_1 = title.slice(-1).match(/[|t|]/i)?59:0;
	var wid4_2 = title.slice(-2,-1).match(/[|t|]/i)?20:0;
	var wid4_3 = title.slice(-1).match(/[|g|]/i)?24:0;
	var wid4_4 = title.slice(-3,-2).match(/[|t|]/i)?10:0;
	var wid4_5 = title.slice(-2,-1).match(/[|g|]/i)?10:0;
	var wid4_6 = title.slice(-3,-2).match(/[|g|]/i)?5:0;
	var wid4_s = title.slice(0,1).match(/[|s|]/i)?28:0; // only add width
	var wid4_t = title.slice(0,1).match(/[t]/i)?7:0;   // only add width
    var wid4_j = title.slice(0,1).match(/[j]/i)?11:0;   // only add width
	var wid4_only = wid4_s + wid4_t + wid4_j;
	var wid4=0;
	if(wid4_1){wid4=wid4_1}else if(wid4_2){wid4 = wid4_2}else if(wid4_3){wid4=wid4_3}else if(wid4_4){wid4=wid4_4}else if(wid4_5){wid4=wid4_5}else{wid4=wid4_6};

	//font6 section
	var wid6_1 = title.slice(-1).match(/[|T|K|F]/)?20:0;
	var wid6_2 = (title.slice(-1).match(/[|Z|I|l|E]/) || title.slice(-2,-1).match(/[|T|]/))?10:0;
	var wid6_3 = title.slice(-1).match(/[|R|L|C]/)?3:0;
	var wid6=0;
	if(wid6_1){wid6=wid6_1}else if(wid6_2){wid6=wid6_2}else{wid6=wid6_3};

	//font7 section
	var hei7=0;
	var hei7_f = title.match(/[|f|Y|G]/)?1.8:0;
	var hei7_uppercase = title.match(/[A-F|H-X|Z]/)?1.6:0;
	var hei7_down=title.match(/[|g|j|p|q|y|z|]/)?0.8:0.25;
	var hei7_up = title.match(/[|b|d|h|k|l|]/)?0.8:0.25;
	hei7_up = hei7_up + hei7_down;
	if(hei7_f){hei7=hei7_f}else if(hei7_uppercase){hei7=hei7_uppercase + hei7_down/4}else{hei7=hei7_up};
	
	var wid7_1 = title.slice(-1).match(/[|V|W]/)?33.5:0;
	var wid7_2 = title.slice(-1).match(/[|N|]/)?28.5:0;
	var wid7_3 = title.slice(-1).match(/[|X|]/)?23:0;
	var wid7_4 = title.slice(-1).match(/[|M||P|]/)?18:0;
	var wid7_5 = title.slice(-1).match(/[B]/)?16:0;
	var wid7_6 = title.slice(-1).match(/[|b|f|h|k|l|T|Y|L|J|I|H|F|D|A|]/)?15.3:0;
	var wid7_7 = title.slice(-1).match(/[|S|R|]/)?12:0;
	var wid7_8 = (title.slice(-2,-1).match(/[|b|f|h|k|l|]/) || title.slice(-1).match(/[|U|O|K|G|C|]/))?7:0;
	var wid7_9 = title.slice(-1).match(/[|E|Q|Z]/)?5:0;
	var wid7=0;
	if(wid7_1){wid7=wid7_1}else if(wid7_2){wid7=wid7_2}else if(wid7_3){wid7=wid7_3}else if(wid7_4){wid7=wid7_4}else if(wid7_5){wid7=wid7_5}else if(wid7_6){wid7=wid7_6}else if(wid7_7){wid7=wid7_7}else if(wid7_8){wid7=wid7_8}else{wid7=wid7_9};

	//font8 section
	var hei8 = title.match(/[|p|g|]/)?0.08:0;
	
	var wid8_1 = title.slice(-1).match(/[|K|]/)?15:0;
	var wid8_2 = title.slice(-1).match(/[|Z|W|V|T|H|C|B|R]/)?7:0;
	var wid8_3 = title.slice(-1).match(/[|P|O|L|I|E]/)?4:0
	var wid8=0;
	if(wid8_1){wid8=wid8_1}else if(wid8_2){wid8=wid8_2}else{wid8=wid8_3};

	//font9 section

	var wid9 = title.slice(-1).match(/[|V|T|L]/)?3:0; 

	var optionsByFont = [
		{ font: '01 Kingthings.ttf',linesize:0.1, factorh: 1, factorw: 1, x: 0, y: 0, anchor: 'top', translatew: 0, divided: function (w) { return 0 }, dividedt: 1 },
		{ font: '01 Kingthings.ttf',linesize:0.1, factorh: 1, factorw: 1, x: 0, y: 0, anchor: 'top', translatew: 0, divided: function (w) { return 0 }, dividedt: 1 },
		{ font: '02 OldLondon.ttf',linesize:0, factorh: 1.6, factorw: 1, x: 0, y: 0, anchor: 'top', translatew: 0, divided: function (w) { return 3 }, dividedt: 1 },
		{ font: '03 SCRIPTIN.ttf',linesize:0.1, factorh: 0.8, factorw: 1, x: 0, y: title.match(/[b|d|h|k|l|t|p|q|g|y|z]/) ? 30:30, anchor: 'middle', translatew:9+font3 * 100, divided:function (w) { return 9+font3 * 100 + width3_1 * 60 }, dividedt: 1 },
		{ font: '04 RatInfestedMailbox.ttf',linesize:0.1, factorh: 1 * hei4, factorw: 1, x: 0, y: title.match(/[|T|]/i) ? 35 : 10 + hei4_up2 * 40, anchor: 'top', translatew:9 + wid4 , divided: function (w) { return 15 + wid4 +wid4_only }, dividedt: 1 },
		{ font: '05 BacanaRegular.ttf',linesize:0.5, factorh: 2.25, factorw: 1, x: 0, y: 20, anchor: 'top', translatew: 5, divided: function (w) { return 8 }, dividedt: 1 },
		{ font: '06 kevinwildfont.ttf',linesize:0, factorh: 1.2, factorw: 1, x: 0, y: 5, anchor: 'top', translatew: 5 + wid6, divided: function (w) { return 7 + wid6 }, dividedt: 1 },
		{ font: '07 porcelai-webfont.woff',linesize:0, factorh: 0.5 * hei7 + (title.match(/[|B|E|]/)?0.04:0), factorw: 1, x:0, y:title.match(/[|B|E|]/)?-2.5:-5, anchor: 'top', translatew:2 + wid7, divided: function (w) { return 3 + wid7}, dividedt: 1 },
		{ font: '08 drjekyll-webfont.woff',linesize:0, factorh: 0.85 + hei8, factorw: 1, x: 0, y: -3, anchor: 'top', translatew: 5 + wid8, divided: function (w) { return 8 + wid8 }, dividedt: 1 },
		{ font: '09 Autery.ttf', factorh: 1,linesize:0, factorw: 1, x: 0, y: 1, anchor: 'top', translatew:3+wid9, divided: function (w) { return 5 + wid9 }, dividedt: 1 },
		{ font: '10 禹卫书法行书繁体.ttf',linesize:0, factorh: 1, factorw: 1, x: 0, y: 0, anchor: 'top', translatew: 0, divided: function (w) { return 0 }, dividedt: 1 }
	]
	let fileName = sanitize(title);
	const textToSVG = loadFont(font, optionsByFont);
	const attributes = { fill: color, transform: "" ,stroke: color ,s:"ss"};
	const options = { x: optionsByFont[font].x, y: optionsByFont[font].y, anchor: optionsByFont[font].anchor, fontSize: 36, attributes: attributes };
	const metrics = textToSVG.getMetrics(title, options);
	attributes.transform = `scale(-1,1) translate(${-metrics.width / optionsByFont[font].dividedt - optionsByFont[font].translatew},0)`
	var svg = textToSVG.getSVG(title, options);
	svg = svg.replace(`height="${metrics.height}"`, `height="${metrics.height * optionsByFont[font].factorh}"`);
	svg = svg.replace(`width="${metrics.width}"`, `width="${metrics.width * optionsByFont[font].factorw + optionsByFont[font].divided(metrics.width)}"`);
	svg = svg.replace(`s="ss"`,`stroke-width="${optionsByFont[font].linesize}"`);
	fs.writeFileSync(svgPath + '/strings/' + fileName + '.svg', svg);
	return fileName;
}

function createSVGFromText_old(svgPath, title, color, size, font) {
	let fileName = sanitize(title);
	const textToSVG = loadFont(font);
	const fontFace = loadFont2(font);
	let attributes = { fill: color, transform: "" };
	const options = { x: 0, y: 0, fontSize: 144, anchor: 'top', attributes: attributes };
	const metrics = textToSVG.getMetrics(title, options);
	const path = fontFace.getPath(title, 0, -fontFace.getPath(title, 0, 0).getBoundingBox().y1); //


	// var svg = textToSVG.getSVG(title, options);

	svg = `<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="${metrics.width}" height="${metrics.height}">`;//
	svg += path.toSVG(0);//
	svg = svg.replace('<path ', '<path transform = "scale(-1,1) translate(' + (-metrics.width) + ', 0)" fill="' + color + '" ');//
	svg += '</svg>'//

	// svg = svg.replace('<svg ',`<svg viewBox="0 0 ${metrics.width} ${metrics.height}"`)
	fs.writeFileSync(svgPath + '/strings/' + fileName + '.svg', svg);
	return fileName;
}

// getting 'ID', 'SVG' and 'QTY' from xlsx file.
var readExcel = function (filename, svgFolder, callbackfunc, csvFile) {

	loadFonts('fonts');

	// open excel file
	var workbook = new Excel.Workbook();
	workbook.xlsx.readFile(filename)
		.then(function () {
			var worksheet = workbook.getWorksheet(1);

			var index = 2;
			var data = [];
			noImgs = [];

			do {
				// read one row cells information
				var tempID = worksheet.getCell('A' + index).value;

				var tempStrings = worksheet.getCell('B' + index).value;
				var tempFont = worksheet.getCell('C' + index).value;
				var tempColor = worksheet.getCell('D' + index).value;
				var tempLengths = worksheet.getCell('E' + index).value;
				var tempPack = worksheet.getCell('F' + index).value;
				var tempFName = worksheet.getCell('G' + index).value;
				var tempQty = worksheet.getCell('H' + index).value;

				var record = {};
				record.id = tempID;
				record.strings = tempStrings;
				record.font = tempFont;
				record.color = tempColor;
				record.lengths = tempLengths;
				record.pack = tempPack;
				record.fname = tempFName;
				record.qty = (tempStrings !== null) ? tempQty * tempPack : tempQty;
				record.isText = false;

				if (!fs.existsSync(svgFolder + '/' + 'strings'))
					fs.mkdirSync(svgFolder + '/' + 'strings');
				if (tempStrings != null) {
					tempFName = createSVGFromText(svgFolder, tempStrings, tempColor, tempLengths, tempFont);
					record.isText = true;
					// tempFName = tempStrings;
				}

				// Check if SVG file is in svgFolder only if row (in Excel) has svg file
				var imgData = getImgData(tempFName, svgFolder, tempStrings != null);
				if (imgData == false && tempStrings == null) {
					// If no SVG file is found, save into CSV file.
					var recordT = [];

					recordT.push(tempID);
					recordT.push(tempFName);
					recordT.push(tempQty);
					noImgs.push(recordT);

					record.imgData = null;

				}
				else {

					// TODO width can't be more than 527px because it goes outside the page
					if (imgData.width > 527) {
						imgData.height = imgData.height * 527 / imgData.width;
						imgData.width = 527;
					}

					// Size of the string in inches 
					if (tempStrings != null) {
						imgData.height = imgData.height * (tempLengths * 72) / imgData.width;
						imgData.width = tempLengths * 72;
					}

					// Read SVG file
					record.isPNG = imgData.isPNG;
					record.width = imgData.width;
					record.height = imgData.height;
					record.data = imgData.svgData;
					record.pngPath = imgData.pngPath;
					data.push(record);
				}

				index = index + 1;
				// check else svg infomation in excel.
				tempID = worksheet.getCell('A' + index).value;
				record.nextID = tempID;
			} while (tempID != null);

			// Create CSV file containing information of SVG file that is not in folder.
			var csvbook = new Excel.Workbook();
			var sheet = csvbook.addWorksheet('My Sheet', { pageSetup: { paperSize: 9, orientation: 'landscape' } });
			sheet.columns = [
				{ header: 'ID', key: 'id', width: 30 },
				{ header: 'SVG', key: 'svg', width: 30 },
				{ header: 'QTY', key: 'qty', width: 10 }
			];

			// Add row with information of SVG that is not in folder.
			for (var i = 0; i < noImgs.length; i++) {
				sheet.addRow([noImgs[i][0], noImgs[i][1], noImgs[i][2]]).commit();
			}

			var noImgPath = './' + csvFile;
			try {
				csvbook.csv.writeFile(noImgPath);
			} catch (err) {
				console.log("already file exist.");
			}

			callbackfunc(data);
		});

}

// readExcel('test.xlsx', './SVGs', 'Output/test.pdf', 'Output/test.csv');

module.exports = readExcel;
