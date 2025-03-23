function Html2Docxml(doc, settings) {

	this.parse = function (htmlText, callback) {

		var parser = new DOMParser();
		var parsedDom = parser.parseFromString(htmlText, "text/html");

		return this.convert(parsedDom, callback);
	}

	this.isProcessImg = function () {
		return settings?.process_img ? true : false;
	}

	this.convert = function (parsedDom, callback) {

		var imgsDom = parsedDom.getElementsByTagName("IMG");

		// console.log("imgsDom:", imgsDom);

		if (imgsDom.length > 0 /*&& this.isProcessImg()*/) {
			return this.generateDocxWithImgs(parsedDom, imgsDom, callback);
		} else {

			return this.generateDocx(parsedDom, callback, false);
		}
	}

	this.recursImg = function (allImg, i, parsedDom, callback) {

		// console.log("recursImg:", allImg[i], i, allImg[i].style.height, allImg[i].style.width)


		// return;

		if (allImg.length == 0) {
			return this.generateDocx(parsedDom, callback, false);
		}

		var that = this;
		var url = allImg[i].src;
		// var res = null;

		let img_height = allImg[i].style.height.replace('px', '');
		let img_width = allImg[i].style.width.replace('px', '');


		// createCanvas 
		var img = new Image();
		img.crossOrigin = 'anonymous';
		img.src = url;

		img.onload = function () {

			var canvas = document.createElement('CANVAS');
			ctx = canvas.getContext('2d');
			canvas.height = img.height;
			canvas.width = img.width;
			ctx.drawImage(img, 0, 0);

			var dataURL = canvas.toDataURL('image/png');

			allImg[i].src = dataURL;
			allImg[i].width = img_width;
			allImg[i].height = img_height;


			canvas = null
			i++;

			if (i < allImg.length) {

				return that.recursImg(allImg, i, parsedDom, callback);
			}
			else {

				return that.generateDocx(parsedDom, callback, true);
			}
		};
	}


	this.generateDocxWithImgs = function (parsedDom, imgsDom, callback) {
		var allImg = [];

		var i = 0;
		for (var j = 0; j < imgsDom.length; j++) {

			var url = imgsDom[j].src;

			if ((url.startsWith('http:') || url.startsWith('https:')) && url.includes(window.location.hostname) && url.match(/.(jpg|jpeg|png|gif)/gi))
				allImg.push(imgsDom[j]);
		}

		// console.log("allImg:", allImg.length)


		// console.log("allImg.length:", allImg.length)

		return this.recursImg(allImg, i, parsedDom, callback);


	}

	this.generateDocx = function (parsedDom, callback, is_img) {
		// console.log("generateDocx,parsedDom:", parsedDom)


		var nodeParent = parsedDom.getElementsByTagName("body");
		var xmlDoc = document.implementation.createDocument(null, "mywordXML");

		var padreXML = xmlDoc.childNodes[0];

		// var zip = new JSZip();

		var numImg = 100;
		var numLink = 1000;
		var relsDocumentXML = "";

		var countList = 0;
		// var itemsList = 0;
		// var ultimoItemAnidado = false;

		var numberingString = '';
		var numIdString = '';

		var cols = 0;

		var tabla = false;
		var orientationRTL = false;

		var imageContentObjArr = [];

		this.checkIfInsertTag = function (newEle) {

			if (newEle.nodeName === 'w:Noinsert') return false;

			else return true;


		}


		this.createNodeBlockquote = function (node, xmlDoc) {

			var newEle = xmlDoc.createElement("w:p");
			var wpPr = xmlDoc.createElement("w:pPr");

			var wstyle = xmlDoc.createElement("w:pStyle"); wstyle.setAttribute('w:val', 'blockQuote');
			wpPr.appendChild(wstyle);

			newEle.appendChild(wpPr);


			return newEle;
		}

		/**
		* Function that checks for association in the branch with a list to insert the branch into the parent = w:body
		* Checks each node before extracting its equivalent in the docx
		* ul, ol, li, and any other element that is associates with a list
		* An element is considered to be associates with a list if its parent is not body and also any of its relatives (siblings or descendants such as children, nephews, grandchildren or children of nephews, great-grandchildren, etc.)
		* @param node the original node to check if it is associates with a list
		* @return boolean true or false
		*
		*/
		this.associatesList = function (node) {

			if (node.parentNode.nodeName === 'BODY' || node.parentNode.nodeName === 'TD') return false;

			if (node.nodeName === 'UL' || node.nodeName === 'OL' || node.nodeName === 'LI') return true;

			else {

				var padre = node.parentNode;

				for (var child = 0; child < node.childNodes.length; child++)
					if (this.associatesList(node.childNodes[child])) return true;

				for (var sibling = 0; sibling < padre.childNodes.length; sibling++) {

					if (padre.childNodes[sibling] != node) {

						var nodSibling = padre.childNodes[sibling];
						for (var nefew = 0; nefew < nodSibling.childNodes.length; nefew++)
							if (this.associatesList(nodSibling.childNodes[nefew])) return true;

					}

				}

			}

			return false;
		}


		this.createNodeXML = function (node, xmlDoc, padreXML) {


			var newEle = this.parseHTMLtoDocx(node, xmlDoc, padreXML);

			var ancestor = xmlDoc.getElementsByTagName('w:body')[0];

			if (ancestor === null) return;

			var insertTag = this.checkIfInsertTag(newEle);

			if (insertTag) {

				if (newEle.nodeName === 'w:tbl') {

					ancestor.appendChild(newEle);
					tabla = newEle;

				}

				else if (newEle.nodeName === 'w:tr') {

					tabla.appendChild(newEle);

				}

				else if (this.associatesList(node)) {

					if (newEle.nodeName === 'w:r') {

						if (ancestor.lastElementChild.nodeName === 'w:p') ancestor.lastElementChild.appendChild(newEle);

						else {

							var p = xmlDoc.createElement("w:p"); p.appendChild(newEle); ancestor.appendChild(p);

						}
					}
					else ancestor.appendChild(newEle);

				}

				else {

					padreXML.appendChild(newEle);
				}

			}

			if (newEle) padreXML = newEle;

			for (var child = 0; child < node.childNodes.length; child++)

				this.createNodeXML(node.childNodes[child], xmlDoc, padreXML);

		}

		this.createNodeTD = function (node, xmlDoc, cols) {

			var widthPercent = 100 / cols;

			var newEle = xmlDoc.createElement("w:tc");
			var wtcPr = xmlDoc.createElement("w:tcPr");
			var valign = xmlDoc.createElement("w:vAlign"); valign.setAttribute('w:val', 'bottom');
			var wtblBorders = xmlDoc.createElement("w:tblBorders");
			var wtop = xmlDoc.createElement("w:top"); wtop.setAttribute('w:val', 'single'); wtop.setAttribute('w:sz', '10'); wtop.setAttribute('w:space', '0'); wtop.setAttribute('w:color', '000000');
			var wstart = xmlDoc.createElement("w:start"); wstart.setAttribute('w:val', 'single'); wstart.setAttribute('w:sz', '10'); wstart.setAttribute('w:space', '0'); wstart.setAttribute('w:color', '000000');
			var wbottom = xmlDoc.createElement("w:bottom"); wbottom.setAttribute('w:val', 'single'); wbottom.setAttribute('w:sz', '10'); wbottom.setAttribute('w:space', '0'); wbottom.setAttribute('w:color', '000000');
			var wend = xmlDoc.createElement("w:end"); wend.setAttribute('w:val', 'single'); wend.setAttribute('w:sz', '10'); wend.setAttribute('w:space', '0'); wend.setAttribute('w:color', '000000');
			wtblBorders.appendChild(wtop);
			wtblBorders.appendChild(wstart);
			wtblBorders.appendChild(wbottom);
			wtblBorders.appendChild(wend);

			var granFather = (node.parentNode).parentNode;
			if (granFather.nodeName === 'THEAD') {
				var wshd = xmlDoc.createElement("w:shd"); wshd.setAttribute('w:val', 'clear'); wshd.setAttribute('w:fill', 'EEEEEE');
				wtcPr.appendChild(wshd);
			}


			var wtcW = xmlDoc.createElement("w:tcW"); wtcW.setAttribute('w:type', 'pct'); wtcW.setAttribute('w:w', widthPercent + '%');


			wtcPr.appendChild(wtblBorders);
			wtcPr.appendChild(wtcW);
			wtcPr.appendChild(valign);

			newEle.appendChild(wtcPr);


			if (node.childNodes.length == 0) {
				var nodep = xmlDoc.createElement("w:p");

				var noder = xmlDoc.createElement("w:r");

				var nodetext = xmlDoc.createElement("w:t");
				var text = xmlDoc.createTextNode(" ");
				nodetext.appendChild(text);
				noder.appendChild(nodetext);
				nodep.appendChild(noder);
				newEle.appendChild(nodep);
			}

			return newEle;
		}

		this.createNodeTR = function (node, xmlDoc) {

			return xmlDoc.createElement("w:tr");
		}

		this.createTableNode = function (node, xmlDoc) {

			var newEle = xmlDoc.createElement("w:tbl");
			var wtblPr = xmlDoc.createElement("w:tblPr");
			var wtblStyle = xmlDoc.createElement("w:tblStyle"); wtblStyle.setAttribute('w:val', 'TableGrid');
			var wtblW = xmlDoc.createElement("w:tblW"); wtblW.setAttribute('w:w', '5000'); wtblW.setAttribute('w:type', 'pct');
			wtblPr.appendChild(wtblStyle);
			wtblPr.appendChild(wtblW);

			var wtblBorders = xmlDoc.createElement("w:tblBorders");
			var wtop = xmlDoc.createElement("w:top"); wtop.setAttribute('w:val', 'single'); wtop.setAttribute('w:sz', '10'); wtop.setAttribute('w:space', '0'); wtop.setAttribute('w:color', '000000');
			var wstart = xmlDoc.createElement("w:start"); wstart.setAttribute('w:val', 'single'); wstart.setAttribute('w:sz', '10'); wstart.setAttribute('w:space', '0'); wstart.setAttribute('w:color', '000000');
			var wbottom = xmlDoc.createElement("w:bottom"); wbottom.setAttribute('w:val', 'single'); wbottom.setAttribute('w:sz', '10'); wbottom.setAttribute('w:space', '0'); wbottom.setAttribute('w:color', '000000');
			var wend = xmlDoc.createElement("w:end"); wend.setAttribute('w:val', 'single'); wend.setAttribute('w:sz', '10'); wend.setAttribute('w:space', '0'); wend.setAttribute('w:color', '000000');

			var windideH = xmlDoc.createElement("w:insideH"); windideH.setAttribute('w:val', 'single'); windideH.setAttribute('w:sz', '5'); windideH.setAttribute('w:space', '0'); windideH.setAttribute('w:color', '000000');
			var windideV = xmlDoc.createElement("w:insideV"); windideV.setAttribute('w:val', 'single'); windideV.setAttribute('w:sz', '5'); windideV.setAttribute('w:space', '0'); windideV.setAttribute('w:color', '000000');


			wtblBorders.appendChild(wtop);
			wtblBorders.appendChild(wstart);
			wtblBorders.appendChild(wbottom);
			wtblBorders.appendChild(wend);

			wtblBorders.appendChild(windideH);
			wtblBorders.appendChild(windideV);

			wtblPr.appendChild(wtblBorders);

			newEle.appendChild(wtblPr);

			return newEle;

		}

		this.createHeading = function (node, xmlDoc, pos) {


			if (node.parentNode.nodeName === 'BODY' || node.parentNode.nodeName === 'TD') {

				var newEle = xmlDoc.createElement("w:p");


				var newEle2 = xmlDoc.createElement("w:pPr");

				var newEle3 = xmlDoc.createElement("w:pStyle"); newEle3.setAttribute('w:val', 'Heading' + pos);

				newEle.appendChild(newEle2);
				newEle2.appendChild(newEle3);

				newEle = this.checkRTL(newEle, xmlDoc, node);
			}
			else {

				var newEle = xmlDoc.createElement("w:r");


			}


			return newEle;
		}

		this.checkRTL = function (newEle, xmlDoc, node) {

			if (node.attributes && node.attributes[0]) {

				var direction = node.attributes[0].nodeValue.replace(' ', '');

				direction = direction.replace('direction:', '');

				var res = direction.substring(0, 3);
				if (res === 'rtl') {
					orientationRTL = true;

					if (newEle.childNodes && newEle.childNodes[0] && newEle.childNodes[0].nodeName === 'w:pPr') {

						var elepPr = newEle.childNodes[0];
					}
					else var elepPr = xmlDoc.createElement("w:pPr");
					var bidi = xmlDoc.createElement("w:bidi"); bidi.setAttribute('w:val', '1');
					elepPr.appendChild(bidi);


					newEle.appendChild(elepPr);
				}
				else orientationRTL = false;
			}

			return newEle;
		}

		this.createNodeStrong = function (node, xmlDoc) {
			var newEle;

			var newEleR = xmlDoc.createElement("w:r");
			var bold = xmlDoc.createElement("w:rPr");
			var propertyBold = xmlDoc.createElement("w:b");
			newEleR.appendChild(bold);
			bold.appendChild(propertyBold);

			if (node.parentNode.nodeName === 'BODY' || node.parentNode.nodeName === 'TD') {

				newEle = xmlDoc.createElement("w:p");
				newEle.appendChild(newEleR);
			}
			else newEle = newEleR;

			return newEle;

		}

		this.createNodeEM = function (node, xmlDoc) {

			var newEle;

			var newEleR = xmlDoc.createElement("w:r");
			var nodeProperty = xmlDoc.createElement("w:rPr");
			var property = xmlDoc.createElement("w:i");
			newEleR.appendChild(nodeProperty);
			nodeProperty.appendChild(property);

			if (node.parentNode.nodeName === 'BODY' || node.parentNode.nodeName === 'TD') {

				newEle = xmlDoc.createElement("w:p");
				newEle.appendChild(newEleR);
			}
			else newEle = newEleR;

			return newEle;
		}

		this.createNodeBR = function (node, xmlDoc) {

			var newEle;

			var newEleR = xmlDoc.createElement("w:r");
			var eleBR = xmlDoc.createElement("w:br");
			newEleR.appendChild(eleBR);


			if (node.parentNode.nodeName === 'BODY' || node.parentNode.nodeName === 'TD') {

				newEle = xmlDoc.createElement("w:p");
				newEle.appendChild(newEleR);
			}
			else newEle = newEleR;


			return newEle;
		}

		this.createHyperlinkNode = function (node, xmlDoc, numLink) {

			for (var j = 0; j < node.childNodes.length; j++) {
				if (node.childNodes[j].nodeName === 'IMG') {

					var nR = xmlDoc.createElement("w:r");
					var nT = xmlDoc.createElement("w:t"); nT.setAttribute('xml:space', 'preserve');
					var t = xmlDoc.createTextNode("link format not available");
					nR.appendChild(nT);
					nT.appendChild(t);

					if (node.parentNode.nodeName === 'BODY') { var p = xmlDoc.createElement("w:p"); p.appendChild(nR); return p; }
					else return nR;
				}
			}


			var hyperEle = xmlDoc.createElement("w:hyperlink");
			if (node.href) {  // EXTERNAL LINK
				hyperEle.setAttribute('r:id', 'link' + numLink);

			}


			if (node.parentNode.nodeName === 'BODY' || node.parentNode.nodeName === 'TD') {

				var newEle = xmlDoc.createElement("w:p");

				newEle.appendChild(hyperEle);

			}
			else var newEle = hyperEle;

			return newEle;
		}

		this.createNodeLi = function (node, xmlDoc, countList, itemsList) {


			var father = node.parentNode;

			var newEle = xmlDoc.createElement("w:p");
			var pr = xmlDoc.createElement("w:pPr");
			var rstyle = xmlDoc.createElement("w:pStyle"); rstyle.setAttribute('w:val', 'ListParagraph'); //('mystyleList');   ListParagraph 

			pr.appendChild(rstyle);
			newEle.appendChild(pr);

			var numpr = xmlDoc.createElement("w:numPr");
			var wilvl = xmlDoc.createElement("w:ilvl"); wilvl.setAttribute('w:val', itemsList);
			var numId = xmlDoc.createElement("w:numId"); numId.setAttribute('w:val', countList);

			numpr.appendChild(wilvl);
			numpr.appendChild(numId);

			pr.appendChild(numpr);


			return newEle;
		}

		/**
		*  function to create a xml document node paragraph or node run
		*  @param: node html node to parse
		*  @return a new document.xml node
		*/
		this.createNodeParagraphOrRun = function (node, xmlDoc) {
			var tx = node.data;

			// si texto undefined. asignar caracter blanco
			if (!tx) tx = '';

			// SI EL TAG NO RECONOCIBLE Y PADRE = BODY. SE TRADUCE EN UN PÁRRAFO
			if (node.parentNode.nodeName === 'BODY' || node.parentNode.nodeName === 'TD') {

				var newEle = xmlDoc.createElement("w:p");

				newEle = this.checkRTL(newEle, xmlDoc, node);

			}
			else {

				var newEle = xmlDoc.createElement("w:r");

			}


			return newEle;
		}

		/**
		*  function to create text node
		*  @param node html node to convert into docx node
		*  @param xmlDoc xml document
		*  @return docx node
		*/
		this.createMyTextNode = function (node, xmlDoc) {

			var tx = node.data;

			if (!tx) tx = ''; // IF TX UNDEFINED  !!!!


			if (node.parentNode.nodeName === 'TR' || node.parentNode.nodeName === 'TABLE') return xmlDoc.createElement("w:Noinsert");
			else if (node.parentNode.nodeName === 'BODY' && (!node.data.trim())) return xmlDoc.createElement("w:Noinsert");

			var newEleR = xmlDoc.createElement("w:r");


			if (orientationRTL) {

				var elerPr = xmlDoc.createElement("w:rPr");
				var rtl = xmlDoc.createElement("w:rtl"); rtl.setAttribute('w:val', '1');
				elerPr.appendChild(rtl);

				newEleR.appendChild(elerPr);
			}

			if (node.parentNode.nodeName === 'A' && node.parentNode.href) {
				var nodeStyleLink = xmlDoc.createElement("w:rStyle"); nodeStyleLink.setAttribute("w:val", "Hyperlink");
				newEleR.appendChild(nodeStyleLink);
			}

			var nodetext = xmlDoc.createElement("w:t"); nodetext.setAttribute('xml:space', 'preserve');
			var texto = xmlDoc.createTextNode(tx);


			newEleR.appendChild(nodetext);
			nodetext.appendChild(texto);

			if (node.parentNode.nodeName === 'BODY' || node.parentNode.nodeName === 'TD') {

				var newEle = xmlDoc.createElement("w:p");
				newEle.appendChild(newEleR);
			}
			else var newEle = newEleR;



			return newEle;
		}

		this.createAstractNumListDecimal = function (numberingString, countList) {
			numberingString = `<w:abstractNum w:abstractNumId="` + countList + `" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
	<w:multiLevelType w:val="hybridMultilevel" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" />
		<w:lvl w:ilvl="0" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:start w:val="1" />
			<w:numFmt w:val="decimal" />
			<w:lvlText w:val="%1." />
			<w:lvlJc w:val="left" />
			<w:pPr><w:ind w:left="720" w:hanging="360" /></w:pPr>
		</w:lvl>
		<w:lvl w:ilvl="1" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:start w:val="1" /><w:numFmt w:val="lowerLetter" /><w:lvlText w:val="%2." /><w:lvlJc w:val="left" /><w:pPr><w:ind w:left="1440" w:hanging="360" /></w:pPr>
		</w:lvl>
		<w:lvl w:ilvl="2" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:start w:val="1" /><w:numFmt w:val="lowerRoman" /><w:lvlText w:val="%3." /><w:lvlJc w:val="right" /><w:pPr><w:ind w:left="2160" w:hanging="180" /></w:pPr>
		</w:lvl>
		<w:lvl w:ilvl="3" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:start w:val="1" /><w:numFmt w:val="decimal" /><w:lvlText w:val="%4." /><w:lvlJc w:val="left" /><w:pPr><w:ind w:left="2880" w:hanging="360" /></w:pPr>
		</w:lvl>
		<w:lvl w:ilvl="4" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:start w:val="1" /><w:numFmt w:val="lowerLetter" /><w:lvlText w:val="%5." /><w:lvlJc w:val="left" /><w:pPr><w:ind w:left="3600" w:hanging="360" /></w:pPr>
		</w:lvl>
		<w:lvl w:ilvl="5" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:start w:val="1" /><w:numFmt w:val="lowerRoman" /><w:lvlText w:val="%6." /><w:lvlJc w:val="right" /><w:pPr><w:ind w:left="4320" w:hanging="180" /></w:pPr>
		</w:lvl>
		<w:lvl w:ilvl="6" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:start w:val="1" /><w:numFmt w:val="decimal" /><w:lvlText w:val="%7." /><w:lvlJc w:val="left" /><w:pPr><w:ind w:left="5040" w:hanging="360" /></w:pPr>
		</w:lvl>
		<w:lvl w:ilvl="7" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:start w:val="1" /><w:numFmt w:val="lowerLetter" /><w:lvlText w:val="%8." /><w:lvlJc w:val="left" /><w:pPr><w:ind w:left="5760" w:hanging="360" /></w:pPr>
		</w:lvl>
		<w:lvl w:ilvl="8" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:start w:val="1" /><w:numFmt w:val="lowerRoman" /><w:lvlText w:val="%9." /><w:lvlJc w:val="right" /><w:pPr><w:ind w:left="6480" w:hanging="180" /></w:pPr>
		</w:lvl>
</w:abstractNum>`+ numberingString;



			return numberingString;
		}

		this.createAstractNumListBullet = function (numberingString, countList) {
			numberingString = `<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="` + countList + `">
	<w:multiLevelType xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="hybridMultilevel"/>
	<w:lvl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ilvl="0">
		<w:start w:val="1"/>
		<w:numFmt w:val="bullet"/>
		<w:lvlText w:val=""/>
		<w:lvlJc w:val="left"/>
		<w:pPr>
		<w:ind w:left="720" w:hanging="360"/>
		</w:pPr>
		<w:rPr>
		<w:rFonts w:hint="default" w:ascii="Symbol" w:hAnsi="Symbol"/>
		</w:rPr>
	</w:lvl>
	<w:lvl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ilvl="1">
		<w:start w:val="1"/>
		<w:numFmt w:val="bullet"/>
		<w:lvlText w:val="o"/>
		<w:lvlJc w:val="left"/>
		<w:pPr>
		<w:ind w:left="1440" w:hanging="360"/>
		</w:pPr>
		<w:rPr>
		<w:rFonts w:hint="default" w:ascii="Courier New" w:hAnsi="Courier New"/>
		</w:rPr>
	</w:lvl>
	<w:lvl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ilvl="2">
		<w:start w:val="1"/>
		<w:numFmt w:val="bullet"/>
		<w:lvlText w:val=""/>
		<w:lvlJc w:val="left"/>
		<w:pPr>
		<w:ind w:left="2160" w:hanging="360"/>
		</w:pPr>
		<w:rPr>
		<w:rFonts w:hint="default" w:ascii="Wingdings" w:hAnsi="Wingdings"/>
		</w:rPr>
	</w:lvl>
	<w:lvl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ilvl="3">
	<w:start w:val="1"/>
	<w:numFmt w:val="bullet"/>
	<w:lvlText w:val=""/>
	<w:lvlJc w:val="left"/>
	<w:pPr>
	<w:ind w:left="2880" w:hanging="360"/>
	</w:pPr>
	<w:rPr>
	<w:rFonts w:hint="default" w:ascii="Symbol" w:hAnsi="Symbol"/>
	</w:rPr>
	</w:lvl>
	<w:lvl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ilvl="4">
	<w:start w:val="1"/>
	<w:numFmt w:val="bullet"/>
	<w:lvlText w:val="o"/>
	<w:lvlJc w:val="left"/>
	<w:pPr>
	<w:ind w:left="3600" w:hanging="360"/>
	</w:pPr>
	<w:rPr>
	<w:rFonts w:hint="default" w:ascii="Courier New" w:hAnsi="Courier New"/>
	</w:rPr>
	</w:lvl>
	<w:lvl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ilvl="5">
	<w:start w:val="1"/>
	<w:numFmt w:val="bullet"/>
	<w:lvlText w:val=""/>
	<w:lvlJc w:val="left"/>
	<w:pPr>
	<w:ind w:left="4320" w:hanging="360"/>
	</w:pPr>
	<w:rPr>
	<w:rFonts w:hint="default" w:ascii="Wingdings" w:hAnsi="Wingdings"/>
	</w:rPr>
	</w:lvl>
	<w:lvl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ilvl="6">
	<w:start w:val="1"/>
	<w:numFmt w:val="bullet"/>
	<w:lvlText w:val=""/>
	<w:lvlJc w:val="left"/>
	<w:pPr>
	<w:ind w:left="5040" w:hanging="360"/>
	</w:pPr>
	<w:rPr>
	<w:rFonts w:hint="default" w:ascii="Symbol" w:hAnsi="Symbol"/>
	</w:rPr>
	</w:lvl>
	<w:lvl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ilvl="7">
	<w:start w:val="1"/>
	<w:numFmt w:val="bullet"/>
	<w:lvlText w:val="o"/>
	<w:lvlJc w:val="left"/>
	<w:pPr>
	<w:ind w:left="5760" w:hanging="360"/>
	</w:pPr>
	<w:rPr>
	<w:rFonts w:hint="default" w:ascii="Courier New" w:hAnsi="Courier New"/>
	</w:rPr>
	</w:lvl>
	<w:lvl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ilvl="8">
	<w:start w:val="1"/>
	<w:numFmt w:val="bullet"/>
	<w:lvlText w:val=""/>
	<w:lvlJc w:val="left"/>
	<w:pPr>
	<w:ind w:left="6480" w:hanging="360"/>
	</w:pPr>
	<w:rPr>
	<w:rFonts w:hint="default" w:ascii="Wingdings" w:hAnsi="Wingdings"/>
	</w:rPr>
	</w:lvl>
</w:abstractNum>`+ numberingString;



			return numberingString;
		}

		this.createNumIdList = function (numIdString, countList) {

			numIdString = '<w:num w:numId="' + countList + '"><w:abstractNumId w:val="' + countList + '"/></w:num>' + numIdString;

			return numIdString;

		}

		/**
		* Function to resize the image if necessary and it exceeds the page size
		* Returns the image dimensions in EMUs for later painting in the docx
		*
		* Perform the conversion correctly, taking into account DPI
		*
		* 1 inch = 914,400 EMUs
		*
		* 1 cm = 360,000 EMUs
		* Considering the page measures 21 x 29.7 cm
		*
		*
		*/
		this.scaleIMG = function (imgwidth, imgheight) {

			var dimensionIMG = { 'width': 0, 'height': 0 }

			// tomando  96px por pulgada . si se quiere la imagen para impresion en papel probar dividir entre 300 x pulgada 
			var width_inch = imgwidth / 96;
			var height_inch = imgheight / 96;

			var width_emu = width_inch * 914400;
			var height_emu = height_inch * 914400;

			var pgSzW = (16 * 360000); // // ancho de página en EMUs
			var pgSzH = (24.7 * 360000); // // alto de página en EMUs



			if (width_emu > pgSzW) {
				var originalW = width_emu;
				width_emu = pgSzW;
				height_emu = Math.floor(width_emu * height_emu / originalW);

			}

			if (height_emu > pgSzH) {

				var originalH = height_emu;
				height_emu = pgSzH;
				width_emu = Math.floor(height_emu * width_emu / originalH);


			}

			dimensionIMG.width = Math.floor(width_emu);
			dimensionIMG.height = Math.floor(height_emu);


			return dimensionIMG;
		}

		/**
		* Function that returns an empty node because it doesn't recognize the image format
		* @param xmlDoc xml docx document
		* @return node element with text = Not available image format
		*/
		this.nodeVoidImg = function (xmlDoc) {

			var newEle = xmlDoc.createElement("w:r");
			var nodetext = xmlDoc.createElement("w:t"); nodetext.setAttribute('xml:space', 'preserve');
			var texto = xmlDoc.createTextNode('IMAGE FORMAT NOT AVAILABLE!');


			newEle.appendChild(nodetext);
			nodetext.appendChild(texto);


			return newEle;

		}

		/**
		* Function that creates the picture node to insert into the docx document.xml
		* @param: node the source IMG node to parse
		*
		*/

		this.createDrawingNodeIMG = function (node, dataImg, xmlDoc, numImg) {


			var format = '.png';
			var nameFile = 'image' + numImg + format;

			var relashionImg = 'rId' + numImg;


			var img = document.createElement('img');
			img.src = dataImg;
			img.width = node.width;
			img.height = node.height;


			var dimensionImg = this.scaleIMG(img.width, img.height);


			var newEleImage = xmlDoc.createElement("w:r");

			var drawEle = xmlDoc.createElement("w:drawing");
			var wpinline = xmlDoc.createElement("wp:inline"); wpinline.setAttribute('distR', '0'); wpinline.setAttribute('distL', '0'); wpinline.setAttribute('distB', '0'); wpinline.setAttribute('distT', '0');
			var wpextent = xmlDoc.createElement("wp:extent"); wpextent.setAttribute('cy', dimensionImg.height); wpextent.setAttribute('cx', dimensionImg.width);
			var wpeffectExtent = xmlDoc.createElement("wp:effectExtent"); wpeffectExtent.setAttribute('b', '0'); wpeffectExtent.setAttribute('r', '0'); wpeffectExtent.setAttribute('t', '0'); wpeffectExtent.setAttribute('l', '0');
			var wpdocPr = xmlDoc.createElement("wp:docPr"); wpdocPr.setAttribute('name', nameFile); wpdocPr.setAttribute('id', numImg);

			var wpcNvGraphicFramePr = xmlDoc.createElement("wp:cNvGraphicFramePr");
			var childcNvGraphicFramePr = xmlDoc.createElement("a:graphicFrameLocks"); childcNvGraphicFramePr.setAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main'); childcNvGraphicFramePr.setAttribute('noChangeAspect', '1');

			wpcNvGraphicFramePr.appendChild(childcNvGraphicFramePr);

			var agraphic = xmlDoc.createElement("a:graphic"); agraphic.setAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
			var agraphicdata = xmlDoc.createElement("a:graphicData"); agraphicdata.setAttribute('uri', 'http://schemas.openxmlformats.org/drawingml/2006/picture');
			var pic = xmlDoc.createElement("pic:pic"); pic.setAttribute('xmlns:pic', 'http://schemas.openxmlformats.org/drawingml/2006/picture');
			var picnvPicPr = xmlDoc.createElement("pic:nvPicPr");
			var piccNvPr = xmlDoc.createElement("pic:cNvPr"); piccNvPr.setAttribute('name', nameFile); piccNvPr.setAttribute('id', numImg);
			var piccNvPicPr = xmlDoc.createElement("pic:cNvPicPr");
			picnvPicPr.appendChild(piccNvPr);
			picnvPicPr.appendChild(piccNvPicPr);

			var picblipFill = xmlDoc.createElement("pic:blipFill");
			var ablip = xmlDoc.createElement("a:blip"); ablip.setAttribute('cstate', 'print'); ablip.setAttribute('r:embed', relashionImg);
			var astretch = xmlDoc.createElement("a:stretch");
			var afillRect = xmlDoc.createElement("a:fillRect");
			astretch.appendChild(afillRect);

			picblipFill.appendChild(ablip);
			picblipFill.appendChild(astretch);


			var picspPr = xmlDoc.createElement("pic:spPr");
			var axfrm = xmlDoc.createElement("a:xfrm");
			var aoff = xmlDoc.createElement("a:off"); aoff.setAttribute('y', '0'); aoff.setAttribute('x', '0');
			var aext = xmlDoc.createElement("a:ext"); aext.setAttribute('cy', dimensionImg.height); aext.setAttribute('cx', dimensionImg.width);
			axfrm.appendChild(aoff);
			axfrm.appendChild(aext);


			var aprstGeom = xmlDoc.createElement("a:prstGeom"); aprstGeom.setAttribute('prst', 'rect');
			var aavLst = xmlDoc.createElement("a:avLst");
			aprstGeom.appendChild(aavLst);

			picspPr.appendChild(axfrm);
			picspPr.appendChild(aprstGeom);


			pic.appendChild(picnvPicPr);
			pic.appendChild(picblipFill);
			pic.appendChild(picspPr);


			agraphicdata.appendChild(pic);

			agraphic.appendChild(agraphicdata);

			wpinline.appendChild(wpextent);
			wpinline.appendChild(wpeffectExtent);
			wpinline.appendChild(wpdocPr);
			wpinline.appendChild(wpcNvGraphicFramePr);
			wpinline.appendChild(agraphic);

			drawEle.appendChild(wpinline);

			newEleImage.appendChild(drawEle);


			return newEleImage;
		}


		this.addImgTypeToContentTypes = function () {
			if (doc === undefined || doc === null || doc.trim() == "") {
				return;
			}
			let contentTypes = '<Default ContentType="image/png" Extension="png"/>';

			let cTypes = doc.getZip().files["[Content_Types].xml"].asText();

			if (!cTypes.includes('Extension="png"')) {
				cTypes = cTypes.replace("</Types>", contentTypes + "</Types>");
				doc.getZip().file("[Content_Types].xml", cTypes);
			}
		}




		/**
		* function that check if string starts with prefix
		*/
		this.stringStartsWith = function (string, prefix) {
			return string.slice(0, prefix.length) == prefix;
		}


		//  FUNCION parseHTMLtoDocx node
		this.parseHTMLtoDocx = function (node, xmlDoc, padreXML) {

			var encabezados = ['H1', 'H2', 'H3', 'H4', 'H5', 'H6'];

			if (node.nodeName === 'BODY') {

				var newEle = xmlDoc.createElement("w:body");

			}

			else if (encabezados.indexOf(node.nodeName) >= 0) {

				var pos = encabezados.indexOf(node.nodeName) + 1;
				var newEle = this.createHeading(node, xmlDoc, pos);

			}

			else if (node.nodeName === 'P') {

				var newEle = this.createNodeParagraphOrRun(node, xmlDoc);

			}

			else if (node.nodeName === 'STRONG') {

				var newEle = this.createNodeStrong(node, xmlDoc);

			}

			else if (node.nodeName === 'EM') {

				var newEle = this.createNodeEM(node, xmlDoc);

			}
			else if (node.nodeName === 'BR') {
				var newEle = this.createNodeBR(node, xmlDoc);

			}

			else if (node.nodeName === 'A') {

				if (node.href) {  // EXTERNAL LINK

					var idLink = 'link' + numLink;
					relsDocumentXML = relsDocumentXML + '<Relationship Id="' + idLink + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="' + node.href + '" TargetMode="External"/>';

					var newEle = this.createHyperlinkNode(node, xmlDoc, numLink);

					numLink++;
				}
				else var newEle = this.createNodeParagraphOrRun(node, xmlDoc);

			}

			else if (node.nodeName === 'UL' || node.nodeName === 'OL') {

				if (node.parentNode.nodeName != 'LI') countList++;

				if (node.nodeName === 'UL') numberingString = this.createAstractNumListBullet(numberingString, countList);
				else if (node.nodeName === 'OL') numberingString = this.createAstractNumListDecimal(numberingString, countList);

				numIdString = this.createNumIdList(numIdString, countList);

				var newEle = xmlDoc.createElement("w:Noinsert");

			}

			else if (node.nodeName === 'LI') {

				var levelList = 0;
				var granFather = node.parentNode.parentNode;

				if (granFather) {

					var abu = true;

					while (abu && granFather.nodeName === 'LI') {
						levelList++;

						if (granFather.parentNode.parentNode) { abu = true; granFather = granFather.parentNode.parentNode; }
						else abu = false;
					}
				}

				var newEle = this.createNodeLi(node, xmlDoc, countList, levelList);

			}

			else if (node.nodeName === '#text') {

				var newEle = this.createMyTextNode(node, xmlDoc);

			}

			else if (node.nodeName === 'IMG'/*&& this.isProcessImg()*/) {

				var format = '.png';
				var nameFile = 'image' + numImg + format;
				var relashionImg = 'rId' + numImg;
				var imgEle;


				// console.log("node.nodeName === 'IMG':", node, node.attributes.src)
				var dataImg = node.attributes.src.value;

				if (!this.stringStartsWith(dataImg, 'data:')) {
					imgEle = this.nodeVoidImg(xmlDoc);
				}
				else {

					var srcImg = dataImg.replace(/^data:image\/.+;base64,/, ' ');

					// zip.file("word/media/" + nameFile, srcImg, { base64: true });
					imageContentObjArr.push({
						imageFileName: nameFile,
						src: srcImg,
						base64: true
					});

					relsDocumentXML = relsDocumentXML + '<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/' + nameFile + '" Id="' + relashionImg + '" />';

					imgEle = this.createDrawingNodeIMG(node, dataImg, xmlDoc, numImg);

					numImg++;
				}


				if (node.parentNode.nodeName === 'BODY' || node.parentNode.nodeName === 'TD') {

					var newEle = xmlDoc.createElement("w:p");
					newEle.appendChild(imgEle);

				}
				else var newEle = imgEle;

			}
			// IF TABLE
			else if (node.nodeName === 'TABLE') {

				var newEle = this.createTableNode(node, xmlDoc);

			}
			else if (node.nodeName === 'THEAD') {

				var newEle = xmlDoc.createElement('w:Noinsert');

			}
			else if (node.nodeName === 'TBODY') {

				var newEle = xmlDoc.createElement('w:Noinsert');

			}
			else if (node.nodeName === 'TR') {

				var newEle = this.createNodeTR(node, xmlDoc);

				cols = 0;
				for (var c = 0; c < node.childNodes.length; c++) {
					if (node.childNodes[c].nodeName === 'TD') cols++;
				}

				if (cols == 0) {

					var eleTD = document.createElement("TD"); node.appendChild(eleTD);

				}


			}
			else if (node.nodeName === 'TD') {

				var newEle = this.createNodeTD(node, xmlDoc, cols);
			}

			else if (node.nodeName === 'BLOCKQUOTE') {

				var newEle = this.createNodeBlockquote(node, xmlDoc);
			}

			else {

				var newEle = this.createNodeParagraphOrRun(node, xmlDoc);

			}




			return newEle;
		}

		this.createNodeXML(nodeParent[0], xmlDoc, padreXML); // first node is body[0]


		let relsNumberingDocumentXML = '<Relationship Id="Rnumbering1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml" />';


		var xml = xmlDoc.getElementsByTagName('w:body');

		// var content = this.createDocx(xml[0].outerHTML, zip);

		// return content;

		// console.log("doc:", doc);

		if (doc && numberingString != "") {

			let cabeceraNumbering = '<?xml version="1.0" encoding="utf-8"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">';
			var footerNumbering = '</w:numbering>';

			let u_o_list = doc.getZip().files["word/numbering.xml"].asText();

			if (u_o_list === undefined || u_o_list === null || u_o_list == "") {
				numberingString = cabeceraNumbering + numberingString + numIdString + footerNumbering;
				doc.getZip().file("word/numbering.xml", numberingString);
			} else {
				numberingString = numberingString + numIdString;

				// console.log("u_o_list:", u_o_list)
				if (!u_o_list.includes('</w:numbering>')) {
					u_o_list = u_o_list.replace("/>", ">");
					u_o_list += numberingString + footerNumbering;
				} else {
					u_o_list = u_o_list.replace("</w:numbering>", numberingString + "</w:numbering>");
				}

				doc.getZip().file("word/numbering.xml", u_o_list);
			}
		}


		if (doc) {

			let rels = doc.getZip().files["word/_rels/document.xml.rels"].asText();

			//relsNumberingDocumentXML
			if (!rels.includes("numbering.xml")) {
				relsDocumentXML += relsNumberingDocumentXML;
			}

			rels = rels.replace("</Relationships>", relsDocumentXML + "</Relationships>");

			doc.getZip().file("word/_rels/document.xml.rels", rels);

			if (is_img && imageContentObjArr && imageContentObjArr.length > 0) {

				imageContentObjArr.forEach((imgObj) => {
					doc.getZip().file("word/media/" + imgObj.imageFileName, imgObj.src, { base64: true });
				});

				// if (this.isProcessImg()){

				// 	this.addImgTypeToContentTypes();
				// }

			}
		}

		callback(xml[0].innerHTML, is_img);
		return xml[0].innerHTML;
	}
}
