//
// The MIT License (MIT)
// Copyright (c) 2016 Ray Jakes
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated
// documentation files (the "Software"), to deal in the Software without restriction, including without limitation the
// rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or substantial portions
// of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
// WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
// COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//

#target photoshop

// Make Photoshop the frontmost application
app.bringToFront();

// define the black color
var scolor = new SolidColor;
scolor.rgb.hexValue = '000000';

// define the opacity to use
var useopacity = 10;

// adds an alignment corner to a document
function addCorner(thisDoc, position, widthpx, heightpx) {


    // create the layer and fill it in
    var cornerLayer = thisDoc.artLayers.add();
    cornerLayer.name = position;

    var rectsize = parseFloat(thisDoc.height.toString().replace(" px", ""))/10;

    thisDoc.selection.select([[0, 0], [0, rectsize], [rectsize, rectsize], [rectsize, 0]], SelectionType.REPLACE, 0, false);
    thisDoc.selection.fill(scolor);

    // move it to the specified corner
    var Position = cornerLayer.bounds;

    switch (position) {

        case "topleft":
            Position[0] = -(rectsize/2);
            Position[1] = -(rectsize/2);
            cornerLayer.translate(Position[0], Position[1]);
            break;

        case "topright":
            Position[0] = -(rectsize/2)+widthpx;
            Position[1] = -(rectsize/2);
            cornerLayer.translate(Position[0], Position[1]);
            break;


        case "bottomleft":
            Position[0] = -(rectsize/2);
            Position[1] = -(rectsize/2)+heightpx;
            cornerLayer.translate(Position[0], Position[1]);
            break;

        case "bottomright":
            Position[0] = -(rectsize/2)+widthpx;
            Position[1] = -(rectsize/2)+heightpx;
            cornerLayer.translate(Position[0], Position[1]);
            break;

    }


    // rotate it
    cornerLayer.rotate(45, AnchorPosition.MIDDLECENTER);

    // opacity
    cornerLayer.opacity = useopacity;


}

function processImage() {
    // all the strings that need localized
    //var strDeleteAllEmptyLayersHistoryStepName = localize("$$$/JavaScripts/DeleteAllEmptyLayers/Menu=Delete All Empty Layers" );

    // define the destination doc/tile size
    var dwidth = 6.71;
    var dheight = 10.14;


    // do the work
    var docRef = app.activeDocument; // remember the document, the selected layer, the visibility setting of the selected layer

    // try and work from a flat file
    try {
        docRef.flatten();
    } catch (e) {
        ; // do nothing
    }

    // set the global dimensions to cm
    var originalUnit = preferences.rulerUnits
    preferences.rulerUnits = Units.INCHES;

    // get the current image dimensions
    var w = parseFloat(docRef.width.toString().replace(" in", ""));
    var h = parseFloat(docRef.height.toString().replace(" in", ""));

    // calculate the number of tiles running left to right, top to bottom
    var countx = Math.ceil(w / dwidth);
    var county = Math.ceil(h / dheight);
    var totalPages = (countx*county);

    if (countx*county > 60){
        alert("Page count is too high ("+totalPages.toString()+").");
        return;
    }

    // loop through this grid
    var tilenum = 1;
    var tempFileList = new Array();   // array of File

    for (var y = 0; y < county; y++) {
        for (var x = 0; x < countx; x++) {


            // set the source document to the front most
            app.activeDocument = docRef;

            // calculate the selection size
            var startx = dwidth * x;
            var starty = dheight * y;
            var endx = startx + dwidth;
            var endy = starty + dheight;

            if (endx > w) {
                endx = w;
            }

            if (endy > h) {
                endy = h;
            }

            // fix selection size for pixels instead of inches (selection only uses pixels)
            startx *= docRef.resolution;
            starty *= docRef.resolution;
            endx *= docRef.resolution;
            endy *= docRef.resolution;

            // make the selection and copy
            var shapeRef = Array = [[startx, starty], [startx, endy], [endx, endy], [endx, starty]];
            docRef.selection.select(shapeRef);
            docRef.selection.copy();

            // create a new document and paste
            var tileDoc = app.documents.add(dwidth, dheight, docRef.resolution, "tile-" + tilenum.toString(), NewDocumentMode.RGB, DocumentFill.WHITE);
            var contentLayer = tileDoc.paste();

            // make sure the paste layer is at 0,0
            var Position = contentLayer.bounds;
            Position[0] = 0 - Position[0];
            Position[1] = 0 - Position[1];
            contentLayer.translate(-Position[0], -Position[1]);

            // stroke the pasted layer with a 1px border
            var strokeLayer = tileDoc.artLayers.add();
            strokeLayer.name = "stroke";

            preferences.rulerUnits = Units.PIXELS;

            tileDoc.selection.select([[0, 0], [0, parseFloat(contentLayer.bounds[3].toString().replace(" px", ""))], [parseFloat(contentLayer.bounds[2].toString().replace(" px", "")), parseFloat(contentLayer.bounds[3].toString().replace(" px", ""))], [parseFloat(contentLayer.bounds[2].toString().replace(" px", "")), 0]], SelectionType.REPLACE, 0, false);
            tileDoc.selection.stroke (scolor, 1, StrokeLocation.INSIDE,ColorBlendMode.NORMAL,100);
            tileDoc.selection.deselect();

            strokeLayer.opacity = useopacity;


            // add the grey alignment corners
            if (x != 0 || y != 0 ) {
                addCorner(tileDoc, "topleft", parseFloat(contentLayer.bounds[2].toString().replace(" px", "")), parseFloat(contentLayer.bounds[3].toString().replace(" px", "")));
            }

            if (x != countx-1 || y != 0) {
                addCorner(tileDoc, "topright", parseFloat(contentLayer.bounds[2].toString().replace(" px", "")), parseFloat(contentLayer.bounds[3].toString().replace(" px", "")));
            }

            if (x != 0 || y != county-1) {
                addCorner(tileDoc, "bottomleft", parseFloat(contentLayer.bounds[2].toString().replace(" px", "")), parseFloat(contentLayer.bounds[3].toString().replace(" px", "")));
            }

            if (x != countx-1 || y != county-1) {
                addCorner(tileDoc, "bottomright", parseFloat(contentLayer.bounds[2].toString().replace(" px", "")), parseFloat(contentLayer.bounds[3].toString().replace(" px", "")));
            }

            preferences.rulerUnits = Units.INCHES;

            // create a new layer and type in the tile number
            var textLayer = tileDoc.artLayers.add();
            textLayer.kind = LayerKind.TEXT;

            var textItem = textLayer.textItem;
            textItem.kind = TextType.POINTTEXT;

            textItem.justification = Justification.CENTER;
            textItem.color = scolor;
            textItem.font = "ArialMT";
            textItem.size = 250;
            textItem.contents = tilenum.toString();

            textItem.position = [dwidth/2, 6];

            textLayer.opacity = useopacity;     // bring the opacity of the layer down

            // save it (for the PDF presentation later)
            var tempFile = docRef.path+"/"+tilenum.toString() + ".psd";
            tempFile = new File(tempFile);
            tempFileList[tilenum - 1] = tempFile;
            tileDoc.saveAs(tempFile);


            tileDoc.close(SaveOptions.DONOTSAVECHANGES);

            // next tile...
            tilenum += 1;
        }
    }

    // close the source document
    app.activeDocument = docRef;
    //docRef.close(SaveOptions.DONOTSAVECHANGES);

    // Restore original ruler unit setting
    app.preferences.rulerUnits = originalUnit;

    // run the PDF Presentation automation
    var outputFile = File(docRef.path+"/"+docRef.name+".pdf");

    var options = new PresentationOptions;

    makePDFPresentation(tempFileList, outputFile, options);

    // delete all the temporary image files
    for (var i = 0; i < tempFileList.length; i++) {
        tempFileList[i].remove();
    }

    // alert the user
    alert("Presentation file saved to: " + outputFile.fsName);

}

if (documents.length) {
    processImage();
} else {
    alert('You must open a file before running this script.');
}