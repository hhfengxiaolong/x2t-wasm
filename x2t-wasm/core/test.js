var x2t = require('./x2t.js');

var isNode = true;

// console.log(x2t);
x2t.run()

function init() {
    x2t.FS.mkdir('/working');
    if (isNode) {
      x2t.FS.mkdir('/source');
      setupNodeFS();
    }
}

function setupNodeFS() {
    x2t.FS.mount(x2t.NODEFS, { root: '.' }, '/source');
    x2t.FS.mount(x2t.MEMFS, { root: './working' }, '/working');
    console.log("Mounted");
}

/*
 Converting File 
 
 This function is for use when testing under nodejs
*/
function convertFile(fileName, outputFormat) {
    console.log("Converting File for " + fileName + " to " + outputFormat);
    var data = x2t.FS.readFile("/source/" + fileName)
    var outputData = convertData(data, fileName, outputFormat);
    return x2t.FS.writeFile("/source/" + fileName + "." + outputFormat, outputData);
}

/*
 Converting Data

 This function converts a data in a specific format to the outputformat
 The filename extension needs to represent the input format
 Example: fileName=cryptpad.bin outputFormat=xlsx
*/
function convertData(data, fileName, outputFormat) {
 console.log("Converting Data for " + fileName + " to " + outputFormat);
 // writing file to mounted working disk (in memory)
 x2t.FS.writeFile('/working/' + fileName, data);
 var params = "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
      + "<TaskQueueDataConvert xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">"
      + "<m_sFileFrom>/working/" + fileName + "</m_sFileFrom>"
      + "<m_sFileTo>/working/" + fileName + "." + outputFormat + "</m_sFileTo>"
      + "<m_bIsNoBase64>false</m_bIsNoBase64>"
      + "</TaskQueueDataConvert>"
 // writing params file to mounted working disk (in memory)
    x2t.FS.writeFile('/working/params.xml', params);
 // running conversion
    x2t.ccall("runX2T", ["number"], ["string"], ["/working/params.xml"]);
 // reading output file from working disk (in memory)
    return x2t.FS.readFile('/working/' + fileName + "." + outputFormat);
}

x2t.onRuntimeInitialized = function() {
 // Init x2t js module
 init();

 if (isNode) {
  console.log("x2t nodejs started");
  convertFile("cryptpad.bin", "xlsx");
  console.log("x2t nodejs finished");
 } else {
  console.log("x2t ready");
 }
}
