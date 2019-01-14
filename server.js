
const http = require("http");
const fs = require("fs");
const path = require("path");
const yaml = require("js-yaml");

const JSZip = require('jszip');
const Docxtemplater = require('docxtemplater');
var expressions= require('angular-expressions');

// define your filter functions here, for example, to be able to write {clientname | lower}
expressions.filters.lower = function(input) {
  // This condition should be used to make sure that if your input is undefined, your output will be undefined as well and will not throw an error
  if(!input) return input;
  return input.toLowerCase();
}

var angularParser = function(tag) {
  return {
    get: tag === '.' ? function(s){ return s;} : function(s) {
      return expressions.compile(tag.replace(/(’|“|”)/g, "'"))(s);
    }
  };
}

// 1. Read contents of the template -> authors.yml
// 1.1 Get document definition, or throw exception on error
try {
  var ymldoc = process.argv[2];
  
  var documentDirectory = path.parse(ymldoc).name;
  var ymlFilePath = path.join(__dirname, 'templates', ymldoc)
  
  var doc = yaml.safeLoad(fs.readFileSync(ymlFilePath, "utf8"));
  //console.log("Dcoument parsed: " + doc);
  
  // 1.2 Get test data
  var testFilePath = path.join(__dirname, 'templates', documentDirectory, 'test.json');
  var contractData = JSON.parse(fs.readFileSync(testFilePath, 'utf8'));
  
  console.log(`Document definition: ${ymldoc}...`);
  console.log(`Document file path: ${ymlFilePath}`);
  console.log(`Test file path: ${testFilePath}`);

} catch (e) {
  console.log(e);
}


// 2. Save generated documnets to file
doc.templates.forEach(template => {
  let templateFilePath = path.join(__dirname, 'templates', template);
  let outputFilePath = path.join(__dirname, 'templates', documentDirectory, 'out-' + path.basename(template));
  console.log(`Processing template: ${template} -> ${outputFilePath} ...`);

  //Load the docx file as a binary
  let content = fs.readFileSync(templateFilePath, 'binary');

  let zip = new JSZip(content);
  let docx = new Docxtemplater();
  docx.loadZip(zip);

  docx.setOptions({parser:angularParser});

  //set the templateVariables
  docx.setData(contractData);

  try {
    // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
    docx.render()
  }
  catch (error) {
      var e = {
          message: error.message,
          name: error.name,
          stack: error.stack,
          properties: error.properties,
      }
      console.log(JSON.stringify({error: e}));
      // The error thrown here contains additional information when logged with JSON.stringify (it contains a property object).
      throw error;
  }

  var buf = docx.getZip().generate({type: 'nodebuffer'});
  // buf is a nodejs buffer, you can either write it to a file or do anything else with it.
  fs.writeFileSync(outputFilePath, buf);
});

/*
http.createServer(function (req, res){
  var filePath = path.join..console

  var size

  res.writeHead(200, {
    'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'Content-Length': size
  });

  var readStream = ...
  readStream.pipe(res);

})
.listen(3000);
*/
