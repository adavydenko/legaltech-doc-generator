const http = require("http");
const fs = require("fs");
const path = require("path");
const yaml = require("js-yaml");

const generateDocx = require("generate-docx");

// 1. Read contents of the template -> authors.yml

// 1.1 Get document definition, or throw exception on error
try {
  var filePath = path.join(__dirname, 'templates', 'authors.yml')
  var doc = yaml.safeLoad(fs.readFileSync(filePath, "utf8"));
  console.log("Dcoument parsed: " + doc);
} catch (e) {
  console.log(e);
}

// 1.2 Get test data
var testPath = path.join(__dirname, 'templates', 'authors', 'test.json');
var contractData = JSON.parse(fs.readFileSync(testPath, 'utf8'));

// 2. Save generated documnets to file

doc.templates.forEach(template => {
  let templateFilePath = path.join(__dirname, 'templates', template);
  let outputFilePath = path.join(__dirname, 'templates', 'authors', 'out-' + path.basename(template));
  console.log(`Processing template: ${template} -> ${outputFilePath} ...`);
  
  let options = {
    template: {
      filePath: templateFilePath,
      data: contractData
    },
    save: {
      filePath: outputFilePath
    }
  };

  generateDocx(options)
    .then(console.log)
    .catch(console.error);

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
