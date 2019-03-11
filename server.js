var Express = require('express');
var multer = require('multer');
var bodyParser = require('body-parser');
var path = require('path');
const excelToJson = require('convert-excel-to-json');
var app = Express();
app.use(bodyParser.json());
var excelFileName=null;

var storage = multer.diskStorage({
  destination: function (req, file, callback) {
    callback(null, './uploads');
  },
  filename: function (req, file, callback) {
    excelFileName = file.fieldname + '-' + Date.now() + '.xlsx';
    callback(null, excelFileName);
    //console.log(excelFileName);
  }
});
var upload = multer({ storage: storage }).single('userPhoto');

//controller for index page
app.get("/", function (req, res) {
  res.sendFile(__dirname + "/index.html");
});

//controller for uploading file
app.post("/api/Upload", function (req, res) {
  //console.log("Calling post");
  var html = null;
  upload(req, res, function (err) {
    if (err) {
      console.log("Error happend");
      return res.end("Error uploading file.");
    }
    //console.log(excelFileName);
    const result = excelToJson({
      sourceFile: path.join(__dirname + '/uploads/' + excelFileName)
    });

    //console.log(result['Sheet1']);
    var obj = result['Sheet1'];
    html = '<html> <head> <p> Data from excel </p> <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.4.0/css/bootstrap.min.css"> <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.min.js"></script> <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.4.0/js/bootstrap.min.js"></script> </head> <body><p>';
    html += '<table class="table table-sm table-dark" style="width:100% background-repeat:no-repeat; width:450px;margin:0;" border="1" cellpadding="0" cellspacing="0" border="0">'
    for (var i = 0; i < obj.length; i++) {
      html += "<tr>";
      console.log(obj[i]);
      for (var key in obj[i]) {
        if (obj[i].hasOwnProperty(key)) {
          //console.log(key + ": " + obj[i][key]);
          html += "<td>" + obj[i][key] + "</td>";
        }
      }
      html += "</tr>"
    }
    html += '</table></p></body></html>'
    //console.log(html);
    return res.send(html);
  });

});

app.listen(3000, function (a) {
  console.log("Listening to port 3000");
});