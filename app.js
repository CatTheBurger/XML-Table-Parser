// app.js - пример дочернего модуля

var fs = require("fs");
var xmlParser = require('xml2js');

class App {

	constructor () {
		this.name = 'NodeJS Application';
	}

	run () {
		fs.readFile('./table.xml','utf-8', function(err, data) {
		    if(err) throw err;
            var content = data;
            
            xmlParser.parseString(content, function (err, result) {
			    var object = result;

			    var data = object["Workbook"]["ss:Worksheet"][0]["Table"][0]["Row"][0];
			    var dataLength = data["Cell"].length;
			    var numbers = [];
			    var sum = 0;
			    var finalResult = 0;

			    for(var i = 0;i < dataLength;i++){
			    	numbers.push(data["Cell"][i]["Data"][0]["_"]);
			    	console.log("LOG:" + numbers[numbers.length-1]);
			    	sum += parseInt(numbers[numbers.length-1]); 
			    }

			    finalResult = sum / dataLength;

			    console.log("Sum:" + sum);
			    console.log("Result:" + finalResult);

			    finalResult = Math.round(finalResult);

			    console.log("Final Result:" + finalResult);

			    console.log("===================");

			    object["Workbook"]["ss:Worksheet"][0]["Table"][0]["Row"][1]["Cell"][0]["Data"][0]["_"] = finalResult; //A2
			    //console.log(object["Workbook"]["ss:Worksheet"][0]["Table"][0]["Row"][1]["Cell"][0]["Data"][0]["_"]);

			    var builder = new xmlParser.Builder();
				var xml = builder.buildObject(object);
				console.log(xml);

				fs.writeFile("./table.xml", xml, function(err, data) {
					if (err) console.log(err);
					console.log("Successfully Written to File.");
				});
			});
        });
	}

}

module.exports = App;