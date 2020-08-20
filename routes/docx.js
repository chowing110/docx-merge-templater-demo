var express = require('express');
var router = express.Router();

var DocxMerger = require('docx-merger');

var PizZip = require('pizzip');
var Docxtemplater = require('docxtemplater');

var fs = require('fs');
var path = require('path');

router.post('/form', function(req, res, next) {
    console.log(req.body);
    //docx-merger
    var data = req.body;
    
    function merge() {
        var files = [];

        var keys = Object.keys(data);
    
        for (i=0;i<keys.length;i++) {
            if (keys[i].match('fname') != null) {
                var temp = fs.readFileSync(path.resolve(__dirname, '../public/docx/input/'+data[keys[i]]+'.docx'), 'binary');
                files.push(temp);
            }
        }
    
        var docx = new DocxMerger({},files);
    
        docx.save('nodebuffer',function (data) {
            // fs.writeFile("output.zip", data, function(err){/*...*/});
            fs.writeFile("./public/docx/input/merged.docx", data, function(err){/*...*/});
        });

        return new Promise(resolve => {
            setTimeout(() => {
              resolve('resolved');
            }, 2000);
        });
    }

    // docxtemplater
    // The error object contains additional information when logged with JSON.stringify (it contains a properties object containing all suberrors).
    function replaceErrors(key, value) {
        if (value instanceof Error) {
            return Object.getOwnPropertyNames(value).reduce(function(error, key) {
                error[key] = value[key];
                return error;
            }, {});
        }
        return value;
    }

    function errorHandler(error) {
        console.log(JSON.stringify({error: error}, replaceErrors));

        if (error.properties && error.properties.errors instanceof Array) {
            const errorMessages = error.properties.errors.map(function (error) {
                return error.properties.explanation;
            }).join("\n");
            console.log('errorMessages', errorMessages);
            // errorMessages is a humanly readable message looking like this :
            // 'The tag beginning with "foobar" is unopened'
        }
        throw error;
    }

    async function template() {
        //Load the docx file as a binary
        const result = await merge();
        console.log(result);

        var fname = "merged";
        var content = fs
            .readFileSync(path.resolve(__dirname, '../public/docx/input/'+fname+'.docx'), 'binary');
    
        var zip = new PizZip(content);
        var doc;
        try {
            doc = new Docxtemplater(zip);
        } catch(error) {
            // Catch compilation errors (errors caused by the compilation of the template : misplaced tags)
            errorHandler(error);
        }
    
        var obj = data;
    
        //set the templateVariables
        doc.setData(obj);
    
        try {
            // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
            doc.render()
        }
        catch (error) {
            // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
            errorHandler(error);
        }
    
        var buf = doc.getZip()
                    .generate({type: 'nodebuffer'});
    
        // buf is a nodejs buffer, you can either write it to a file or do anything else with it.
        fs.writeFileSync(path.resolve(__dirname, '../public/docx/output/output.docx'), buf);
    }
    
    template();

    res.render('pages/download');
});

module.exports = router;