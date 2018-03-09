var gulp = require('gulp');
var path = require('path');
const gutil = require('gulp-util');
const gulpif = require('gulp-if');
const rename = require('gulp-rename');
var convertExcel = require('excel-as-json').processFile;
const loadJsonFile = require('load-json-file');
var foreach = require('gulp-foreach');

var Vtt = require('vtt-creator');
var fs = require('fs');

function hmsToSecondsOnly(str) {
    var p = str.toString().split(':'),
        s = 0,
        m = 1;
    while (p.length > 0) {
        s += m * parseInt(p.pop(), 10);
        m *= 60;
    }
    return s;
}

gulp.task('convertXL', () => {
    /*
    convertExcel('transcript.xlsx', 'transcript.json', {
        omitEmtpyFields: true
    }, function () {
        var v = new Vtt();
        loadJsonFile('transcript.json').then(json => {
            // promise to process it 
            for (var i = 0; i < json.length; i++) {
                v.add(hmsToSecondsOnly(json[i]['Start']), hmsToSecondsOnly(json[i]['End']), `${json[i]['Who?']} ${json[i]['Responses']}`);
            }
            fs.writeFileSync('transcript.vtt', v.toString());
        });
    }); */
    return (gulp.src('transcripts/excel/**.xlsx').pipe(foreach(function (stream, file) {
        var name = path.basename(file.path, '.xlsx');
        //console.log(name, 'filename');
        convertExcel('transcripts/excel/' + name + '.xlsx', 'json/' + name + '.json', {
            omitEmtpyFields: true
        }, function () {
            //console.log('here');
            var v = new Vtt();
            loadJsonFile('json/' + name + '.json').then(json => {
                // promise to process it 
                console.log(name);
                for (var i = 0; i < json.length; i++) {
                    v.add(hmsToSecondsOnly(json[i]['Start']), hmsToSecondsOnly(json[i]['End']), `${json[i]['Who']}: ${json[i]['Quote']}`, 'line:-1 position:12% align:start size:75%');
                }
                fs.writeFileSync('vtt/' + name + '.vtt', v.toString());
               
            });
           
        });
        return stream;
    })))
});