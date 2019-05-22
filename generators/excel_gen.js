const officegen = require('officegen')
const fs = require('fs')
const _path = require('path');

var outDir = _path.join(__dirname,'../docs/gen');
// Create an empty Excel object:
let xlsx = officegen('xlsx')
let docx = officegen('docx')

// Officegen calling this function after finishing to generate the xlsx document:
xlsx.on('finalize', function(written) {
  console.log(
    'Finish to create a Microsoft Excel document.'
  )
})

// Officegen calling this function to report errors:
xlsx.on('error', function(err) {
  console.log(err)
})

let sheet = xlsx.makeNewSheet()
sheet.name = 'Officegen Excel'

// Create a new paragraph:
let pObj = docx.createP()

pObj.addText(sheet.setCell('A1', "Num."), { highlight: true });

// Add data using setCell:


sheet.setCell('A2', 1)
sheet.setCell('A3', 2)
sheet.setCell('A4', 3)
sheet.setCell('A5', 4)
sheet.setCell('A6', 5)
sheet.setCell('B1', "Name")
sheet.setCell('B2', "JJ")
sheet.setCell('B3', "SS")
sheet.setCell('B4', "YY")
sheet.setCell('B5', "AA")
sheet.setCell('B6', "FF")
sheet.setCell('C1', "Second Name")
sheet.setCell('C2', "SSF")
sheet.setCell('C3', "FFF")
sheet.setCell('C4', "WEEE")
sheet.setCell('C5', "SEEE")
sheet.setCell('C6', "GEEE")

// Let's generate the Excel document into a file:

let out = fs.createWriteStream(_path.join(outDir,'example.xlsx'))

out.on('error', function(err) {
  console.log(err)
})

// Async call to generate the output file:
xlsx.generate(out)
