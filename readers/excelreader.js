const readXlsxFile = require('read-excel-file/node');

function main(){
    if (process.argv.length <= 2) {
        throw new Error("\033[1;31m\nMissing Excel file.\n");
    }
    var excel_file = process.argv[2];

    readXlsxFile(excel_file).then((rows) => {
	//console.log(rows);
	let row = rows[0];
	let ok_num = false;
	let ok_name = false;
	let ok_sn = false;
	if(row[0] == 'Num.')
	{
	    ok_num = true;
	}
	else
	{
	    console.log('\033[1;31mError on cell A1 ' + 'value given: ' + row[0] + ' expected value: Num. \n');
	}
	if(row[1] == 'Name')
	{
	    ok_name = true;
	}
	else
	{
	    console.log('\033[1;31mError on cell B1 ' + 'value given: ' + row[1] + ' expected value: Name \n');
	}
	if(row[2] == 'Second Name')
	{
	    ok_sn = true;
	}
	else
	{
	   console.log('\033[1;31mError on cell C1 ' + 'value given: ' + row[2] + ' expected value: Second Name \n');
	}
	if(ok_num && ok_name && ok_sn)
	{
	    console.log('\033[1;92m\Everything is ok');
	}
	else
	{
	    console.log('Something is wrong with excel file');
	}
    });
}
main();



