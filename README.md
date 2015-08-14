# ExcelArrayAccess
Access to PHPExcel cells as a php array

## Require
You need to install PHPExcel library
https://github.com/PHPOffice/PHPExcel

## Use the ExcelArrayAccess
You can read the test file to look how the script works.

You can create a new file by using that command :

`$arr = new ExcelArray(__DIR__.'/files/test.xlsx'); //works event test.xlsx doesn't exist

$arr[0] = array('test', 'autre', 'test');

$arr[1] = array('test1', 'autre4', 'test1');`

You can get values from an excel file like this :

`$arr = new ExcelArray(__DIR__.'/files/test.xlsx'); //test.xlsx needs to exist
 
var_dump($arr);
 
foreach($arr as $k => $val){

    var_dump($val);
    
}`

More informations : http://www.raphael-goncalves.fr/blog/fichier-excel-et-l-interface-arrayaccess
