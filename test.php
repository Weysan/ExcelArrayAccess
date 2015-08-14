<?php

require __DIR__.'/PhpExcel/Classes/PHPExcel.php';
require __DIR__.'/ExcelArray.php';

/* add data to an excel file */
$arr = new ExcelArray(__DIR__.'/files/test.xlsx');
$arr[0] = array('test', 'autre', 'test');
$arr[1] = array('test1', 'autre4', 'test1');

/* get data from an existing file */
$arr = new ExcelArray(__DIR__.'/files/test.xlsx');
 
var_dump($arr);
 
foreach($arr as $k => $val){
    var_dump($val);
}
