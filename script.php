<?php

require_once('vendor/autoload.php');

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\IOFactory;
use Shuchkin\SimpleXLSX;

$file = 'noduplicates.xlsx';
$output = 'test.xlsx';

$reader = IOFactory::createReader('Xlsx');
$spreadsheet = $reader->load($file);

if( $xlsx = SimpleXLSX::parse($file) ){

    $emailsContainer = [];
    print_r("loading...");

    for($x = 0; $x < count( $xlsx->rows()); $x+=1)
    {
        $sheet = $spreadsheet->getActiveSheet();
        $email = $xlsx->rows()[$x][1];
        $incrementor = 1;
        // print_r($email);

        if(!in_array($email, $emailsContainer)){
            $tempArray = array_push($emailsContainer, $email);
        } else {
            print_r("\nDuplicate found at line " . $x+1);
            $sheet->removeRow($x+1, 1); // delete row
            $sheet->insertNewRowBefore($x+1); // insert blank row where the deleted row was originally placed
        }
    }
} else {
    echo SimpleXLSX::parseError();
}

$writer = IOFactory::createWriter($spreadsheet, "Xlsx");
$writer->save($output);


?>