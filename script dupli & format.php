<?php

require_once('vendor/autoload.php');

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\IOFactory;
use Shuchkin\SimpleXLSX;

$date = date("Y-m-d");
$file = 'noduplicates.xlsx';
$output = "test-$date.xlsx";
$columnTargeted = 1;

$reader = IOFactory::createReader('Xlsx');
$spreadsheet = $reader->load($file);

function useRegex($input) {
    $regex = '/([a-zA-Z]+(\\.[a-zA-Z]+)+)@fr\\.[a-zA-Z]ran[a-zA-Z]avi[a-zA-Z]\\.com/i';
    return preg_match($regex, $input);
}

// $input1 = "MNI helene.orsini@fr.transavia.com";

// print_r(useRegex($input1));

if( $xlsx = SimpleXLSX::parse($file) ){

    $emailsContainer = [];
    print_r("loading...");

    for($x = 1; $x < count( $xlsx->rows()); $x+=1)
    {
        $sheet = $spreadsheet->getActiveSheet();
        $email = $xlsx->rows()[$x][$columnTargeted];

        if(!in_array($email, $emailsContainer)){
            $tempArray = array_push($emailsContainer, $email);
            if(useRegex($email) == 0)
            {
                print_r("\nWrong input found at line " . $x+1);
                $sheet->removeRow($x+1, $columnTargeted); // delete row
                $sheet->insertNewRowBefore($x+1); // insert blank row where the deleted row was originally placed
            }
        } else {
            print_r("\nDuplicate found at line " . $x+1);
            $sheet->removeRow($x+1, $columnTargeted); // delete row
            $sheet->insertNewRowBefore($x+1); // insert blank row where the deleted row was originally placed
        }
    }
} else {
    echo SimpleXLSX::parseError();
}

$writer = IOFactory::createWriter($spreadsheet, "Xlsx");
$writer->save($output);


?>