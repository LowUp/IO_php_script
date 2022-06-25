<?php

require_once('vendor/autoload.php');

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\IOFactory;
use Shuchkin\SimpleXLSX;

$date = date("Y-m-d");
$file = 'raw_data_export-1655557327.xlsx';
// $file = 'noduplicates.xlsx';
$output = "Output-$date.csv";
$columnTargeted = 1;

$reader = IOFactory::createReader('Xlsx');
$spreadsheet = $reader->load($file);

function useRegex($input) {
    $regex = '/([a-zA-Z]+(\\.[a-zA-Z]+)+)@fr\\.[a-zA-Z]ran[a-zA-Z]avi[a-zA-Z]\\.com/i';
    return preg_match($regex, $input);
}

// Removes duplicates and wrong input 
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
                print_r("\nWrong email format found at line " . $x+1);
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

// $writer = IOFactory::createWriter($spreadsheet, "Xlsx");
// $writer->save($output);

$writer = IOFactory::createWriter($spreadsheet, "Csv");
$writer -> setEnclosure('');
$writer->save($output);


// Removes blank lines
print_r("\nloading...");
$lines = file($output);
$num_rows = count($lines);

foreach ($lines as $lineNo => $line) {
    
    $csv = str_getcsv($line);

    if (!array_filter($csv)) {
        unset($lines[$lineNo]);
    }
}

file_put_contents($output, $lines);

print_r("\nDone...");

?>