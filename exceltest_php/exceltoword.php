<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;

function joinList($items): String
{
    $letters = array("A", "B", "C", "D", "E", "F", "G");
    $count = count($items);
    $items[0] = trim($items[0]) . "\n";
    for ($i = 0; $i < $count - 2; $i++) {
        if ($items[$i + 1] && trim($items[$i + 1]) != "") {
            $items[$i + 1] = " " . $letters[$i] . trim($items[$i + 1]);
        } else {
            $item[$i + 1] = "";
        }
    }
    $items[$count - 1] = " 答案：" . trim($items[$count - 1]) . "\n";
    return implode("", $items);
}

function process($array): array
{
    $content = array();
    $choice = array();
    $answer = array();
    $questionindex = array();
    $count = count($array);
    for ($i = 0; $i < $count; $i++) {
        if (preg_match("/题目/", $array[$i])) {
            array_push($content, $i);
        }
        if (preg_match("/A|B|C|D|E|F|G/i", $array[$i])) {
            array_push($choice, $i);
        }
        if (preg_match("/答案/", $array[$i])) {
            array_push($answer, $i);
        }
    }
    array_push($questionindex, $content[0]);
    array_push($questionindex, ...$choice);
    array_push($questionindex, $answer[0]);
    return $questionindex;
}

function createWord($str, $filename)
{
    $phpWord = new \PhpOffice\PhpWord\PhpWord();
    $fontStyle = new \PhpOffice\PhpWord\Style\Font();
    $fontStyle->setSize(8);
    $content = explode("\n", $str);
    $section = $phpWord->addSection();
    for ($i = 0; $i < count($content); $i++) {
        $myTextElement = $section->addText($content[$i]);
        $myTextElement->setFontStyle($fontStyle);
    }
    $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
    $objWriter->save($filename);
}

$inputFileType = 'Xlsx';
$inputFileName = 'safety.xlsx';

$reader = IOFactory::createReader($inputFileType);
$reader->setLoadAllSheets();
$spreadsheet = $reader->load($inputFileName);
$totalsheets = $spreadsheet->getSheetCount();

for ($i = 0; $i < $totalsheets; $i++) {
    $worksheet = $spreadsheet->getSheet($i);
    $data = array();
    $sum = array();
    foreach ($worksheet->getRowIterator() as $row) {
        $highestColumn = $worksheet->getHighestColumn();
        $highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn);
        if ($row->getRowIndex() == 1) {
            $temp = array();
            for ($col = 1; $col <= $highestColumnIndex; ++$col) {
                array_push($temp, $worksheet->getCellByColumnAndRow($col, $row->getRowIndex()));
            }
            $data = process($temp);
        }
        if ($row->getRowIndex() > 1) {
            $everyrow = array();
            $index = "(" . (string)($row->getRowIndex() - 1) . ")";
            for ($j = 0; $j < count($data); $j++) {
                array_push($everyrow, $worksheet->getCellByColumnAndRow($data[$j] + 1, $row->getRowIndex()));
            }
            $everyrowlist = joinList($everyrow);
            $everyrowlist = $index . $everyrowlist;
            array_push($sum, $everyrowlist);
        }
        $total = implode("", $sum);
        $title = $worksheet->getTitle() . "php.docx";
        createWord($total, $title);
    }
}
