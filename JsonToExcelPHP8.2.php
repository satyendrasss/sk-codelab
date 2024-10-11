<?php
// error_reporting(0);
//FOR PHP 8.2 JSON TO EXCEL
require realpath(dirname(__FILE__)) . '../../exportExcel/phpExcel/autoload.php';
    use PhpOffice\PhpSpreadsheet\IOFactory;
    use PhpOffice\PhpSpreadsheet\Spreadsheet;

function convertJsonToExcel($jsondata)
{
    ini_set('memory_limit', '-1');
    // require realpath(dirname(__FILE__)) . '../../PHPExcel/Classes/PHPExcel.php';
    

    // $objPHPExcel = new PHPExcel();
    $objPHPExcel = new Spreadsheet();
    $dataArr = json_decode($jsondata, true);
    $arrKeys = array_keys($dataArr);

    function getColLetter($i)
    {
        $COLS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
        $ct = ($i > 25) ? floor($i / 26) : 0;
        $ret = $COLS[$i % 26];
        while ($ct--)
            $ret .= $ret;
        return $ret;
    }

    foreach ($arrKeys as $keyIndex => $arrKey) {

        $objPHPExcel->createSheet($keyIndex);
        $objPHPExcel->setActiveSheetIndex($keyIndex);
        $activeSheet1 = $objPHPExcel->getActiveSheet();
        $cols = array_keys($dataArr[$arrKey][0]);
        // prepare header row
        foreach ($cols as $i => $col) {
            $activeSheet1->setCellValue(getColLetter($i) . 1, $col);
        }
        // prepare the rest
        foreach ($dataArr[$arrKey] as $i => $row) {
            foreach ($cols as $j => $col) {
                $activeSheet1->setCellValue(getColLetter($j) . ($i + 2), $row[$col]);
            }
        }
        $objPHPExcel->getActiveSheet($keyIndex)->setTitle($arrKey);
    }
    $writer = IOFactory::createWriter($objPHPExcel, 'Xls');
$fileName = 'exported_excel_' . time() . '.xls';

// Send headers to force download the file
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="' . $fileName . '"');
header('Cache-Control: max-age=0');

// Save the file to output
$writer->save('php://output');


    // $file_name = 'Excel-Data';
    // header('Content-Type: application/vnd.ms-excel');
    // header('Content-Disposition: attachment;filename="' . $file_name . '.xls"');
    // header('Cache-Control: max-age=0');

    // $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
    // $objWriter->save('php://output');
}

$jsondata  = array(
    'data' => array(
        array('name' => 'sateyndra', 'age' => '26', 'gender' => 'male'),
        array('name' => 'sateyndra-1', 'age' => '24', 'gender' => 'Male'),
    ),
    'members' => array(
        array('memname' => 'abc', 'aa' => '22', 'edu' => '10th'),
        array('memname' => 'abc1', 'aa' => '20', 'edu' => '12th'),
    )
);
$json = json_encode($jsondata);

convertJsonToExcel($json);
