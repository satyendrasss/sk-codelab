<?php
    error_reporting(E_ALL);
    require realpath(dirname(__FILE__)) . '/Classes/PHPExcel.php';
    
    $objPHPExcel = new PHPExcel();
    
    // DEMO ONLY (potentially unsafe)
    
    $data = json_decode($_POST['json']);
    $key = $_POST['key'];
    $cols = explode('|', $_POST['cols']);
    
    $objPHPExcel->setActiveSheetIndex(0);
    $activeSheet = $objPHPExcel->getActiveSheet();
    
    // TODO deal with more than 26 columns... does Excel double letters up or what?
    function getColLetter ($i) {
    $COLS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    $ct = ($i > 25) ? floor($i / 26) : 0;
    $ret = $COLS[$i % 26];
        while ($ct--)
            $ret .= $ret;
    return $ret;
    }

    // prepare header row
    foreach ($cols as $i=>$col) {
        $activeSheet->setCellValue(getColLetter($i) . 1, $col);
    }

    // prepare the rest
    foreach ($data->$key as $i=>$row) {
        foreach ($cols as $j=>$col) {
            $activeSheet->setCellValue(getColLetter($j) . ($i + 2), $row->$col);
        }
    }
    
    
    header('Content-Type: application/vnd.ms-excel');
    header('Content-Disposition: attachment;filename="'. $_POST['filename'] . '.xls"');
    header('Cache-Control: max-age=0');
    
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
    $objWriter->save('php://output');

  /*
  Expamle
  {
   "data": [
     {
       "columna1": "lorem ipsum",
       "columna2": "ipsum presure",
       "columna3": "ignords ipsum",
     }
   ]
}

*/
?>
