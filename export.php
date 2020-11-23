<?php 
include("connection.php");
include('Classes/PHPExcel.php');

$objPHPExcel    =   new PHPExcel();


$gJobID = $_GET["JobID"];
$gSeqID =  $_GET["SeqID"];
$gOption =  $_GET["Option"];

if($gOption == 1){
  $sql = "EXEC [dbo].[TH-SP-ExportInfo] @jobid = $gJobID, @importSeq = $gSeqID ";
}else{
   $sql = "EXEC [dbo].[TH-SP-ExportInfo] @jobid = $gJobID, @importSeq = -1";
}

function findNum($str){

  $num = '';
  for ($i = 0; $i < strlen($str); $i++) {

      if (is_numeric($str[$i])) {
          $num .= $str[$i];
      }
  }
  substr($num, 0, -4);
  return $num;

}

function __CompareRinning($pNum,$cNum)
{
  $nResult = 1;
   if ($pNum+1 < $cNum)  
   
   // cnum+1 == pnum สลับบรรทัด
   
   {$nResult = 0;} 
   
   return  $nResult;
}

$sqlHeader = "SELECT FieldName FROM ImExFieldMapping
where JobID = $gJobID ORDER BY FieldSequence";

$resultHeader = sqlsrv_query($conn, $sqlHeader);
 
$objPHPExcel->setActiveSheetIndex(0);

$rowArray = 'A';
$countRowArray = 1;
$columnCount = 0;

while($rowH =  sqlsrv_fetch_array($resultHeader)){
$objPHPExcel->getActiveSheet()->SetCellValue($rowArray.$countRowArray , $rowH['FieldName']);
$objPHPExcel->getActiveSheet()->getColumnDimension($rowArray)->setAutoSize(true);
$rowArray++;
$columnCount++;
}

 $colString = PHPExcel_Cell::stringFromColumnIndex($columnCount-1);


$rCol = $colString.'1';

$objPHPExcel->getActiveSheet()
    ->getStyle("A1:$rCol")
    ->applyFromArray(
      array(
          'fill' => array(
              'type' => PHPExcel_Style_Fill::FILL_SOLID,
              'color' => array('rgb' => 'FFA500')
          )
      )

  );
  $objPHPExcel->getActiveSheet()->getStyle("A1:$rCol")->getFont()->setBold(true);
  $objPHPExcel->getActiveSheet()->getStyle("A1:$rCol")->getFont()->setSize(12);
  $objPHPExcel->getActiveSheet()->getStyle("A1:$rCol")->getAlignment()
  ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
  $objPHPExcel->getActiveSheet()->getStyle("A1:$rCol")->getAlignment()
    ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    
$result = sqlsrv_query($conn, $sql);

$i = 2;
$num1 = 1;
$num2 = 0;
$row = 1;
$fNum = -1; //First Numbering
$cNum = -1;  //Compareing Numbering

while($rowData =  sqlsrv_fetch_array($result)){


   for($column = 0; $column< $columnCount; $column++){
    if($column == 1){
      $cNum = findNum($rowData[1]); 
      if ($fNum != -1 ) {
          if (__CompareRinning($fNum,$cNum) != 0 )  {
                $objPHPExcel
                ->getActiveSheet()
                ->setCellValueByColumnAndRow($column,$i,$rowData[$column]);
                //set style
                $objPHPExcel->getActiveSheet()
                ->getStyle("A$num1:AA$num1")
                ->applyFromArray(
                  array(
                      'fill' => array(
                          'type' => PHPExcel_Style_Fill::FILL_SOLID,
                          'color' => array('rgb' => 'FFFF00')
                      )
                  )
                );

          } else {
            $objPHPExcel
            ->getActiveSheet()
            ->setCellValueByColumnAndRow($column,$i,($rowData[$column]));
          }
      } else {

            $objPHPExcel
            ->getActiveSheet()
            ->setCellValueByColumnAndRow($column,$i,($rowData[$column]));



      }
    
    }
    else{
      $objPHPExcel
      ->getActiveSheet()
      ->setCellValueByColumnAndRow($column,$i,($rowData[$column]));
    } //if
   
    
    
    } //for
   $i++;
   $row++;
   $num1++;
   //if $i != 2 and rowdata[1] == '' $fnum = $fnmu else   $fNum = findNum($rowData[1]);
   $fNum = findNum($rowData[1]);
} //While

$objPHPExcel->getActiveSheet()
->getStyle("H1:H$i")
->getNumberFormat()
->setFormatCode('0');

$sqlFileName = "SELECT * FROM ImExJobs where JobID = $gJobID ";


$mADate = date('Ymd_H:m:s');
$JobDesc = "Motor-Claims_".$mADate;

$objWriter  =   new PHPExcel_Writer_Excel2007($objPHPExcel);

header('Content-Type: application/vnd.ms-excel'); 
header('Content-Disposition: attachment;filename="'.$JobDesc.'.xlsx"');
header('Cache-Control: max-age=0');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');  
$objWriter->save('php://output');



?>
