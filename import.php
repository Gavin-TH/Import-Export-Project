<?php  
include("connection.php");

     $pJobID = $_POST["JobID"];
     $pJobSeq = $_POST["countRound"];
     $pOption = $_POST["radio1"];

     function ms_escape_string($data) 
     {
     if ( !isset($data) or empty($data) ) return '';
     if ( is_numeric($data) ) return $data;
     
     $data = str_replace('="', '', $data);
    
     if(substr($data,strlen($data)-1,1) == '"')
     {
          $data = substr($data,0,strlen($data)-1);
     }
      
     
     
     $non_displayables = array(
         '/%0[0-8bcef]/',            // url encoded 00-08, 11, 12, 14, 15
         '/%1[0-9a-f]/',             // url encoded 16-31
         '/[\x00-\x08]/',            // 00-08
         '/\x0b/',                   // 11
         '/\x0c/',                   // 12
         '/[\x0e-\x1f]/',             // 14-31
         '/%3D%22/'
     );

     foreach ( $non_displayables as $regex )
     $data = preg_replace( $regex, '', $data );
     $data = str_replace("'", "''", $data );
     return $data;
     }


     
     function removeSymbol($pData){
          
          $pData = str_replace('="', '', $pData);
          $pData = str_replace('"', '', $pData);

           return $pData;
 
 
      }

      

 if(!empty($_FILES["excel_file"]))  
 {  
      
      $file_array = explode(".", $_FILES["excel_file"]["name"]);  
      if($file_array[1] == "xlsx" or $file_array[1] == "xls" )  
      {  
           include("PHPExcel/Classes/PHPExcel/IOFactory.php");  

          // $User = getenv('USERNAME');
          // $strUsr = rtrim($User, "$");

           $query = "SELECT CASE WHEN MAX(ImportSequence) IS NULL THEN 1
           ELSE MAX(ImportSequence) +1  END AS [ROUND]
           FROM ImExInfo100";

           $stmt = sqlsrv_query($conn, $query);

           $Round = sqlsrv_fetch_array($stmt);
       
           $dRound =  $Round['ROUND'];
           $mADate = date('Y-m-d H:m:s');
          
           $object = PHPExcel_IOFactory::load($_FILES["excel_file"]["tmp_name"]); 
           
           $break = false;

           foreach($object->getWorksheetIterator() as $worksheet)  
           {
              if (!$break)
              {
                $highestRow = $worksheet->getHighestRow();  
                for($row=2; $row<=$highestRow; $row++)  
                {  


                    if (removeSymbol(ms_escape_string( $worksheet->getCellByColumnAndRow(0, $row)->getValue())) == '')
                    {
                      
                         // echo '<script type="text/javascript">';
                         // echo ' alert("กรุณาตรวจสอบข้อมูลใน Excel ใหม่")'; 
                         // echo '</script>';
                         echo "<script> swal(\"\",\" พบค่าว่างที่คอลัมน์แรก กรุณาตรวจสอบข้อมูล\", \"error\")
                         .then(willRedirect => {
                              if (willRedirect) {
                                window.location.reload();
                              }
                            })
                            ; </script>";

                         $break = TRUE;
                    break;
                    }
              
                    $JobID = $_POST["JobID"];

                    $ImportDate = $ADate;

                    $ImportSequence =  $_POST["countRound"];
                    
                    $ImportRows = $row;

                    $value001   = ms_escape_string( $worksheet->getCellByColumnAndRow(0, $row)->getValue()); 
                    $value002   = ms_escape_string( $worksheet->getCellByColumnAndRow(1, $row)->getvalue());  
                    //$value002   = PHPExcel_Shared_Date::ExcelToPHPObject($value002)->format('d/m/Y');
                    $value003   = ms_escape_string( $worksheet->getCellByColumnAndRow(2, $row)->getValue());  
                    $value004   = ms_escape_string( $worksheet->getCellByColumnAndRow(3, $row)->getValue());  
                    $value005   = ms_escape_string( $worksheet->getCellByColumnAndRow(4, $row)->getValue());  
                    $value006   = ms_escape_string( $worksheet->getCellByColumnAndRow(5, $row)->getValue());  
                    $value007   = ms_escape_string( $worksheet->getCellByColumnAndRow(6, $row)->getValue());                
                    $value008   = ms_escape_string( $worksheet->getCellByColumnAndRow(7, $row)->getValue()); 
                    $value009   = ms_escape_string( $worksheet->getCellByColumnAndRow(8, $row)->getValue()); 
                    //$value009   = PHPExcel_Shared_Date::ExcelToPHPObject($value009)->format('d/m/Y H:i');
                    $value010   = ms_escape_string( $worksheet->getCellByColumnAndRow(9, $row)->getValue()); 
                    $value011   = ms_escape_string( $worksheet->getCellByColumnAndRow(10, $row)->getValue()); 
                    $value012   = ms_escape_string( $worksheet->getCellByColumnAndRow(11, $row)->getValue()); 
                    $value013   = ms_escape_string( $worksheet->getCellByColumnAndRow(12, $row)->getValue()); 
                    $value014   = ms_escape_string( $worksheet->getCellByColumnAndRow(13, $row)->getValue()); 
                    $value015   = ms_escape_string( $worksheet->getCellByColumnAndRow(14, $row)->getValue()); 
                    $value016   = ms_escape_string( $worksheet->getCellByColumnAndRow(15, $row)->getValue()); 
                    $value017   = ms_escape_string( $worksheet->getCellByColumnAndRow(16, $row)->getValue()); 
                    $value018   = ms_escape_string( $worksheet->getCellByColumnAndRow(17, $row)->getValue()); 
                    $value019   = ms_escape_string( $worksheet->getCellByColumnAndRow(18, $row)->getValue()); 
                    $value020   = ms_escape_string( $worksheet->getCellByColumnAndRow(19, $row)->getValue()); 
                    $value021   = ms_escape_string( $worksheet->getCellByColumnAndRow(20, $row)->getValue()); 
                    $value022   = ms_escape_string( $worksheet->getCellByColumnAndRow(21, $row)->getValue()); 
                    $value023   = ms_escape_string( $worksheet->getCellByColumnAndRow(22, $row)->getValue()); 
                    $value024   = ms_escape_string( $worksheet->getCellByColumnAndRow(23, $row)->getValue()); 
                    $value025   = ms_escape_string( $worksheet->getCellByColumnAndRow(24, $row)->getValue()); 
                    $value026   = ms_escape_string( $worksheet->getCellByColumnAndRow(25, $row)->getValue()); 
                    $value027   = ms_escape_string( $worksheet->getCellByColumnAndRow(26, $row)->getValue()); 
                    $value028   = ms_escape_string( $worksheet->getCellByColumnAndRow(27, $row)->getValue()); 
                    $value029   = ms_escape_string( $worksheet->getCellByColumnAndRow(28, $row)->getValue()); 
                    $value030   = ms_escape_string( $worksheet->getCellByColumnAndRow(29, $row)->getValue()); 
                    $value031   = ms_escape_string( $worksheet->getCellByColumnAndRow(30, $row)->getValue());  
                    $value032   = ms_escape_string( $worksheet->getCellByColumnAndRow(31, $row)->getValue());  
                    $value033   = ms_escape_string( $worksheet->getCellByColumnAndRow(32, $row)->getValue()); 
                    $value034   = ms_escape_string( $worksheet->getCellByColumnAndRow(33, $row)->getValue());   
                    $value035   = ms_escape_string( $worksheet->getCellByColumnAndRow(34, $row)->getValue()); 
                    $value036   = ms_escape_string( $worksheet->getCellByColumnAndRow(35, $row)->getValue()); 
                    $value037   = ms_escape_string( $worksheet->getCellByColumnAndRow(36, $row)->getValue()); 
                    $value038   = ms_escape_string( $worksheet->getCellByColumnAndRow(37, $row)->getValue()); 
                    $value039   = ms_escape_string( $worksheet->getCellByColumnAndRow(38, $row)->getValue()); 
                    $value040   = ms_escape_string( $worksheet->getCellByColumnAndRow(39, $row)->getValue()); 
                    $value041   = ms_escape_string( $worksheet->getCellByColumnAndRow(40, $row)->getValue()); 
                    $value042   = ms_escape_string( $worksheet->getCellByColumnAndRow(41, $row)->getValue()); 
                    $value043   = ms_escape_string( $worksheet->getCellByColumnAndRow(42, $row)->getValue()); 
                    $value044   = ms_escape_string( $worksheet->getCellByColumnAndRow(43, $row)->getValue()); 
                    $value045   = ms_escape_string( $worksheet->getCellByColumnAndRow(44, $row)->getValue()); 
                    $value046   = ms_escape_string( $worksheet->getCellByColumnAndRow(45, $row)->getValue()); 
                    $value047   = ms_escape_string( $worksheet->getCellByColumnAndRow(46, $row)->getValue()); 
                    $value048   = ms_escape_string( $worksheet->getCellByColumnAndRow(47, $row)->getValue()); 
                    $value049   = ms_escape_string( $worksheet->getCellByColumnAndRow(48, $row)->getValue()); 
                    $value050   = ms_escape_string( $worksheet->getCellByColumnAndRow(49, $row)->getValue()); 
                    $value051   = ms_escape_string( $worksheet->getCellByColumnAndRow(50, $row)->getValue()); 
                    $value052   = ms_escape_string( $worksheet->getCellByColumnAndRow(51, $row)->getValue()); 
                    $value053   = ms_escape_string( $worksheet->getCellByColumnAndRow(52, $row)->getValue()); 
                    $value054   = ms_escape_string( $worksheet->getCellByColumnAndRow(53, $row)->getValue()); 
                    $value055   = ms_escape_string( $worksheet->getCellByColumnAndRow(54, $row)->getValue()); 
                    $value056   = ms_escape_string( $worksheet->getCellByColumnAndRow(55, $row)->getValue()); 
                    $value057   = ms_escape_string( $worksheet->getCellByColumnAndRow(56, $row)->getValue()); 
                    $value058   = ms_escape_string( $worksheet->getCellByColumnAndRow(57, $row)->getValue()); 
                    $value059   = ms_escape_string( $worksheet->getCellByColumnAndRow(58, $row)->getValue()); 
                    $value060   = ms_escape_string( $worksheet->getCellByColumnAndRow(59, $row)->getValue()); 
                    $value061   = ms_escape_string( $worksheet->getCellByColumnAndRow(60, $row)->getValue()); 
                    $value062   = ms_escape_string( $worksheet->getCellByColumnAndRow(61, $row)->getValue()); 
                    $value063   = ms_escape_string( $worksheet->getCellByColumnAndRow(62, $row)->getValue()); 
                    $value064   = ms_escape_string( $worksheet->getCellByColumnAndRow(63, $row)->getValue()); 
                    $value065   = ms_escape_string( $worksheet->getCellByColumnAndRow(64, $row)->getValue()); 
                    $value066   = ms_escape_string( $worksheet->getCellByColumnAndRow(65, $row)->getValue()); 
                    $value067   = ms_escape_string( $worksheet->getCellByColumnAndRow(66, $row)->getValue()); 
                    $value068   = ms_escape_string( $worksheet->getCellByColumnAndRow(67, $row)->getValue()); 
                    $value069   = ms_escape_string( $worksheet->getCellByColumnAndRow(68, $row)->getValue()); 
                    $value070   = ms_escape_string( $worksheet->getCellByColumnAndRow(69, $row)->getValue()); 
                    $value071   = ms_escape_string( $worksheet->getCellByColumnAndRow(70, $row)->getValue()); 
                    $value072   = ms_escape_string( $worksheet->getCellByColumnAndRow(71, $row)->getValue()); 
                    $value073   = ms_escape_string( $worksheet->getCellByColumnAndRow(72, $row)->getValue()); 
                    $value074   = ms_escape_string( $worksheet->getCellByColumnAndRow(73, $row)->getValue()); 
                    $value075   = ms_escape_string( $worksheet->getCellByColumnAndRow(74, $row)->getValue()); 
                    $value076   = ms_escape_string( $worksheet->getCellByColumnAndRow(75, $row)->getValue()); 
                    $value077   = ms_escape_string( $worksheet->getCellByColumnAndRow(76, $row)->getValue()); 
                    $value078   = ms_escape_string( $worksheet->getCellByColumnAndRow(77, $row)->getValue()); 
                    $value079   = ms_escape_string( $worksheet->getCellByColumnAndRow(78, $row)->getValue()); 
                    $value080   = ms_escape_string( $worksheet->getCellByColumnAndRow(79, $row)->getValue()); 
                    $value081   = ms_escape_string( $worksheet->getCellByColumnAndRow(80, $row)->getValue()); 
                    $value082   = ms_escape_string( $worksheet->getCellByColumnAndRow(81, $row)->getValue()); 
                    $value083   = ms_escape_string( $worksheet->getCellByColumnAndRow(82, $row)->getValue()); 
                    $value084   = ms_escape_string( $worksheet->getCellByColumnAndRow(83, $row)->getValue()); 
                    $value085   = ms_escape_string( $worksheet->getCellByColumnAndRow(84, $row)->getValue()); 
                    $value086   = ms_escape_string( $worksheet->getCellByColumnAndRow(85, $row)->getValue()); 
                    $value087   = ms_escape_string( $worksheet->getCellByColumnAndRow(86, $row)->getValue()); 
                    $value088   = ms_escape_string( $worksheet->getCellByColumnAndRow(87, $row)->getValue()); 
                    $value089   = ms_escape_string( $worksheet->getCellByColumnAndRow(88, $row)->getValue()); 
                    $value090   = ms_escape_string( $worksheet->getCellByColumnAndRow(89, $row)->getValue()); 
                    $value091   = ms_escape_string( $worksheet->getCellByColumnAndRow(90, $row)->getValue()); 
                    $value092   = ms_escape_string( $worksheet->getCellByColumnAndRow(91, $row)->getValue()); 
                    $value093   = ms_escape_string( $worksheet->getCellByColumnAndRow(92, $row)->getValue()); 
                    $value094   = ms_escape_string( $worksheet->getCellByColumnAndRow(93, $row)->getValue()); 
                    $value095   = ms_escape_string( $worksheet->getCellByColumnAndRow(94, $row)->getValue()); 
                    $value096   = ms_escape_string( $worksheet->getCellByColumnAndRow(95, $row)->getValue()); 
                    $value097   = ms_escape_string( $worksheet->getCellByColumnAndRow(96, $row)->getValue()); 
                    $value098   = ms_escape_string( $worksheet->getCellByColumnAndRow(97, $row)->getValue()); 
                    $value099   = ms_escape_string( $worksheet->getCellByColumnAndRow(98, $row)->getValue()); 
                    $value100   = ms_escape_string( $worksheet->getCellByColumnAndRow(99, $row)->getValue()); 

                    $CreateDT = $ADate;
                    $CreateID = "T562195";
                    $LastUpdateDT = "";
                    $LastUpdateID = "";

                    $SQL = "SELECT COUNT(Value005) as [num] FROM ImExInfo100 WHERE Value005 = '$value005'";
                    $result = sqlsrv_query($conn,$SQL);
                    $num = sqlsrv_fetch_array($result);

                    if($num['num'] == 0)
                    {
                         $query = "INSERT INTO [dbo].[ImExInfo100]
                                             ([JobID]
                                             ,[ImportDate]
                                             ,[ImportSequence]
                                             ,[ImportRows]
                                             ,[Value001]
                                             ,[Value002]
                                             ,[Value003]
                                             ,[Value004]
                                             ,[Value005]
                                             ,[Value006]
                                             ,[Value007]
                                             ,[Value008]
                                             ,[Value009]
                                             ,[Value010]
                                             ,[Value011]
                                             ,[Value012]
                                             ,[Value013]
                                             ,[Value014]
                                             ,[Value015]
                                             ,[Value016]
                                             ,[Value017]
                                             ,[Value018]
                                             ,[Value019]
                                             ,[Value020]
                                             ,[Value021]
                                             ,[Value022]
                                             ,[Value023]
                                             ,[Value024]
                                             ,[Value025]
                                             ,[Value026]
                                             ,[Value027]
                                             ,[Value028]
                                             ,[Value029]
                                             ,[Value030]
                                             ,[Value031]
                                             ,[Value032]
                                             ,[Value033]
                                             ,[Value034]
                                             ,[Value035]
                                             ,[Value036]
                                             ,[Value037]
                                             ,[Value038]
                                             ,[Value039]
                                             ,[Value040]
                                             ,[Value041]
                                             ,[Value042]
                                             ,[Value043]
                                             ,[Value044]
                                             ,[Value045]
                                             ,[Value046]
                                             ,[Value047]
                                             ,[Value048]
                                             ,[Value049]
                                             ,[Value050]
                                             ,[Value051]
                                             ,[Value052]
                                             ,[Value053]
                                             ,[Value054]
                                             ,[Value055]
                                             ,[Value056]
                                             ,[Value057]
                                             ,[Value058]
                                             ,[Value059]
                                             ,[Value060]
                                             ,[Value061]
                                             ,[Value062]
                                             ,[Value063]
                                             ,[Value064]
                                             ,[Value065]
                                             ,[Value066]
                                             ,[Value067]
                                             ,[Value068]
                                             ,[Value069]
                                             ,[Value070]
                                             ,[Value071]
                                             ,[Value072]
                                             ,[Value073]
                                             ,[Value074]
                                             ,[Value075]
                                             ,[Value076]
                                             ,[Value077]
                                             ,[Value078]
                                             ,[Value079]
                                             ,[Value080]
                                             ,[Value081]
                                             ,[Value082]
                                             ,[Value083]
                                             ,[Value084]
                                             ,[Value085]
                                             ,[Value086]
                                             ,[Value087]
                                             ,[Value088]
                                             ,[Value089]
                                             ,[Value090]
                                             ,[Value091]
                                             ,[Value092]
                                             ,[Value093]
                                             ,[Value094]
                                             ,[Value095]
                                             ,[Value096]
                                             ,[Value097]
                                             ,[Value098]
                                             ,[Value099]
                                             ,[Value100]
                                             ,[CreateDT]
                                             ,[CreateID])

                                VALUES 

                                ('".$JobID."'
                                ,GETDATE()
                                ,'".$ImportSequence."'
                                ,'".$ImportRows."'
                                ,N'".$value001."'
                                ,N'".$value002."'
                                ,N'".$value003."'
                                ,N'".$value004."'
                                ,N'".$value005."'
                                ,N'".$value006."'
                                ,N'".$value007."'
                                ,N'".$value008."'
                                ,N'".$value009."'
                                ,N'".$value010."'
                                ,N'".$value011."'
                                ,N'".$value012."'
                                ,N'".$value013."'
                                ,N'".$value014."'
                                ,N'".$value015."'
                                ,N'".$value016."'
                                ,N'".$value017."'
                                ,N'".$value018."'
                                ,N'".$value019."'
                                ,N'".$value020."'
                                ,N'".$value021."'
                                ,N'".$value022."'
                                ,N'".$value023."'
                                ,N'".$value024."'
                                ,N'".$value025."'
                                ,N'".$value026."'
                                ,N'".$value027."'
                                ,N'".$value028."'
                                ,N'".$value029."'
                                ,N'".$value030."'
                                ,N'".$value031."'
                                ,N'".$value032."'
                                ,N'".$value033."'
                                ,N'".$value034."'
                                ,N'".$value035."'
                                ,N'".$value036."'
                                ,N'".$value037."'
                                ,N'".$value038."'
                                ,N'".$value039."'
                                ,N'".$value040."'
                                ,N'".$value041."'
                                ,N'".$value042."'
                                ,N'".$value043."'
                                ,N'".$value044."'
                                ,N'".$value045."'
                                ,N'".$value046."'
                                ,N'".$value047."'
                                ,N'".$value048."'
                                ,N'".$value049."'
                                ,N'".$value050."'
                                ,N'".$value051."'
                                ,N'".$value052."'
                                ,N'".$value053."'
                                ,N'".$value054."'
                                ,N'".$value055."'
                                ,N'".$value056."'
                                ,N'".$value057."'
                                ,N'".$value058."'
                                ,N'".$value059."'
                                ,N'".$value060."'
                                ,N'".$value061."'
                                ,N'".$value062."'
                                ,N'".$value063."'
                                ,N'".$value064."'
                                ,N'".$value065."'
                                ,N'".$value066."'
                                ,N'".$value067."'
                                ,N'".$value068."'
                                ,N'".$value069."'
                                ,N'".$value070."'
                                ,N'".$value071."'
                                ,N'".$value072."'
                                ,N'".$value073."'
                                ,N'".$value074."'
                                ,N'".$value075."'
                                ,N'".$value076."'
                                ,N'".$value077."'
                                ,N'".$value078."'
                                ,N'".$value079."'
                                ,N'".$value080."'
                                ,N'".$value081."'
                                ,N'".$value082."'
                                ,N'".$value083."'
                                ,N'".$value084."'
                                ,N'".$value085."'
                                ,N'".$value086."'
                                ,N'".$value087."'
                                ,N'".$value088."'
                                ,N'".$value089."'
                                ,N'".$value090."'
                                ,N'".$value091."'
                                ,N'".$value092."'
                                ,N'".$value093."'
                                ,N'".$value094."'
                                ,N'".$value095."'
                                ,N'".$value096."'
                                ,N'".$value097."'
                                ,N'".$value098."'
                                ,N'".$value099."'
                                ,N'".$value100."'
                                ,GETDATE()
                                ,'".$CreateID."'
                                )";  
                    }
                    else 
                    {
                         $query = " UPDATE [DBO].[ImExInfo100]

                                        SET   [ImportDate] = GETDATE()
                                             ,[ImportSequence] = '".$ImportSequence."'
                                             ,[ImportRows] = '".$ImportRows."'
                                             ,[Value001] = N'".$value001."'
                                             ,[Value002] = N'".$value002."'
                                             ,[Value003] = N'".$value003."'
                                             ,[Value004] = N'".$value004."'
                                             ,[Value005] = N'".$value005."'
                                             ,[Value006] = N'".$value006."'
                                             ,[Value007] = N'".$value007."'
                                             ,[Value008] = N'".$value008."'
                                             ,[Value009] = N'".$value009."'
                                             ,[Value010] = N'".$value010."'
                                             ,[Value011] = N'".$value011."'
                                             ,[Value012] = N'".$value012."'
                                             ,[Value013] = N'".$value013."'
                                             ,[Value014] = N'".$value014."'
                                             ,[Value015] = N'".$value015."'
                                             ,[Value016] = N'".$value016."'
                                             ,[Value017] = N'".$value017."'
                                             ,[Value018] = N'".$value018."'
                                             ,[Value019] = N'".$value019."'
                                             ,[Value020] = N'".$value020."'
                                             ,[Value021] = N'".$value021."'
                                             ,[Value022] = N'".$value022."'
                                             ,[Value023] = N'".$value023."'
                                             ,[Value024] = N'".$value024."'
                                             ,[Value025] = N'".$value025."'
                                             ,[Value026] = N'".$value026."'
                                             ,[Value027] = N'".$value027."'
                                             ,[Value028] = N'".$value028."'
                                             ,[Value029] = N'".$value029."'
                                             ,[Value030] = N'".$value030."'
                                             ,[Value031] = N'".$value031."'
                                             ,[Value032] = N'".$value032."'
                                             ,[Value033] = N'".$value033."'
                                             ,[Value034] = N'".$value034."'
                                             ,[Value035] = N'".$value035."'
                                             ,[Value036] = N'".$value036."'
                                             ,[Value037] = N'".$value037."'
                                             ,[Value038] = N'".$value038."'
                                             ,[Value039] = N'".$value039."'
                                             ,[Value040] = N'".$value040."'
                                             ,[Value041] = N'".$value041."'
                                             ,[Value042] = N'".$value042."'
                                             ,[Value043] = N'".$value043."'
                                             ,[Value044] = N'".$value044."'
                                             ,[Value045] = N'".$value045."'
                                             ,[Value046] = N'".$value046."'
                                             ,[Value047] = N'".$value047."'
                                             ,[Value048] = N'".$value048."'
                                             ,[Value049] = N'".$value049."'
                                             ,[Value050] = N'".$value050."'
                                             ,[Value051] = N'".$value051."'
                                             ,[Value052] = N'".$value052."'
                                             ,[Value053] = N'".$value053."'
                                             ,[Value054] = N'".$value054."'
                                             ,[Value055] = N'".$value055."'
                                             ,[Value056] = N'".$value056."'
                                             ,[Value057] = N'".$value057."'
                                             ,[Value058] = N'".$value058."'
                                             ,[Value059] = N'".$value059."'
                                             ,[Value060] = N'".$value060."'
                                             ,[Value061] = N'".$value061."'
                                             ,[Value062] = N'".$value062."'
                                             ,[Value063] = N'".$value063."'
                                             ,[Value064] = N'".$value064."'
                                             ,[Value065] = N'".$value065."'
                                             ,[Value066] = N'".$value066."'
                                             ,[Value067] = N'".$value067."'
                                             ,[Value068] = N'".$value068."'
                                             ,[Value069] = N'".$value069."'
                                             ,[Value070] = N'".$value070."'
                                             ,[Value071] = N'".$value071."'
                                             ,[Value072] = N'".$value072."'
                                             ,[Value073] = N'".$value073."'
                                             ,[Value074] = N'".$value074."'
                                             ,[Value075] = N'".$value075."'
                                             ,[Value076] = N'".$value076."'
                                             ,[Value077] = N'".$value077."'
                                             ,[Value078] = N'".$value078."'
                                             ,[Value079] = N'".$value079."'
                                             ,[Value080] = N'".$value080."'
                                             ,[Value081] = N'".$value081."'
                                             ,[Value082] = N'".$value082."'
                                             ,[Value083] = N'".$value083."'
                                             ,[Value084] = N'".$value084."'
                                             ,[Value085] = N'".$value085."'
                                             ,[Value086] = N'".$value086."'
                                             ,[Value087] = N'".$value087."'
                                             ,[Value088] = N'".$value088."'
                                             ,[Value089] = N'".$value089."'
                                             ,[Value090] = N'".$value090."'
                                             ,[Value091] = N'".$value091."'
                                             ,[Value092] = N'".$value092."'
                                             ,[Value093] = N'".$value093."'
                                             ,[Value094] = N'".$value094."'
                                             ,[Value095] = N'".$value095."'
                                             ,[Value096] = N'".$value096."'
                                             ,[Value097] = N'".$value097."'
                                             ,[Value098] = N'".$value098."'
                                             ,[Value099] = N'".$value099."'
                                             ,[Value100] = N'".$value100."'
                                             ,[LastUpdateDT] = GETDATE()
                                             ,[LastUpdateID] = '".$CreateID."'
                                        WHERE Value005 = N'".$value005."'
                                        ";
                    }

                    $stmt = sqlsrv_query($conn, $query);
                }

              

          } 
          } 

          if(!$break){
               
          //   echo '<script type="text/javascript">';
          //   echo ' alert("อัพโหลดไฟล์เรียบร้อย")'; 
          //   echo '</script>';
            echo "<script> swal(\"\",\" อัพโหลดไฟล์เรียบร้อย\", \"success\"
            ,         {
               closeOnClickOutside: false, 
               buttons: \"Download\", 
               
          }
            
            )
            .then(willRedirect => {
               if (willRedirect) {
                 window.open(\"export.php?JobID=$pJobID&Option=$pOption&SeqID=$pJobSeq\",\"_blank\");
                 window.location.reload();
               }
             })
            ; </script>";

          //    echo '<script type="text/javascript">';
          //    echo 'window.open("export.php?JobID='.$pJobID.'&Option='.$pOption.'&SeqID='.$pJobSeq.'","_blank")';
          //    echo '</script>';

          //   echo '<script type="text/javascript">';
          //   echo 'window.location.href = "index.php" ';
          //   echo '</script>';

          }

          else
          {

               $sqlDelete = "DELETE FROM [FPG_MIS_THAI_UAT].[DBO].[ImExInfo100]
               WHERE ImportSequence = $pJobSeq";
               
               sqlsrv_query($conn,$sqlDelete);

          
               // echo '<script type="text/javascript">';
               // echo ' alert("ไ่มได้อัพโหลดไฟล์")'; 
               // echo '</script>';
               // echo "<script> swal(\"\",\" ไ่มได้อัพโหลดไฟล์\", \"success\"); </script>";
               
          }
            
            
      }  
      else  
      {  
     //    echo '<script type="text/javascript">';
     //    echo ' alert("ไฟล์ไม่ถูกต้อง !")';  //not showing an alert box.
     //    echo '</script>';
           echo "<script> swal(\"\",\" ชนิดของไฟล์ไม่ถูกต้อง ! \", \"warning\")
           .then(willRedirect => {
               if (willRedirect) {
                 window.location.reload();
               }
             })
           ; </script>";
      }  
 }  
 ?>  
