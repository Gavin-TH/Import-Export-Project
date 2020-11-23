<?php
function findNum($str){

    $num = '';
    for ($i = 0; $i < strlen($str); $i++) {
  
        if (is_numeric($str[$i])) {
            $num .= $str[$i];
        }
    }
  
    return $num;
  
  }

  echo  ("FPGD00251/0163");
  echo  "---\n";
  echo  findNum("FPGD00251/0163");
  echo  "---";
  echo $string = substr(findNum("FPGD00251/0163"), 0, -4);
  echo  "---";
  echo $string - 5;

  if($string == 246)
  {
    echo  "TRUE";
  }
  else{
    echo  "FALSE";
  }
?>
