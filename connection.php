<?php

//Connect MSSQL
$dbServer  = '10.62.2.24';
$dbUser  = 'sa';
$dbPassword  = 'P@ssw0rd123';
$dbName = 'FPG_MIS_THAI_UAT';

$connectionInfo = array(
    "Database" => $dbName, 
    "CharacterSet" => "UTF-8",
    "UID" => $dbUser,
    "PWD" => $dbPassword

);
 
$conn = sqlsrv_connect($dbServer, $connectionInfo);

?>