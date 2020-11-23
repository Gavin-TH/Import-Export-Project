<?php
include("nav.html");
include("connection.php");


?>
<!doctype html>

<script type="text/javascript">

function getFileID(){
    var fileID = document.getElementById("IndexID").value
    return fileID;
}

</script>



<head>
    <title>หน้าแรก</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
</head>
<body>  


    <div class="container box" style="margin-top:40px;" >  
        <h3 align="center">หน้าแรก</h3>  
       
    </div>  
    
    <div class="container text-center" style="margin-top:100px;">  
    <div class="row">
        <div class="col-md-3"></div>
        <div class="col-md-6 ">
                <div class="row">
                    
                    <div class="col-md-3" style="display: flex; justify-content:flex-end"> 
                    
                    <script>
                    function addDate(){
                    date = new Date();
                    var month = date.getMonth()+1;
                    var day = date.getDate();
                    var year = date.getFullYear();

                    if (document.getElementById('date').value == ''){
                    document.getElementById('date').value = "วันที่ " + day + '-' + month + '-' + year;
                    }
                    }
                    </script>
                      <form mehtod="post" id="export_excel">
                    <body onload="addDate();">
                        <input name="date" type="text" id="date" readonly style="text-align:center; width:150px;">
                    </body>
                    </div>

                    <div class="col-md-6">
                        <select name='JobID' class="custom-select">
                        
                        <?php 
                            

                            $query = "SELECT * FROM [ImExJobs]";
                            $stmt = sqlsrv_query($conn, $query);
                            $word = "ชื่องาน: ";

                            while($report = sqlsrv_fetch_array($stmt)){

                                echo "<option value='".$report['JobID']."'>   $word  ".$report['JobName_TH']."  </option>";
                            }
                        ?>
                        </select>
                    </div>

                    <div class="col-md-3" style="display: flex; justify-content: flex-start">
                        <select name='countRound' class="custom-select" style="width:110px;">
                            <?php
                            $query = "SELECT CASE WHEN MAX(ImportSequence) IS NULL THEN 1
                            ELSE MAX(ImportSequence) +1  END AS [ROUND]
                            FROM ImExInfo100";
                            
                            $stmt = sqlsrv_query($conn, $query);
                            $word = "รอบที่";
                            $Round = sqlsrv_fetch_array($stmt);
                        
                            echo "<option value='".$Round['ROUND']."'>   $word  ".$Round['ROUND']."  </option>";
                            ?>
                        </select>
                    </div>
                    
                </div>
        </div>
        <div class="col-md-3"></div>
        </div>
    </div>

    <div class="container text-center" style="margin-top:40px;">  
    <div class="row">

    <div class="col-md-3"></div>

    <div class="col-md-6">
    <div class="row">

    <div class="col-md-8">
      
            <div class="custom-file"> 
                <input type="file" name="excel_file" id="excel_file" class="custom-file-input"/>  
                <label id="excel_label" name="labelname" class="custom-file-label" for="excel_file">เลือกไฟล์</label>
            </div>
        
    </div>
    <div class="col-md-4"> 
   
        <input type="submit" value="อัพโหลดไฟล์" id="btn_submit" class="btn btn-lg btn-primary mr-2">
       
    </div>
    </div>
    </div>

    <div class="col-md-3"></div>
    <div id="result" style="text-align: center;">  
    </div>

    </div>
                             
    <div class="container box" style="margin-top:25px;">  
    <div class="form-check">
        
        <input class="form-check-input" type="radio" name="radio1" id="radio1" value="1" checked>
        <label class="form-check-label" for="exampleRadios1">รอบล่าสุด</label>
     
        <input class="form-check-input" type="radio" name="radio1" id="radio2" value="2" style="margin-left: 20px;">
        <label class="form-check-label" for="exampleRadios2"  style="margin-left: 40px;">เลือกทั้งหมด</label>
     
    </div>

    
        </form>
    </div>
    </div>


        

        <div class="modal fade" id="wait" tabindex="-1" role="dialog" aria-labelledby="exampleModalCenterTitle" aria-hidden="true" data-keyboard="false" data-backdrop="static">
          <div class="modal-dialog modal-dialog-centered" role="document">
            <div class="modal-content">
             
              <div class="modal-body">
            
                <center><p><img src="images/loading.gif" width="100%" height="100%"/></p></center>
               

                

              </div>
            </div>
          </div>
        </div>


</body>  

</html>
<script>  

var myApp;
myApp = myApp || (function () {
    var pleaseWaitDiv = $('#wait');
    return {
        showPleaseWait: function() {
            pleaseWaitDiv.modal();
        },
        hidePleaseWait: function () {
            pleaseWaitDiv.modal('hide');
        },

    };
})();


$(".custom-file-input").on("change", function() 
{
  var fileName = $(this).val().split("\\").pop();
  $(this).siblings(".custom-file-label").addClass("selected").html(fileName);

});



 $(document).ready(function()
 {  
      $('#btn_submit').click(function(){  
           //$('#export_excel').submit();  
           

      });  


      $('#export_excel').on('submit', function(event){  
      myApp.showPleaseWait();
           event.preventDefault();  
           $.ajax({  
                url:"import.php",  
                method:"POST",  
                
                data:new FormData(this),  
                contentType:false,  
                processData:false,  
                success:function(data){  
                     $('#result').html(data);  
                     $('#excel_file').val('');  
                     myApp.hidePleaseWait();

                }  
           });
         
            
      });  
 });  

 
 </script>  

 <script>


</script>
