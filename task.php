<?php
if($_POST && $_FILES){
    require_once "PHPExcel/Classes/PHPExcel.php";
    date_default_timezone_set("Asia/Kolkata");
    $recv_code = isset($_POST['recv_code']) ? $_POST['recv_code'] : '';
    $callsign_code = isset($_POST['callsign_code']) ? $_POST['callsign_code'] : '';
    $path = $_FILES['my_file_input']['tmp_name'];
    $reader = PHPExcel_IOFactory::createReaderForFile($path);
    $sheetnames = $reader->listWorksheetNames($path);
    $objPHPExcel = $reader->load($path);
    $getResultData  = '';
    $contcount = 0;
    $line = 0;
    foreach ($objPHPExcel->getWorksheetIterator() as $worksheet) {
        $worksheet = $worksheet->toArray();
        $dt = date('D M d Y H:i:s O');
        $refno = get_date_str($dt, "");
        $getResultData .= "UNB+UNOA:2+KMT+".$recv_code."+".get_date_str($dt, "daterawonly").":".get_date_str($dt, "timetominrawonly")."+".$refno."'\n";
        $getResultData .= "UNH+".$refno."+COPRAR:D:00B:UN:SMDG21+LOADINGCOPRAR'\n";
        $line++;
        $report_dt = ""; 
        $voyage = "";
        $vslname = "";
        $callsign = "";
        $opr = "";
        for ($singleRow = 0; $singleRow < count($worksheet); $singleRow++) {
            if($singleRow>6) break;
            $rowCells = $worksheet[$singleRow];
            if($singleRow==1) {
                $tmpdt = explode("/", $rowCells[1]);
                $day = $tmpdt[0];
                $month = $tmpdt[1];
                $tmpyear = explode(" ", $tmpdt[2]);
                $report_date = date($tmpyear[0]."-".$month."-".$day." " .$tmpyear[1]);
                $report_dt = get_date_str($report_date, "");
            }
            if($singleRow==3) {
                if(gettype($rowCells[3])!="undefined") {
                    $tmp = explode("/", $rowCells[3]);
                    $voyage = $tmp[0];
                    $callsign = $tmp[1];
                    $opr = $tmp[2];
                    $vslname = $rowCells[1];
                }
            }
        }  
        $getResultData .= "BGM+45+".$report_dt."+5'\n";
        $line++;
        $getResultData .= "TDT+20+".$voyage."+1++172:".$opr."+++".$callsign_code.":103::".$vslname."'\n";
        $line++;
        $getResultData .= "RFF+VON:".$voyage."'\n";
        $line++;
        $getResultData .= "NAD+CA+".$opr."'\n";
        $line++;
        $tmp;
        $dim;
        for ($singleRow = 0; $singleRow < count($worksheet); $singleRow++) {
            if(gettype($worksheet[$singleRow]) !="undefined") {
                $rowCells = $worksheet[$singleRow];
                if($singleRow>7) {
                    $contcount++;
                    //print_r($rowCells);die;
                    $fe = "5";
                    if(gettype($rowCells[3]) !="undefined" && $rowCells[3]=="E"){
                        $fe = "4";
                    }
                    $type = "2";
                    if(gettype($rowCells[11])!="undefined" && $rowCells[11]=="Y"){
                        $type = "6";
                    }
                    if(gettype($rowCells[1])!="undefined" && gettype($rowCells[7])!="undefined") { 
                        $getResultData .= "EQD+CN+".$rowCells[1]."+".$rowCells[7].":102:5++".$type."+".$fe."'\n";
                        $line++;
                    }
                    if(gettype($rowCells[6])!="undefined") { 
                        $getResultData .= "LOC+11+".$rowCells[5].":139:6'\n";
                        $line++;
                    }
                    if(gettype($rowCells[6])!="undefined") { 
                        $getResultData .= "LOC+7+".$rowCells[6].":139:6'\n";
                        $line++;
                    }
                    if(gettype($rowCells[19])!="undefined") { 
                        $getResultData .= "LOC+9+".$rowCells[19].":139:6'\n";
                        $line++;
                    }
                    if(gettype($rowCells[13])!="undefined") { 
                        $getResultData .= "MEA+AAE+VGM+KGM:".$rowCells[13]."'\n";
                        $line++;
                    }
                    if(gettype($rowCells[17])!="undefined" && trim($rowCells[17])!="" && trim($rowCells[17])!="/") {
                        $tmp = explode(",",$rowCells[17]);
                        for($i=0; $i<count($tmp); $i++) {
                            $dim = explode("/",$rowCells[17]);
                            if(trim($dim[0])=="OF") {
                                $getResultData .= "DIM+5+CMT:".trim($dim[1])."'\n";
                                $line++;
                            }
                            if(trim($dim[0])=="OB") {
                                $getResultData .= "DIM+6+CMT:".trim($dim[1])."'\n";
                                $line++;
                            }
                            if(trim($dim[0])=="OR") {
                                $getResultData .= "DIM+7+CMT::".trim($dim[1])."'\n";
                                $line++;
                            }
                            if(trim($dim[0])=="OL") {
                                $getResultData .= "DIM+8+CMT::".trim($dim[1])."'\n";
                                $line++;
                            }
                            if(trim($dim[0])=="OH") {
                                $getResultData .= "DIM+9+CMT:::".trim($dim[1])."'\n";
                                $line++;
                            }
                        }
                    }
                    if(gettype($rowCells[15])!="undefined" && trim($rowCells[15])!="" && trim($rowCells[15])!="/") {
                        $temperature = $rowCells[15];
                        $temperature = str_replace(' ','', $temperature);
                        $temperature = str_replace("C", "", $temperature);
                        $temperature = str_replace("+", "", $temperature);
                        $getResultData .= "TMP+2+".$temperature.":CEL'\n";
                        $line++;
                    }
                    if(gettype($rowCells[25])!="undefined" && trim($rowCells[25])!="" && trim($rowCells[25])!="/") {
                        $tmp = explode(",",$rowCells[25]);
                        if($tmp[0]=="L") {
                            $getResultData .= "SEL+".$tmp[1]."+CA'\n";
                            $line++;
                        }
                        if($tmp[0]=="S") {
                            $getResultData .= "SEL+".$tmp[1]."+SH'\n";
                            $line++;
                        }
                        if($tmp[0]=="M") {
                            $getResultData .= "SEL+".$tmp[1]."+CU'\n";
                            $line++;
                        }
                    }
                    if(gettype($rowCells[8])!="undefined") { 
                        $getResultData .= "FTX+AAI+++".$rowCells[8]."'\n";
                        $line++;
                    }                      
                            
                    if(gettype($rowCells[12])!="undefined" && trim($rowCells[12])!="" && trim($rowCells[12])!="/") {
                        $getResultData .= "FTX+AAA+++".trim(cleanString($rowCells[12]))."'\n";
                        $line++;
                    }
                    if(gettype($rowCells[18])!="undefined" && trim($rowCells[18])!="" && trim($rowCells[18])!="/") {
                        $getResultData .= "FTX+HAN++".$rowCells[18]."'\n";
                        $line++;
                    }
                    if(gettype($rowCells[14])!="undefined" && $rowCells[14]!="" && trim($rowCells[14])!="/") {
                        $tmp = $rowCells[14].split('/');
                        $getResultData .= "DGS+IMD+".$tmp[0]."+".$tmp[1]."'\n";
                        $line++;
                    }
                    if(gettype($rowCells[2])!="undefined" && trim($rowCells[2])!="") { 
                        $getResultData .= "NAD+CF+".$rowCells[2].":160:ZZZ'\n";
                        $line++;
                    } //box 
                }
            }
        }
        //$contcount--;
        $getResultData .= "CNT+16:".$contcount."'\n";
        $line++;
        $line++;
        $getResultData .= "UNT+".$line."+".$refno."'\n";
        $getResultData .= "UNZ+1+".$refno."'";
    }
    echo $getResultData;die;
}
function get_date_str($d, $type) {
    $now = $d;
    $dt = date("d", strtotime($now));
    $dt = (strlen($dt)<2)? "0".$dt : $dt;
    $hrs = date("H", strtotime($now));
    $hrs = (strlen($hrs)<2)? "0".$hrs : $hrs;
    $min = date("i", strtotime($now));
    $min = (strlen($min)<2)? "0".$min : $min;
    $sec = date("s", strtotime($now));
    $sec = (strlen($sec)<2)? "0".$sec : $sec;
    $mth = date("m", strtotime($now));
    $mth = (strlen($mth)<2)? "0".$mth : $mth;
    if($type=="daterawonly") {
        return date("Y", strtotime($now)).''.$mth.''.$dt;
    } else if($type=="timetominrawonly"){
        return $hrs.''.$min;
    } else {
        return date("Y", strtotime($now)).''.$mth.''.$dt.''.$hrs.''.$min.''.$sec;
    }
    //return now.getHours()+':'+String(min)+':'+String(sec);
}
function cleanString($input) {
    $output = "";
    for ($i=0; $i<strlen($input); $i++) {
        $utf16 = mb_convert_encoding($input, 'UTF-16LE', 'UTF-8');
        $charCodeAt = ord($utf16[$i * 2]) + (ord($utf16[$i * 2 + 1]) << 8);
        if ($charCodeAt <= 127) {
            $output .= $input[$i];
        }
    }
    return $output;
}
?>

<head>
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
    <script src="https://code.jquery.com/jquery-3.1.1.min.js"></script>
</head>
<body> 
    <div class="container">
        <div class="card" style="">
            <div class="card-body">
                <form id="data" method="post" enctype="multipart/form-data">
                    <h5 class="card-title">Export Booking Excel to Coprar Converter</h5>
                    <div class="form-group">
                        <label for="recv_code">Receiver Code:</label><input class="form-control" type="text" name="recv_code" id="recv_code" value="RECEIVER" />
                        <p><small>Please change before file select.</small></p>
                    </div>
                    <div class="form-group">
                        <label for="recv_code">Callsign Code:</label><input class="form-control" type="text" name="callsign_code" id="callsign_code" value="XXXXX" />
                        <p><small>Please change before file select.</small></p>
                    </div>
                    <div class="form-group">
                        <label for="my_file_input">Export booking excel file:</label><input class="form-control" type="file" name="my_file_input" id="my_file_input" />
                        <p><small><a href="https://westports.github.io/ETP/sample.xlsx">Sample Excel</a></small></p>
                    </div>
                    <div class="form-group"><textarea class="form-control" rows="20" cols="40" id='my_file_output'></textarea></div>
                </form>
            </div>
        </div>
    </div>
</body>
<script>

$(document).ready(function(){
    console.log(String(new Date().getDate()).length)
    $("#my_file_input").on('change', function(){
        var form = $('#data')[0];
        var formData = new FormData(form);
        $.ajax({
            url: 'task.php',
            type: 'POST',
            data: formData,
            success: function (data) {
               $("#my_file_output").val(data)
            },
            cache: false,
            contentType: false,
            processData: false
        });
    })
})

</script>