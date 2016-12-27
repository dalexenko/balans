<?
function insert_data ()
{
    global $odbc_access, $Workbook, $sheet1;

$del_query = "delete * from balans_bars";

$res_del=odbc_exec($odbc_access, $del_query);

$i=5;
$excel_result_balacc = '0000';

while ($excel_result_balacc !='') 
{
		$coord_razom = "C" . $i;
        $coord_balacc = "B" . $i;
        $coord_bars_deb = "E" . $i;
        $coord_bars_kred = "F" . $i;

        $Worksheet = $Workbook->Worksheets($sheet1);
        $Worksheet->activate;

        $excel_cell_razom = $Worksheet->Range($coord_razom);
        $excel_cell_razom->activate;
        $excel_result_razom = $excel_cell_razom->value;


        $excel_cell_balacc = $Worksheet->Range($coord_balacc);
        $excel_cell_balacc->activate;
        $excel_result_balacc = $excel_cell_balacc->value;

        $excel_cell_bars_deb = $Worksheet->Range($coord_bars_deb);
        $excel_cell_bars_deb->activate;
        $excel_result_bars_deb = $excel_cell_bars_deb->value;

$bars_deb_p = explode(".", $excel_result_bars_deb);
$bars_deb = implode(",", $bars_deb_p);


        $excel_cell_bars_kred = $Worksheet->Range($coord_bars_kred);
        $excel_cell_bars_kred->activate;
        $excel_result_bars_kred = $excel_cell_bars_kred->value;

$bars_kred_p = explode(".", $excel_result_bars_kred);
$bars_kred = implode(",", $bars_kred_p);

//print $excel_result_razom.":::".$excel_cell_balacc. "\n";

$ins_query = "insert into balans_bars (balacc, bars_deb, bars_kred) values(".$excel_result_balacc.",'".$bars_deb."','".$bars_kred."')";


echo "\n";


$res=odbc_exec($odbc_access, $ins_query);

        $i = $i + 1;
    } 
} 

echo $odbc_access=odbc_connect("balans","balans","balans");


// открытие XLS файла

$filename = "D:/balans/bars.xls";
$sheet1 = "Лист1";

$excel_app = new COM("Excel.application") or Die ("Did not connect");

$excel_app->Visible = 1;

$Workbook = $excel_app->Workbooks->Open("$filename") or Die("Did not open $filename $Workbook");

// чтение XLS файла

insert_data ();

// closing excel

$excel_app->Quit();

// free the object
//$excel_app->Release();

$excel_app = null;

?>