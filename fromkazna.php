<?
function insert_data ()
{
    global $odbc_access, $Workbook, $sheet1;

$del_query = "delete * from balans_kazna";

$res_del=odbc_exec($odbc_access, $del_query);

$i=3;
$excel_result_balacc = '0000';
while ($excel_result_balacc !='') 
	{
		$coord_razom = "C" . $i;
        $coord_balacc = "B" . $i;
        $coord_kazna_deb = "N" . $i;
        $coord_kazna_kred = "O" . $i;

        $Worksheet = $Workbook->Worksheets($sheet1);
        $Worksheet->activate;

        $excel_cell_razom = $Worksheet->Range($coord_razom);
        $excel_cell_razom->activate;
        $excel_result_razom = $excel_cell_razom->value;


        $excel_cell_balacc = $Worksheet->Range($coord_balacc);
        $excel_cell_balacc->activate;
        $excel_result_balacc = $excel_cell_balacc->value;

        $excel_cell_kazna_deb = $Worksheet->Range($coord_kazna_deb);
        $excel_cell_kazna_deb->activate;
        $excel_result_kazna_deb = $excel_cell_kazna_deb->value;

$kazna_deb_p = explode(".", $excel_result_kazna_deb);
$kazna_deb = implode(",", $kazna_deb_p);


        $excel_cell_kazna_kred = $Worksheet->Range($coord_kazna_kred);
        $excel_cell_kazna_kred->activate;
        $excel_result_kazna_kred = $excel_cell_kazna_kred->value;

$kazna_kred_p = explode(".", $excel_result_kazna_kred);
$kazna_kred = implode(",", $kazna_kred_p);

print $excel_result_razom.":::".$excel_cell_balacc. "\n";

$ins_query = "insert into balans_kazna (balacc, fond, kazna_deb, kazna_kred) values(".$excel_result_balacc.",'".$excel_result_razom."','".$kazna_deb."','".$kazna_kred."')";

$res=odbc_exec($odbc_access, $ins_query);

        $i = $i + 1;
    } 
} 

echo $odbc_access=odbc_connect("balans","balans","balans");


// открытие XLS файла

$filename = "D:/balans/bal.xls";
$sheet1 = "bal";

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