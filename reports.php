<?php
ini_set('error_reporting', E_ALL);
// Подключаем класс для работы с excel
require_once('/var/www/html/Server/PHPExcel.php');
// Подключаем класс для вывода данных в формате excel
require_once('/var/www/html/Server/PHPExcel/Writer/Excel2007.php');
// Подключаем конфиг для подключения к базе
require_once('/var/www/html/Server/config.php');
// Подключаем модуль для рендеринга пдф
$rendererName = PHPExcel_Settings::PDF_RENDERER_TCPDF;
$rendererLibrary = '/var/www/html/Server/tcpdf';
$rendererLibraryPath = '' . $rendererLibrary;


if (!PHPExcel_Settings::setPdfRenderer(
        $rendererName,
        $rendererLibraryPath
    )) {
    die(
        'Пожалуйста задайте значения $rendererName и $rendererLibraryPath' 
    );
}

// подключаемся к базе

  $conn = pg_connect("host=localhost port=5432 dbname=".$dbname." user=".$dbuser." password=".$dbpass."");
if (!$conn) {
  echo "Произошла ошибка соединения.\n";
  exit;
}

$result = pg_query($conn, "SELECT * from public.american_football_fumbles_stats");
if (!$result) {
  echo "Произошла ошибка при загрузке данных.\n";
	if (!$result) {
  	echo "ошибочный запрос\n";
  	exit;
	}

  exit;
}


$arr = pg_fetch_all($result);

// Создаем объект класса PHPExcel
$xls = new PHPExcel();
// Устанавливаем индекс активного листа
$xls->setActiveSheetIndex(0);
// Получаем активный лист
$sheet = $xls->getActiveSheet();
//ориентация альбомная
$xls->getActiveSheet()
    ->getPageSetup()
    ->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);


//отступы
$xls->getActiveSheet()
    ->getPageMargins()->setTop(0.50);
$xls->getActiveSheet()
    ->getPageMargins()->setRight(0.66);
$xls->getActiveSheet()
    ->getPageMargins()->setLeft(0.28);
$xls->getActiveSheet()
    ->getPageMargins()->setBottom(0.50);





// Подписываем лист
$sheet->setTitle('Тестовая таблица');

$xls->getActiveSheet()->fromArray($arr, null, 'A1');




      $xls->getActiveSheet()->getStyle('A1')->getBorders()->applyFromArray(
              array(
                  'bottom'     => array(
                      'style' => PHPExcel_Style_Border::BORDER_SLANTDASHDOT,
                      'color' => array(
                          'rgb' => '808080'
                      )
                  ),
                  'top'     => array(
                      'style' => PHPExcel_Style_Border::BORDER_SLANTDASHDOT,
                      'color' => array(
                          'rgb' => '808080'
                      )
                  )
              )
      );

// Вставляем массив в объект
$xls->getActiveSheet()->fromArray($arr, null, 'A1');






$objWriter = new PHPExcel_Writer_Excel5($xls);

$objWriter->save('/var/www/html/tmp/file1.xls');


$objPHPExcelWriter = PHPExcel_IOFactory::createWriter($xls,'HTML');
//print"$objPHPExcelWriter";

$objPHPExcelWriter->save('/var/www/html/tmp/file1.html');


// Создайте Writer для PDF, который автоматически подберет настройки, которые вы определили
$objWriter = PHPExcel_IOFactory::createWriter($xls, 'PDF');
//print "<pre>";
//print_r($xls);
//print "</pre>";

//$objWriter->SetLineStyle(array('width' => 0.5, 'cap' => 'butt', 'join' => 'miter', 'dash' => 4, 'color' => array(255, 0, 0)));
// Сохранить PDF в файл
$objWriter->save('/var/www/html/tmp/file1.pdf');

$objWorksheet = $xls->setActiveSheetIndex(0);
$highest_row = $objWorksheet->getHighestRow();
$highest_col = $objWorksheet->getHighestColumn();

//$highest_col_index = PHPExcel_Cell::columnIndexFromString($highest_col);
        
// start $row from 2, if you want to skip header

$newarr =[];

for ($counter = 1; $counter <= $highest_row; $counter++)
{
    $row = $objWorksheet->rangeToArray('A'.$counter.':'.$highest_col.$counter);


    $newarr = array_merge($newarr, $row); 
    $row = reset($row);            
}


function html_table($data = array())
{
    $rows = array();
    foreach ($data as $row) {
        $cells = array();
        foreach ($row as $cell) {
            $cells[] = "<td style='border-right-width:0.1px'>{$cell}</td>";
        }
        $rows[] = "<tr>" . implode('', $cells) . "</tr>";
    }
    return '<table border="1" cellspacing="3" cellpadding="4" style="border-right-width:1px">' . implode('', $rows) . '</table>';
}

class MYPDF extends TCPDF {

      public function Header() {
        // Logo
      //  $image_file = K_PATH_IMAGES.'logo_example.jpg';
     //   $this->Image($image_file, 10, 10, 15, '', 'JPG', '', 'T', false, 300, '', false, false, 0, false, false, false);
        // Set font

        $this->SetFont('freemono', 'B', 20);
        // Title
        $this->Cell(0, 15, 'Образец отчета', 0, false, 'C', 0, '', 0, false, 'M', 'M');
    }

    // Load table data from file
    public function LoadData($file) {
        // Read file lines
        $lines = file($file);
        $data = array();
        foreach($lines as $line) {
            $data[] = explode(';', chop($line));
        }
        return $data;
    }

    // Colored table
    public function ColoredTable($header,$data) {
        // Colors, line width and bold font
        $this->SetFillColor(255, 0, 0);
        $this->SetTextColor(255);
        $this->SetDrawColor(128, 0, 0);
        $this->SetLineWidth(0.3);
        $this->SetFont('', 'B');
        // Header
        $w = array(19, 19,19,19,19,19,19,19,19,19,19,19,19,19);

        $num_headers = count($header);
        for($i = 0; $i < $num_headers; ++$i) {
            $this->Cell($w[$i], 7, $header[$i], 1, 0, 'C', 1);
        //    $this->Cell($w[$i], 0, $header[$i], 1, 1, 'C', 0);
            
            //         Cell(0, 0, 'TEST CELL S, 1, 1, 'C', 0, '', 4);
            //MultiCell(55, 60, '[FIT CELL] '.$txt."\n", 1, 'J', 1, 1, 125, 145, true, 0, false, true, 60, 'M', true);
        }
        $this->Ln();
        // Color and font restoration
        $this->SetFillColor(224, 235, 255);
        $this->SetTextColor(0);
        $this->SetFont('');
        // Data
        $fill = 0;
        foreach($data as $row) {

                  foreach ($row as $cell) {
             $this->Cell($w[0], 0, $cell, 'LR', 0, 'L', $fill);


        }

          //  $this->Cell($w[0], 6, $row[0], 'LR', 0, 'L', $fill);
          //  $this->Cell($w[1], 6, $row[1], 'LR', 0, 'L', $fill);
          //  $this->Cell($w[2], 6, number_format($row[2]), 'LR', 0, 'R', $fill);
          //  $this->Cell($w[3], 6, number_format($row[3]), 'LR', 0, 'R', $fill);
            $this->Ln();
            $fill=!$fill;
        }
        $this->Cell(array_sum($w), 0, '', 'T');
    }
}







$html=  html_table($newarr) ;


$pdf = new MYPDF(PDF_PAGE_ORIENTATION, PDF_UNIT, PDF_PAGE_FORMAT, true, 'UTF-8', false);

// set document information
$pdf->SetCreator(PDF_CREATOR);
$pdf->SetAuthor('АСУДТ');
$pdf->SetTitle('Отчёт...');
$pdf->SetSubject('TCPDF Tutorial');
$pdf->SetKeywords('TCPDF, PDF, example, test, guide');

$pdf->setHeaderFont(array('freemono', '', 20));




// set default header data
$pdf->SetHeaderData(PDF_HEADER_LOGO, PDF_HEADER_LOGO_WIDTH, 'fasdfываыва'.' 011', 'ывафыва');

// set header and footer fonts
$pdf->setHeaderFont(Array(PDF_FONT_NAME_MAIN, '', PDF_FONT_SIZE_MAIN));
$pdf->setFooterFont(Array(PDF_FONT_NAME_DATA, '', PDF_FONT_SIZE_DATA));

// set default monospaced font
$pdf->SetDefaultMonospacedFont(PDF_FONT_MONOSPACED);

// set margins
$pdf->SetMargins(PDF_MARGIN_LEFT, PDF_MARGIN_TOP, PDF_MARGIN_RIGHT);
$pdf->SetHeaderMargin(PDF_MARGIN_HEADER);
$pdf->SetFooterMargin(PDF_MARGIN_FOOTER);

// set auto page breaks
$pdf->SetAutoPageBreak(TRUE, PDF_MARGIN_BOTTOM);

// set image scale factor
$pdf->setImageScale(PDF_IMAGE_SCALE_RATIO);

// set some language-dependent strings (optional)
if (@file_exists(dirname(__FILE__).'/lang/rus.php')) {
    require_once(dirname(__FILE__).'/lang/rus.php');
    $pdf->setLanguageArray($l);
}

// ---------------------------------------------------------

// set font
$pdf->SetFont('freemono', '', 12);

// add a page

$pdf->AddPage('L', 'A4');

// column titles
$header = array('столбец', 'столбец','столбец','столбец','столбец','столб','столбец','столбец','столбец','столбец','столбец','столбец','столбец','столбец',);


// print colored table
$pdf->ColoredTable($header, $newarr);

// ---------------------------------------------------------

// close and output PDF document
$pdf->Output('example_011.pdf', 'I');

//print "<pre>";
//print_r ($newarr);
//print "</pre>";




?>