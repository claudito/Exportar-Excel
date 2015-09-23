<?php
$conexion = new mysqli('localhost','root','','control_de_transporte',3306);
if (mysqli_connect_errno()) {
printf("La conexión con el servidor de base de datos falló: %s\n", mysqli_connect_error());
exit();
}
$CODIGO=$_REQUEST[codigo];
$NOMBRE=$_REQUEST[nombre];



$consulta = "SELECT codigo_reporte,area,date_format(fecha,'%d-%m-%Y')as fecha, centro_costo, SUM( costo ) AS costo
FROM reporte_det
WHERE codigo_reporte =  '$CODIGO'
GROUP BY centro_costo";
$resultado = $conexion->query($consulta);
if($resultado->num_rows > 0 ){

date_default_timezone_set('America/Lima');

if (PHP_SAPI == 'cli')
die('Este archivo solo se puede ver desde un navegador web');

/** Se agrega la libreria PHPExcel */
require_once 'lib/PHPExcel/PHPExcel.php';

// Se crea el objeto PHPExcel
$objPHPExcel = new PHPExcel();

// Se asignan las propiedades del libro
$objPHPExcel->getProperties()->setCreator("LUIS CLAUDIO") //Autor
->setLastModifiedBy("LUIS CLAUDIO") //Ultimo usuario que lo modificó
->setTitle("Reporte Excel")
->setSubject("Reporte Excel")
->setDescription("Reporte de Transportista")
->setKeywords("reporte ade transportista")
->setCategory("Reporte excel");

$tituloReporte = "REPORTE N° ".$CODIGO." DEL TRANSPORTISTA".' '.$NOMBRE.' (DETALLE CC)';
$titulosColumnas=
array('CODIGO', 'ÁREA', 'CENTRO DE COSTO', 'COSTO','FECHA');

$objPHPExcel->setActiveSheetIndex(0)
->mergeCells('A1:Q1');

// Se agregan los titulos del reporte
$objPHPExcel->setActiveSheetIndex(0)
->setCellValue('A1',$tituloReporte)
->setCellValue('A3',  $titulosColumnas[0])
->setCellValue('B3',  $titulosColumnas[1])
->setCellValue('C3',  $titulosColumnas[2])
->setCellValue('D3',  $titulosColumnas[3])
->setCellValue('E3',  $titulosColumnas[4]);


//Se agregan los datos de los alumnos
$i = 4;
while ($fila = $resultado->fetch_array()) {
$objPHPExcel->setActiveSheetIndex(0)
->setCellValueExplicit('A'.$i,  $fila['codigo_reporte'],PHPExcel_Cell_DataType::TYPE_STRING)
->setCellValueExplicit('B'.$i,  $fila['area'],PHPExcel_Cell_DataType::TYPE_STRING)
->setCellValueExplicit('C'.$i,  $fila['centro_costo'],PHPExcel_Cell_DataType::TYPE_STRING)
->setCellValueExplicit('D'.$i,  $fila['costo'],PHPExcel_Cell_DataType::TYPE_NUMERIC)
->setCellValueExplicit('E'.$i,  $fila['fecha'],PHPExcel_Cell_DataType::TYPE_STRING);
$i++;
}

$estiloTituloReporte = array(
'font' => array(
'name'      => 'Verdana',
'bold'      => true,
'italic'    => false,
'strike'    => false,
'size' =>18,
'color'     => array(
'rgb' => 'FFFFFF'
)
),
'fill' => array(
'type'	=> PHPExcel_Style_Fill::FILL_SOLID,
'color'	=> array('argb' => '064D06')
),
'borders' => array(
'allborders' => array(
'style' => PHPExcel_Style_Border::BORDER_NONE                    
)
), 
'alignment' =>  array(
'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
'vertical'   => PHPExcel_Style_Alignment::VERTICAL_CENTER,
'rotation'   => 0,
'wrap'          => TRUE
)
);

$estiloTituloColumnas = array(
'font' => array(
'name'      => 'Arial',
'bold'      => true,
'size'     => 9  ,                       
'color'     => array(
'rgb' => 'FFFFFF'//color letra
)
),
'fill' 	=> array(
'type'		=> PHPExcel_Style_Fill::FILL_GRADIENT_LINEAR,
'rotation'   => 90,
'startcolor' => array(
'rgb' => '159415'//color cabecera reporte-inicio
),
'endcolor'   => array(
'argb' => '159415'//color cabecera reporte-fin
)
),
'borders' => array(
'top'     => array(
'style' => PHPExcel_Style_Border::BORDER_MEDIUM ,
'color' => array(
'rgb' => '143860'
)
),
'bottom'     => array(
'style' => PHPExcel_Style_Border::BORDER_MEDIUM ,
'color' => array(
'rgb' => '143860'
)
)
),
'alignment' =>  array(
'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
'vertical'   => PHPExcel_Style_Alignment::VERTICAL_CENTER,
'wrap'          => TRUE
));

$estiloInformacion = new PHPExcel_Style();
$estiloInformacion->applyFromArray(
array(
'font' => array(
'name'      => 'Arial',  
'size'     => 9  ,               
'color'     => array(
'rgb' => '000000'
)
),
'fill' 	=> array(
'type'		=> PHPExcel_Style_Fill::FILL_SOLID,
'color'		=> array('argb' => 'CBECCB')//color de datos
),
'borders' => array(
'left'     => array(
'style' => PHPExcel_Style_Border::BORDER_THIN ,
'color' => array(
'rgb' => '3a2a47'
)
)             
)
));

$objPHPExcel->getActiveSheet()->getStyle('A1:E1')->applyFromArray($estiloTituloReporte);
$objPHPExcel->getActiveSheet()->getStyle('A3:E3')->applyFromArray($estiloTituloColumnas);		
$objPHPExcel->getActiveSheet()->setSharedStyle($estiloInformacion, "A4:E".($i-1));

for($i = 'A'; $i <= 'G'; $i++){
$objPHPExcel->setActiveSheetIndex(0)			
->getColumnDimension($i)->setAutoSize(TRUE);
}

// Se asigna el nombre a la hoja
$objPHPExcel->getActiveSheet()->setTitle('TRANSPORTISTA');

// Se activa la hoja para que sea la que se muestre cuando el archivo se abre
$objPHPExcel->setActiveSheetIndex(0);
// Inmovilizar paneles 
//$objPHPExcel->getActiveSheet(0)->freezePane('A4');
$objPHPExcel->getActiveSheet(0)->freezePaneByColumnAndRow(0,4);

// Se manda el archivo al navegador web, con el nombre que se indica (Excel2007)
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="Reporte-Detalle-CC.xlsx"');
header('Cache-Control: max-age=0');

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save('php://output');
exit;

}
else{
print_r('No hay resultados para mostrar');
}
?>