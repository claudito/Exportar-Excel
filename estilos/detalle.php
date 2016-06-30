<?php 


function detalle()

{

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


return $estiloInformacion;


}



 ?>