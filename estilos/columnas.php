<?php 


function columnas()

{

		$Columnas = array(
            'font' => array(
                'name'      => 'Arial',
                'bold'      => true,
                 'size'     => 8  ,                       
                'color'     => array(
                    'rgb' => 'FFFFFF'//color letra
                )
            ),
            'fill' 	=> array(
				'type'		=> PHPExcel_Style_Fill::FILL_GRADIENT_LINEAR,
				'rotation'   => 90,
        		'startcolor' => array(
            		'rgb' => '159415'//color cabecera 
        		),
        		'endcolor'   => array(
            		'argb' => '159415'//color cabecera 
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



return $Columnas;


}










 ?>