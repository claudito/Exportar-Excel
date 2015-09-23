<?php
    $conexion = new mysqli('localhost','root','','control_de_transporte',3306);
	if (mysqli_connect_errno()) {
    	printf("La conexión con el servidor de base de datos falló: %s\n", mysqli_connect_error());
    	exit();
	}
	$CODIGO=$_REQUEST[codigo];
    $NOMBRE=$_REQUEST[nombre];
	$consulta = "SELECT * FROM reporte_det WHERE codigo_reporte='$CODIGO'";
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

		$tituloReporte = "REPORTE DE DE TRANSPORTISTA".' '.$NOMBRE;
		$titulosColumnas=
		 array('CODIGO', 'TRANSPORTISTA', 'FECHA', 'AREA','CC','PROVEEDOR/CLIENTE',
		 	'DISTRITO','DIRECCIÓN','TIPO','COSTO','G/R','OC-OS','O/T','COMENTARIO');
		
		$objPHPExcel->setActiveSheetIndex(0)
        		    ->mergeCells('A1:N1');
						
		// Se agregan los titulos del reporte
		$objPHPExcel->setActiveSheetIndex(0)
					->setCellValue('A1',$tituloReporte)
        		    ->setCellValue('A3',  $titulosColumnas[0])
		            ->setCellValue('B3',  $titulosColumnas[1])
        		    ->setCellValue('C3',  $titulosColumnas[2])
        		    ->setCellValue('D3',  $titulosColumnas[3])
            		->setCellValue('E3',  $titulosColumnas[4])
            		->setCellValue('F3',  $titulosColumnas[5])
            		->setCellValue('G3',  $titulosColumnas[6])
            		->setCellValue('H3',  $titulosColumnas[7])
            		->setCellValue('I3',  $titulosColumnas[8])
            		->setCellValue('J3',  $titulosColumnas[9])
            		->setCellValue('k3',  $titulosColumnas[10])
            		->setCellValue('L3',  $titulosColumnas[11])
            		->setCellValue('M3',  $titulosColumnas[12])
            		->setCellValue('N3',  $titulosColumnas[13]);
		
		
		//Se agregan los datos de los alumnos
		$i = 4;
		while ($fila = $resultado->fetch_array()) {
			$objPHPExcel->setActiveSheetIndex(0)
        		    ->setCellValueExplicit('A'.$i,  $fila['codigo_reporte'],PHPExcel_Cell_DataType::TYPE_STRING)
		            ->setCellValueExplicit('B'.$i,  $fila['transportista'],PHPExcel_Cell_DataType::TYPE_STRING)
        		    ->setCellValueExplicit('C'.$i,  $fila['fecha'],PHPExcel_Cell_DataType::TYPE_STRING)
        		     ->setCellValueExplicit('D'.$i,  $fila['area'],PHPExcel_Cell_DataType::TYPE_STRING)
            		->setCellValueExplicit('E'.$i, $fila['centro_costo'],PHPExcel_Cell_DataType::TYPE_STRING)
            		->setCellValueExplicit('F'.$i,  $fila['proveedor_cliente'],PHPExcel_Cell_DataType::TYPE_STRING)
            		->setCellValueExplicit('G'.$i,  $fila['distrito'],PHPExcel_Cell_DataType::TYPE_STRING)
				    ->setCellValueExplicit('H'.$i,  $fila['direccion'],PHPExcel_Cell_DataType::TYPE_STRING)
				    ->setCellValueExplicit('I'.$i,  $fila['tipo_unidad'],PHPExcel_Cell_DataType::TYPE_STRING)
				    ->setCellValueExplicit('J'.$i,  $fila['costo'],PHPExcel_Cell_DataType::TYPE_NUMERIC)
				    ->setCellValueExplicit('K'.$i,  $fila['guia'],PHPExcel_Cell_DataType::TYPE_STRING)
				    ->setCellValueExplicit('L'.$i,  $fila['oc_os'],PHPExcel_Cell_DataType::TYPE_STRING)
				    ->setCellValueExplicit('M'.$i,  $fila['ot'],PHPExcel_Cell_DataType::TYPE_STRING)
				    ->setCellValueExplicit('N'.$i,  $fila['comentario'],PHPExcel_Cell_DataType::TYPE_STRING);
					$i++;
		}
		
		$estiloTituloReporte = array(
        	'font' => array(
	        	'name'      => 'Verdana',
    	        'bold'      => true,
        	    'italic'    => false,
                'strike'    => false,
               	'size' =>16,
	            	'color'     => array(
    	            	'rgb' => 'FFFFFF'
        	       	)
            ),
	        'fill' => array(
				'type'	=> PHPExcel_Style_Fill::FILL_SOLID,
				'color'	=> array('argb' => 'FF220835')
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
                'color'     => array(
                    'rgb' => 'FFFFFF'
                )
            ),
            'fill' 	=> array(
				'type'		=> PHPExcel_Style_Fill::FILL_GRADIENT_LINEAR,
				'rotation'   => 90,
        		'startcolor' => array(
            		'rgb' => 'c47cf2'
        		),
        		'endcolor'   => array(
            		'argb' => 'FF431a5d'
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
               	'color'     => array(
                   	'rgb' => '000000'
               	)
           	),
           	'fill' 	=> array(
				'type'		=> PHPExcel_Style_Fill::FILL_SOLID,
				'color'		=> array('argb' => 'CBECCB')
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
		 
		$objPHPExcel->getActiveSheet()->getStyle('A1:N1')->applyFromArray($estiloTituloReporte);
		$objPHPExcel->getActiveSheet()->getStyle('A3:N3')->applyFromArray($estiloTituloColumnas);		
		$objPHPExcel->getActiveSheet()->setSharedStyle($estiloInformacion, "A4:N".($i-1));
				
		for($i = 'A'; $i <= 'N'; $i++){
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
		header('Content-Disposition: attachment;filename="Reportedetransportista.xlsx"');
		header('Cache-Control: max-age=0');

		$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
		$objWriter->save('php://output');
		exit;
		
	}
	else{
		print_r('No hay resultados para mostrar');
	}
?>