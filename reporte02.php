<?php

  include('configuracion.php');
  include('bd/conexion.php');
  $db = new Conexion();

	$query     = "SELECT * FROM alumno";
	$result    = $db->query($query);
	$numfilas  = $result->num_rows;
	if($numfilas > 0 )
	{
	 					
		if (PHP_SAPI == 'cli')
			die('Este archivo solo se puede ver desde un navegador web');

		/** Se agrega la libreria PHPExcel */
		include('librerias/PHPExcel/PHPExcel.php');

		// Se crea el objeto PHPExcel
		$objPHPExcel = new PHPExcel();

		// Se asignan las propiedades del libro
		$objPHPExcel->getProperties()->setCreator("LUIS CLAUDIO") //Autor
							 ->setLastModifiedBy("LUIS CLAUDIO") //Ultimo usuario que lo modificó
							 ->setTitle("Reporte Excel")
							 ->setSubject("Reporte Excel")
							 ->setDescription("Reporte de Alumnos")
							 ->setKeywords("Reporte de Alumnos")
							 ->setCategory("Reporte excel");

		$tituloReporte   = "REPORTE DE ALUMNO";
		$titulosColumnas = array('CÓDIGO', 'NOMBRES', 'APELLIDOS','EDAD','USUARIO','CONTRASEÑA');
        
        $objPHPExcel->setActiveSheetIndex(0)
        		    ->mergeCells('A1:F1');

						
		// Se agregan los titulos del reporte
		$objPHPExcel->setActiveSheetIndex(0)
					->setCellValue('A1',$tituloReporte)
        		    ->setCellValue('A3',  $titulosColumnas[0])
		            ->setCellValue('B3',  $titulosColumnas[1])
        		    ->setCellValue('C3',  $titulosColumnas[2])
        		    ->setCellValue('D3',  $titulosColumnas[3])
        		    ->setCellValue('E3',  $titulosColumnas[4])
        		    ->setCellValue('F3',  $titulosColumnas[5])
        		   
        		    ;
		
		
		//Se agregan los datos de los alumnos
		$i = 4;
		while ($fila = $result->fetch_array()) {
			$objPHPExcel->setActiveSheetIndex(0)
        		    ->setCellValueExplicit('A'.$i,  $fila['codigo'],PHPExcel_Cell_DataType::TYPE_NUMERIC)
		            ->setCellValueExplicit('B'.$i,  $fila['nombres'],PHPExcel_Cell_DataType::TYPE_STRING)
        		    ->setCellValueExplicit('C'.$i,  $fila['apellidos'],PHPExcel_Cell_DataType::TYPE_STRING)
        		    ->setCellValueExplicit('D'.$i,  $fila['edad'],PHPExcel_Cell_DataType::TYPE_NUMERIC)
        		    ->setCellValueExplicit('E'.$i,  $fila['usuario'],PHPExcel_Cell_DataType::TYPE_STRING)
        		    ->setCellValueExplicit('F'.$i,  $fila['contrasena'],PHPExcel_Cell_DataType::TYPE_STRING)
        		    

        		     ;
					$i++;
		}

       
       #estilos de la celda
       include('estilos/titulo.php');
       include('estilos/columnas.php');
       include('estilos/detalle.php');
       $estiloTituloReporte  = titulo();
       $estiloTituloColumnas = columnas();
       $estiloInformacion    = detalle();

		
		$objPHPExcel->getActiveSheet()->getStyle('A1:F1')->applyFromArray($estiloTituloReporte);
		$objPHPExcel->getActiveSheet()->getStyle('A3:F3')->applyFromArray($estiloTituloColumnas);
		$objPHPExcel->getActiveSheet()->setSharedStyle($estiloInformacion, "A4:F".($i-1));
				
		for($i = 'A'; $i <= 'F'; $i++){
			$objPHPExcel->setActiveSheetIndex(0)			
				->getColumnDimension($i)/*->setAutoSize(TRUE)*/;
		}
		
		// Se asigna el nombre a la hoja
		$objPHPExcel->getActiveSheet()->setTitle('ALUMNOS');
        
         // Inmovilizar paneles 
		$objPHPExcel->getActiveSheet(0)->freezePane('A4');
		$objPHPExcel->getActiveSheet(0)->freezePaneByColumnAndRow(0,4);



		// Se activa la hoja para que sea la que se muestre cuando el archivo se abre
		$objPHPExcel->setActiveSheetIndex(0);

		// Se manda el archivo al navegador web, con el nombre que se indica (Excel2007)
		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header('Content-Disposition: attachment;filename="REPORTE DE ALUMNOS-'.date('d-m-Y H:i:s').'.xlsx"');
		header('Cache-Control: max-age=0');

		$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
		$objWriter->save('php://output');
		exit;

		
	}
	else
	{
	echo "No hay resultados para mostrar";
	}
?>