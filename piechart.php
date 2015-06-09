<?php


// include PHPExxcel class
include 'phpexcel186/classes/PHPExcel.php';

	$objReader = PHPExcel_IOFactory::createReader('Excel2007');
	$objPHPExcel = $objReader->load('pie.xlsx');

    $sheet = $objPHPExcel->getActiveSheet();

	$dataseriesLabels1 = array();
	 
	$xAxisTickValues1 = array(
	  new PHPExcel_Chart_DataSeriesValues('String', "'Pie'!\$A$3:\$A$6", NULL, 4),  
	);
	 
	$dataSeriesValues1 = array(
	  new PHPExcel_Chart_DataSeriesValues('String', "'Pie'!\$B$3:\$B$6", NULL, 4), 
	);
	 
	$series1 = new PHPExcel_Chart_DataSeries(
		PHPExcel_Chart_DataSeries::TYPE_PIECHART,       // plotType
		PHPExcel_Chart_DataSeries::GROUPING_STANDARD,     // plotGrouping
		range(0, count($dataSeriesValues1)-1),          // plotOrder
		$dataseriesLabels1,                   // plotLabel
		$xAxisTickValues1,                    // plotCategory
		$dataSeriesValues1                    // plotValues
	);
 
    //  Set up a layout object for the Pie chart
    $layout1 = new PHPExcel_Chart_Layout();
    $layout1->setShowVal(TRUE);
    $layout1->setShowPercent(TRUE);
 
    //  Set the series in the plot area
    $plotarea1 = new PHPExcel_Chart_PlotArea($layout1, array($series1));
    //  Set the chart legend
    $legend1 = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);

    $title1 = new PHPExcel_Chart_Title('PIE CHART');


    //  Create the chart
    $chart1 = new PHPExcel_Chart(
      'chart1',   // name
      $title1,    // title
      $legend1,   // legend
      $plotarea1,   // plotArea
      true,     // plotVisibleOnly
      0,        // displayBlanksAs
      NULL,     // xAxisLabel
      NULL      // yAxisLabel   - Pie charts don't have a Y-Axis
    );

    //  Set the position where the chart should appear in the worksheet
    $chart1->setTopLeftPosition('H2');
    $chart1->setBottomRightPosition('N21'); 
    $sheet->addChart($chart1);
	
	
	$writer = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
             
	$writer->setIncludeCharts(true);
	
	header("Content-type: application/x-msdownload");
    header("Content-Disposition: attachment; filename=PieChart.xlsx");
    $writer->save('php://output');
    exit;