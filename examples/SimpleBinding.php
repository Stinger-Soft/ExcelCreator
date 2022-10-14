<?php

use StingerSoft\ExcelCreator\ExcelFactory;
use StingerSoft\ExcelCreator\ColumnBinding;

include __DIR__.'/../vendor/autoload.php';
include __DIR__.'/Person.php';

//Create excel file
$excel = ExcelFactory::createConfiguredExcel();

//and a first sheet called 'Party guests'
$sheet1 = $excel->addSheet('Party guests');

//Add a column with the caption 'name'
$nameBinding = new ColumnBinding();
$nameBinding->setLabel('Name');

//bind it to the property of the data to be rendered by specifying the property path (see http://symfony.com/doc/current/components/property_access.html )
$nameBinding->setBinding('name');

//Width will be calculated by excel
$nameBinding->setColumnWidth('auto');
$sheet1->addColumnBinding($nameBinding);


//Add a column with the caption 'E-Mail'
$emailBinding = new ColumnBinding();
$emailBinding->setLabel('E-Mail');
$emailBinding->setBinding('email');
$emailBinding->setColumnWidth('auto');
$sheet1->addColumnBinding($emailBinding);

$guests = array();
$guests[] = new Person('Peter Mobb', 'peter@mobbtrix.de');
$guests[] = new Person('Peter und Uschi', 'peter_uschi@meppen.de');

$sheet1->setData($guests);
$sheet1->applyData();

$excel->writeToFile(__DIR__.'/simple_binding.xlsx');