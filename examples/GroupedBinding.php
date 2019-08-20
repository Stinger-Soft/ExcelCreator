<?php

use StingerSoft\ExcelCreator\ColumnBinding;
use StingerSoft\ExcelCreator\ExcelFactory;

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

//Add other information in outline 1
$streetBinding = new ColumnBinding();
$streetBinding->setLabel('Address');
$streetBinding->setBinding('address');
$streetBinding->setColumnWidth('auto');
$sheet1->addColumnBinding($streetBinding);

$cityBinding = new ColumnBinding();
$cityBinding->setLabel('City');
$cityBinding->setBinding('city');
$cityBinding->setColumnWidth('auto');
$sheet1->addColumnBinding($cityBinding);

$zipBinding = new ColumnBinding();
$zipBinding->setLabel('Zipcode');
$zipBinding->setBinding('zipcode');
$zipBinding->setColumnWidth('auto');
$sheet1->addColumnBinding($zipBinding);

//Group results by zip code
$sheet1->setGroupByBinding($zipBinding);

$guests = array();

$person = new Person('Peter Mobb', 'peter@mobbtrix.de');
$person->setAddress('Musterstraße 1');
$person->setCity('Bremen');
$person->setZipCode(28357);
$guests[] = $person;

$person = new Person('Oliver Kotte', 'oliver.kotte@stinger-soft.net');
$person->setAddress('Musterstraße 2');
$person->setCity('Bremen');
$person->setZipCode(28357);
$guests[] = $person;

$person = new Person('Florian Meyer', 'florian.meyer@stinger-soft.net');
$person->setAddress('Musterstraße 3');
$person->setCity('Bremen');
$person->setZipCode(28357);
$guests[] = $person;

$person = new Person('Peter und Uschi', 'peter_uschi@meppen.de');
$person->setAddress('Meppstraße 1');
$person->setCity('Meppen');
$person->setZipCode(49716);
$guests[] = $person;

$sheet1->setData($guests);
$sheet1->applyData();

$excel->writeToFile(__DIR__.'/grouped_binding.xlsx');