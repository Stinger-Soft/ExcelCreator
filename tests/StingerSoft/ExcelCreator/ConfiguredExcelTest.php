<?php

/*
 * This file is part of the Stinger Excel Creator package.
 *
 * (c) Oliver Kotte <oliver.kotte@stinger-soft.net>
 * (c) Florian Meyer <florian.meyer@stinger-soft.net>
 *
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */
namespace StingerSoft\ExcelCreator;

class ConfiguredExcelTest extends \PHPUnit_Framework_TestCase {

	public function testSetters() {
		$excel = new ConfiguredExcel();
		
		$excel->setTitle('TestTitle');
		$this->assertEquals('TestTitle', $excel->getPhpExcel()->getProperties()->getTitle());
		$this->assertEquals('TestTitle', $excel->getTitle());
		
		$excel->setCompany('TestCompany');
		$this->assertEquals('TestCompany', $excel->getPhpExcel()->getProperties()->getCompany());
		$this->assertEquals('TestCompany', $excel->getCompany());
		
		$excel->setCreator('TestCreator');
		$this->assertEquals('TestCreator', $excel->getPhpExcel()->getProperties()->getCreator());
		$this->assertEquals('TestCreator', $excel->getCreator());
	}

	public function testAddSheet() {
		$excel = new ConfiguredExcel();
		$this->assertCount(0, $excel->getSheets());
		$this->assertCount(1, $excel->getPhpExcel()->getAllSheets());
		
		$sheet = $excel->addSheet('TestSheet');
		$this->assertCount(1, $excel->getSheets());
		$this->assertCount(1, $excel->getPhpExcel()->getAllSheets());
		$this->assertEquals('TestSheet', $sheet->getSheet()->getTitle());
		
		$sheet = $excel->addSheet('TestSheet2');
		$this->assertCount(2, $excel->getSheets());
		$this->assertCount(2, $excel->getPhpExcel()->getAllSheets());
		$this->assertEquals('TestSheet2', $sheet->getSheet()->getTitle());
	}
}