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

use PHPUnit\Framework\TestCase;

abstract class ConfiguredExcelTest extends TestCase {

	abstract public function getImplementation(): string;

	public function testSetters() {
		if($this->getImplementation() === ExcelFactory::TYPE_SPOUT) {
			$this->markTestSkipped('Spout doesn\'t support metadata');
			return;
		}
		$excel = ExcelFactory::createConfiguredExcel($this->getImplementation());

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
		$excel = ExcelFactory::createConfiguredExcel($this->getImplementation());

		$this->assertCount(0, $excel->getSheets());
		if($this->getImplementation() === ExcelFactory::TYPE_PHP_SPREADSHEET) {
			$this->assertCount(1, $excel->getPhpExcel()->getAllSheets());
		}

		$sheet = $excel->addSheet('TestSheet');
		$this->assertCount(1, $excel->getSheets());
		if($this->getImplementation() === ExcelFactory::TYPE_PHP_SPREADSHEET) {
			$this->assertCount(1, $excel->getPhpExcel()->getAllSheets());
			$this->assertEquals('TestSheet', $sheet->getSheet()->getTitle());
		}

		$sheet = $excel->addSheet('TestSheet2');
		$this->assertCount(2, $excel->getSheets());
		if($this->getImplementation() === ExcelFactory::TYPE_PHP_SPREADSHEET) {
			$this->assertCount(2, $excel->getPhpExcel()->getAllSheets());
			$this->assertEquals('TestSheet2', $sheet->getSheet()->getTitle());
		}
	}
}