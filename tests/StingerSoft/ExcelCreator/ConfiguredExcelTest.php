<?php
declare(strict_types=1);

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

	public function testSetters(): void {
		if($this->getImplementation() === ExcelFactory::TYPE_SPOUT) {
			// Spout doesn't support metadata
			self::assertTrue(true);
			return;
		}
		$excel = ExcelFactory::createConfiguredExcel($this->getImplementation());

		$excel->setTitle('TestTitle');
		self::assertEquals('TestTitle', $excel->getPhpExcel()->getProperties()->getTitle());
		self::assertEquals('TestTitle', $excel->getTitle());

		$excel->setCompany('TestCompany');
		self::assertEquals('TestCompany', $excel->getPhpExcel()->getProperties()->getCompany());
		self::assertEquals('TestCompany', $excel->getCompany());

		$excel->setCreator('TestCreator');
		self::assertEquals('TestCreator', $excel->getPhpExcel()->getProperties()->getCreator());
		self::assertEquals('TestCreator', $excel->getCreator());
	}

	public function testAddSheet(): void {
		$excel = ExcelFactory::createConfiguredExcel($this->getImplementation());

		self::assertCount(0, $excel->getSheets());
		if($this->getImplementation() === ExcelFactory::TYPE_PHP_SPREADSHEET) {
			self::assertCount(1, $excel->getPhpExcel()->getAllSheets());
		}

		$sheet = $excel->addSheet('TestSheet');
		self::assertCount(1, $excel->getSheets());
		if($this->getImplementation() === ExcelFactory::TYPE_PHP_SPREADSHEET) {
			self::assertCount(1, $excel->getPhpExcel()->getAllSheets());
			self::assertEquals('TestSheet', $sheet->getSheet()->getTitle());
		}

		$sheet = $excel->addSheet('TestSheet2');
		self::assertCount(2, $excel->getSheets());
		if($this->getImplementation() === ExcelFactory::TYPE_PHP_SPREADSHEET) {
			self::assertCount(2, $excel->getPhpExcel()->getAllSheets());
			self::assertEquals('TestSheet2', $sheet->getSheet()->getTitle());
		}
	}
}