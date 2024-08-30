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

abstract class ConfiguredExcelTestCase extends TestCase {

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
		} elseif($this->getImplementation() === ExcelFactory::TYPE_SPOUT) {
			self::assertCount(1, $excel->getPhpExcel()->getSheets());
		}

		$sheet = $excel->addSheet('TestSheet');
		self::assertCount(1, $excel->getSheets());
		if($this->getImplementation() === ExcelFactory::TYPE_PHP_SPREADSHEET) {
			self::assertCount(1, $excel->getPhpExcel()->getAllSheets());
			self::assertEquals('TestSheet', $sheet->getSourceSheet()->getTitle());
		} elseif($this->getImplementation() === ExcelFactory::TYPE_SPOUT) {
			self::assertCount(1, $excel->getPhpExcel()->getSheets());
			self::assertEquals('TestSheet', $sheet->getSourceSheet()->getName());
		}

		$sheet = $excel->addSheet('TestSheet2');
		self::assertCount(2, $excel->getSheets());
		if($this->getImplementation() === ExcelFactory::TYPE_PHP_SPREADSHEET) {
			self::assertCount(2, $excel->getPhpExcel()->getAllSheets());
			self::assertEquals('TestSheet2', $sheet->getSourceSheet()->getTitle());
		} elseif($this->getImplementation() === ExcelFactory::TYPE_SPOUT) {
			self::assertCount(2, $excel->getPhpExcel()->getSheets());
			self::assertEquals('TestSheet2', $sheet->getSourceSheet()->getName());
		}
	}

	public function testSetActiveSheet() : void {
		$excel = ExcelFactory::createConfiguredExcel($this->getImplementation());

		self::assertCount(0, $excel->getSheets());

		$sheet1 = $excel->addSheet('TestSheet');
		self::assertCount(1, $excel->getSheets());

		$sheet2 = $excel->addSheet('TestSheet2');
		self::assertCount(2, $excel->getSheets());

		if($this->getImplementation() === ExcelFactory::TYPE_PHP_SPREADSHEET) {
			// in spreadsheet newly added sheets are NOT marked as active
			self::assertSame($excel->getPhpExcel()->getActiveSheet(), $sheet1->getSourceSheet());
			$excel->setActiveSheet($sheet2);
			self::assertSame($excel->getPhpExcel()->getActiveSheet(), $sheet2->getSourceSheet());
		} elseif($this->getImplementation() === ExcelFactory::TYPE_SPOUT) {
			// in spout newly added sheets are marked as active
			self::assertSame($excel->getPhpExcel()->getCurrentSheet(), $sheet2->getSourceSheet());
			$excel->setActiveSheet($sheet1);
			self::assertSame($excel->getPhpExcel()->getCurrentSheet(), $sheet1->getSourceSheet());
		}
	}
}