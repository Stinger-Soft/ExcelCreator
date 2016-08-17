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

class IntegrationTest extends \PHPUnit_Framework_TestCase {

	public function testSimpleCycle() {
		$excel = new ConfiguredExcel();
		
		$excel->setTitle('TestCase');
		$excel->setCreator('Me');
		$excel->setCompany('StingerSoft');
		
		$sheet1 = $excel->addSheet('Test Sheet1');
		
		for($i = 0; $i < 10; $i++) {
			$binding = new ColumnBinding();
			$binding->setLabel('Column ' . $i);
			$binding->setLabelTranslationDomain(false);
			$binding->setBinding('[' . $i . ']'); //
			$binding->setColumnWidth('auto');
			$binding->setWrapText(true);
			$binding->setOutline(1);
			$binding->setHeaderBackgroundColor('000000');
			$binding->setHeaderFontColor('FFFFFF');
			$binding->setDataFontColor('$FFFFFF');
			$binding->setDataBackgroundColor('$FAFAFA');
			$binding->setFormatter(function ($value) {
				return strtoupper($value);
			});
			$sheet1->addColumnBinding($binding);
		}
		$sheet1->setData($this->getArrayData());
		$sheet1->applyData();
	}

	public function testBindingFunctionCycle() {
		$excel = new ConfiguredExcel();
		
		$excel->setTitle('TestCase');
		$excel->setCreator('Me');
		
		$sheet1 = $excel->addSheet('Test Sheet1');
		
		for($i = 0; $i < 10; $i++) {
			$binding = new ColumnBinding();
			$binding->setLabel('Column ' . $i);
			$binding->setLabelTranslationDomain(false);
			$binding->setColumnWidth('auto');
			$binding->setBinding(function (ColumnBinding $bind, $item) {
				return 'Bound via callable';
			});
			$sheet1->addColumnBinding($binding);
		}
		$sheet1->setData($this->getArrayData());
		$sheet1->applyData();
	}

	protected function getArrayData($count = 10, $columns = 10) {
		$data = array();
		for($i = 0; $i < $count; $i++) {
			$item = array();
			for($j = 0; $j < $columns; $j++) {
				$item[$j] = 'Test ' . $i . ':' . $j;
			}
			$data[] = $item;
		}
		return $data;
	}
}