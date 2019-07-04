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
use StingerSoft\ExcelCreator\Spreadsheet\ConfiguredExcel;
use Symfony\Contracts\Translation\TranslatorInterface;

class HelperTest extends TestCase {

	use Helper;
	
	public function testDecodeHtmlEntity() {
		$this->assertEquals('Laphroaig', $this->decodeHtmlEntity('La&shy;phroaig'));
		$this->assertEquals('Laphroaig', $this->decodeHtmlEntity('Laphroaig'));
	}
	
	public function testTranslate() {
		$this->assertEquals('Laphroaig', $this->translate('Laphroaig', false));
		
		$translator = $this->getMockBuilder(TranslatorInterface::class)->setMethods(array('trans'))->getMockForAbstractClass();
		$translator->expects($this->exactly(2))->method('trans')->willReturn('translated');
		$this->translator = $translator;
		
		$this->assertEquals('translated', $this->translate('Laphroaig', null));
		$this->assertEquals('translated', $this->translate('Laphroaig', 'domain'));
	}
	
	public function testSetSheetTitle() {
		$excel = new ConfiguredExcel();
		$sheet = $excel->addSheet('test')->getSheet();
		
		$this->setSheetTitle($sheet, 'TestTitle');
		$this->assertEquals('TestTitle', $sheet->getTitle());
		
		$this->setSheetTitle($sheet, '0123456789012345678901234567890123456789');
		$this->assertEquals('0123456789012345678901234567890', $sheet->getTitle());
		
		$this->setSheetTitle($sheet, 'Test:::::');
		$this->assertEquals('Test_____', $sheet->getTitle());
	}

}