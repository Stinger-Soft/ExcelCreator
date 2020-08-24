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
use StingerSoft\ExcelCreator\Spreadsheet\ConfiguredExcel;
use Symfony\Contracts\Translation\TranslatorInterface;

class HelperTest extends TestCase {

	use Helper;

	public function testDecodeHtmlEntity(): void {
		self::assertEquals('Laphroaig', $this->decodeHtmlEntity('La&shy;phroaig'));
		self::assertEquals('Laphroaig', $this->decodeHtmlEntity('Laphroaig'));
	}

	public function testTranslate(): void {
		self::assertEquals('Laphroaig', $this->translate('Laphroaig', false));

		$translator = $this->getMockBuilder(TranslatorInterface::class)->setMethods(['trans'])->getMockForAbstractClass();
		$translator->expects(self::exactly(2))->method('trans')->willReturn('translated');
		$this->translator = $translator;

		self::assertEquals('translated', $this->translate('Laphroaig', null));
		self::assertEquals('translated', $this->translate('Laphroaig', 'domain'));
	}

	public function testCreateTemporaryFile(): void {
		$fileName = self::createTemporaryFile('tmp', 'test');
		self::assertContains($fileName, self::getTemporaryFileNames());

		self::removeTemporaryFiles();
		self::assertEmpty(self::getTemporaryFileNames());
	}

	public function testSetSheetTitle(): void {
		$excel = new ConfiguredExcel();
		$sheet = $excel->addSheet('test')->getSheet();

		$this->setSheetTitle($sheet, 'TestTitle');
		self::assertEquals('TestTitle', $sheet->getTitle());

		$this->setSheetTitle($sheet, '0123456789012345678901234567890123456789');
		self::assertEquals('0123456789012345678901234567890', $sheet->getTitle());

		$this->setSheetTitle($sheet, 'Test:::::');
		self::assertEquals('Test_____', $sheet->getTitle());
	}

}