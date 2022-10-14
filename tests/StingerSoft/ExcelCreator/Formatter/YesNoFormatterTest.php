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
namespace StingerSoft\ExcelCreator\Formatter;

use PHPUnit\Framework\TestCase;

class YesNoFormatterTest extends TestCase {

	public function testCreateTranslationFormatter(): void {
		$pirateYesNoFormatter = YesNoFormatter::createYesNoFormatter('Arrrr!', 'Avast!');
		
		self::assertEquals('Arrrr!', $pirateYesNoFormatter(true));
		self::assertEquals('Avast!', $pirateYesNoFormatter(false));
		self::assertEquals('Avast!', $pirateYesNoFormatter('nonono'));
	}
}