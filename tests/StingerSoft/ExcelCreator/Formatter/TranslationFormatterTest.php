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

namespace StingerSoft\ExcelCreator\Formatter;

use PHPUnit\Framework\MockObject\Exception;
use PHPUnit\Framework\MockObject\MockObject;
use PHPUnit\Framework\TestCase;
use Symfony\Contracts\Translation\TranslatorInterface;

class TranslationFormatterTest extends TestCase {

    /**
     * @throws Exception
     */
    public function testCreateTranslationFormatter(): void {
		/** @var TranslatorInterface|MockObject $translator */
		$translator = $this->createMock(TranslatorInterface::class);
		$translator->method('trans')->willReturn('translated');

		$formatter = TranslationFormatter::createTranslationFormatter($translator, 'test');

		self::assertEquals('translated', $formatter('test'));
	}

}