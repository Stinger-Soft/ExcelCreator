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

use OpenSpout\Common\Exception\IOException;
use Symfony\Contracts\Translation\TranslatorInterface;

class ExcelFactory {

	/**
	 * @string
	 */
	public const TYPE_SPOUT = 'spout';

	public const TYPE_PHP_SPREADSHEET = 'spreadsheet';

	private function __construct() {
	}

	/**
	 * @param string $type
	 * @param TranslatorInterface|null $translator
	 * @return ConfiguredExcelInterface
	 */
	public static function createConfiguredExcel(string $type = self::TYPE_PHP_SPREADSHEET, ?TranslatorInterface $translator = null): ConfiguredExcelInterface {
		switch($type) {
			case self::TYPE_SPOUT:
				try {
					return new Spout\ConfiguredExcel($translator);
				} catch(IOException $e) {
					throw new \RuntimeException('Cannot create instance of Spout ConfiguredExcel', 0, $e);
				}
			default:
				return new Spreadsheet\ConfiguredExcel($translator);
		}
	}

}