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

namespace StingerSoft\ExcelCreator\CellModifier;


use Box\Spout\Common\Entity\Cell as SpoutCell;
use Box\Spout\Common\Entity\Style\Style;
use PhpOffice\PhpSpreadsheet\Cell\Cell as SpreadsheetCell;
use StingerSoft\ExcelCreator\ColumnBinding;

class DateTimeModifier {

	/**
	 * @return callable
	 */
	public static function createDateModifier(): callable {
		return self::createFormatModifier('m/d/yyyy');
	}

	/**
	 * @param bool $addSeconds
	 * @return callable
	 */
	public static function createDateTimeModifier(bool $addSeconds = false): callable {
		return self::createFormatModifier('m/d/yyyy h:mm' . ($addSeconds ? ':ss' : ''));
	}

	public static function createFormatModifier(string $formatCode): callable {
		/**
		 *
		 * @param ColumnBinding $binding
		 * @param SpoutCell|SpreadsheetCell $cell
		 */
		return static function(ColumnBinding $binding, &$cell) use ($formatCode) {
			if($cell instanceof SpoutCell) {
				$style = $cell->getStyle();
				if($style === null) {
					$style = new Style();
				}
				$style->setFormat($formatCode);
				$cell->setStyle($style);
				return;
			}
			if($cell instanceof SpreadsheetCell) {
				$style = $cell->getStyle();
				$style->applyFromArray(['numberFormat' => ['formatCode' => $formatCode]]);
				return;
			}
		};
	}
}