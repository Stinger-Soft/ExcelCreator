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

use OpenSpout\Common\Entity\Cell as SpoutCell;
use OpenSpout\Common\Entity\Style\Style;
use PhpOffice\PhpSpreadsheet\Cell\Cell as SpreadsheetCell;
use StingerSoft\ExcelCreator\ColumnBinding;

abstract class AbstractModifier {

	public static function createFormatModifier(string $formatCode, ?string $spoutFormatCode = null): callable {
		$spoutFormatCode = $spoutFormatCode ?? $formatCode;

		/**
		 *
		 * @param ColumnBinding             $binding
		 * @param SpoutCell|SpreadsheetCell $cell
		 */
		return static function (ColumnBinding $binding, $cell) use ($formatCode, $spoutFormatCode) {
			if($cell instanceof SpoutCell) {
				$style = $cell->getStyle();
				if($style === null) {
					$style = new Style();
				} else {
					$style = clone $style;
					$style->setId(null);
				}
				$style->setFormat($spoutFormatCode);
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