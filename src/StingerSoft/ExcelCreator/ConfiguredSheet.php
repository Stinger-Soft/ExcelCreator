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

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use Symfony\Contracts\Translation\TranslatorInterface;

/**
 * Abstraction layer to represent a single worksheet inside an excel file
 */
class ConfiguredSheet implements ConfiguredSheetInterface {

	protected $configuredSheet = null;

	/**
	 * Default constructor
	 *
	 * @param ConfiguredExcel $excel
	 *            The parent excel file
	 * @param Worksheet $sheet
	 *            The underlying PHP Excel worksheet
	 * @param TranslatorInterface $translator
	 *            Translator to support translation of bindings
	 */
	public function __construct(ConfiguredExcel $excel, Worksheet $sheet, TranslatorInterface $translator = null) {
		$this->configuredSheet = new Spreadsheet\ConfiguredSheet($excel, $sheet, $translator);
	}

	/**
	 * @inheritDoc
	 */
	public function addColumnBinding(ColumnBinding $binding): ConfiguredSheetInterface {
		$this->configuredSheet->addColumnBinding($binding);
	}

	/**
	 * @inheritDoc
	 */
	public function getIndexForBinding(ColumnBinding $binding) {
		return $this->configuredSheet->getIndexForBinding($binding);
	}

	/**
	 * @inheritDoc
	 */
	public function setData($data): ConfiguredSheetInterface {
		$this->configuredSheet->setData($data);
		return $this;
	}

	/**
	 * @inheritDoc
	 */
	public function applyData(int $startColumn = 1, int $headerRow = 1): ConfiguredSheetInterface {
		$this->configuredSheet->applyData($startColumn, $headerRow);
		return $this;
	}

	/**
	 * @inheritDoc
	 */
	public function setExtraData($extraData): ConfiguredSheetInterface {
		$this->configuredSheet->setExtraData($extraData);
		return $this;
	}

	/**
	 * Returns the underlying PHP Excel Sheet object
	 *
	 * @return Worksheet
	 */
	public function getSheet(): Worksheet {
		return $this->configuredSheet->getSheet();
	}

	/**
	 * @inheritDoc
	 */
	public function getGroupByBinding(): ?ColumnBinding {
		return $this->configuredSheet->getGroupByBinding();
	}

	/**
	 * @inheritDoc
	 */
	public function setGroupByBinding(?ColumnBinding $groupByBinding = null): ConfiguredSheetInterface {
		$this->configuredSheet->setGroupByBinding($groupByBinding);
		return $this;
	}
}
