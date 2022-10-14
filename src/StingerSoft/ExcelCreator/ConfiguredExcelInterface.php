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

use Doctrine\Common\Collections\Collection;
use Exception;

interface ConfiguredExcelInterface {

	/**
	 * Adds and returns a new sheet
	 *
	 * @param string $title
	 *            The title of the new sheet
	 * @return ConfiguredSheetInterface
	 */
	public function addSheet(string $title): ConfiguredSheetInterface;

	/**
	 * Returns the worksheets of this excel file
	 *
	 * @return ConfiguredSheetInterface[]|Collection
	 */
	public function getSheets(): Collection;

	/**
	 * Set the given sheet to be the active one.
	 *
	 * @param ConfiguredSheetInterface $sheet the sheet to make active
	 * @throws Exception in case the sheet cannot be found or cannot be made active
	 */
	public function setActiveSheet(ConfiguredSheetInterface $sheet): void;

	/**
	 * Returns the title of this excel file
	 *
	 * @return string|null The title of this excel file
	 */
	public function getTitle(): ?string;

	/**
	 * Sets the title of this excel file
	 *
	 * @param string|null $title
	 *            The title of this excel file
	 * @return ConfiguredExcelInterface
	 */
	public function setTitle(?string $title = null): self;

	/**
	 * Get the author of this excel file
	 *
	 * @return string|null The author of this excel file
	 */
	public function getCreator(): ?string;

	/**
	 * Sets the author of this excel file
	 *
	 * @param string|null $creator
	 *            The author of this excel file
	 * @return ConfiguredExcelInterface
	 */
	public function setCreator(?string $creator = null): self;

	/**
	 *
	 * @return string|null The company this excel file is created by
	 */
	public function getCompany(): ?string;

	/**
	 *
	 * @param string|null $company
	 *            The company this excel file is created by
	 * @return ConfiguredExcelInterface
	 */
	public function setCompany(?string $company = null): self;

	/**
	 * @param string $filename
	 * @return ConfiguredExcelInterface
	 */
	public function writeToFile(string $filename): self;

}