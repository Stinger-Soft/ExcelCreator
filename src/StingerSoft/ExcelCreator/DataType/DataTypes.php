<?php
declare(strict_types=1);

/*
 * This file is part of the PEC Platform ExcelCreator.
 *
 * (c) PEC project engineers &amp; consultants
 *
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace StingerSoft\ExcelCreator\DataType;

use OpenSpout\Common\Entity\Cell as SpoutCell;
use PhpOffice\PhpSpreadsheet\Cell\DataType;

class DataTypes {

	public const DATA_TYPE_STRING2 = 'str';
	public const DATA_TYPE_STRING = 's';
	public const DATA_TYPE_FORMULA = 'f';
	public const DATA_TYPE_NUMERIC = 'n';
	public const DATA_TYPE_BOOL = 'b';
	public const DATA_TYPE_NULL = 'null';
	public const DATA_TYPE_INLINE = 'inlineStr';
	public const DATA_TYPE_ERROR = 'e';

	protected const SPOUT_DATA_TYPE_STRING2 = SpoutCell::TYPE_STRING;
	protected const SPOUT_DATA_TYPE_STRING = SpoutCell::TYPE_STRING;
	protected const SPOUT_DATA_TYPE_FORMULA = SpoutCell::TYPE_FORMULA;
	protected const SPOUT_DATA_TYPE_NUMERIC = SpoutCell::TYPE_NUMERIC;
	protected const SPOUT_DATA_TYPE_BOOL = SpoutCell::TYPE_BOOLEAN;
	protected const SPOUT_DATA_TYPE_NULL = SpoutCell::TYPE_EMPTY;
	protected const SPOUT_DATA_TYPE_INLINE = SpoutCell::TYPE_STRING;
	protected const SPOUT_DATA_TYPE_ERROR = SpoutCell::TYPE_ERROR;

	protected const SPREADSHEET_DATA_TYPE_STRING2 = DataType::TYPE_STRING2;
	protected const SPREADSHEET_DATA_TYPE_STRING = DataType::TYPE_STRING;
	protected const SPREADSHEET_DATA_TYPE_FORMULA = DataType::TYPE_FORMULA;
	protected const SPREADSHEET_DATA_TYPE_NUMERIC = DataType::TYPE_NUMERIC;
	protected const SPREADSHEET_DATA_TYPE_BOOL = DataType::TYPE_BOOL;
	protected const SPREADSHEET_DATA_TYPE_NULL = DataType::TYPE_NULL;
	protected const SPREADSHEET_DATA_TYPE_INLINE = DataType::TYPE_INLINE;
	protected const SPREADSHEET_DATA_TYPE_ERROR = DataType::TYPE_ERROR;

	protected const SPOUT_MAPPING = [
		self::DATA_TYPE_STRING2 => self::SPOUT_DATA_TYPE_STRING2,
		self::DATA_TYPE_STRING  => self::SPOUT_DATA_TYPE_STRING,
		self::DATA_TYPE_FORMULA => self::SPOUT_DATA_TYPE_FORMULA,
		self::DATA_TYPE_NUMERIC => self::SPOUT_DATA_TYPE_NUMERIC,
		self::DATA_TYPE_BOOL    => self::SPOUT_DATA_TYPE_BOOL,
		self::DATA_TYPE_NULL    => self::SPOUT_DATA_TYPE_NULL,
		self::DATA_TYPE_INLINE  => self::SPOUT_DATA_TYPE_INLINE,
		self::DATA_TYPE_ERROR   => self::SPOUT_DATA_TYPE_ERROR,
	];

	protected const SPREADSHEET_MAPPING = [
		self::DATA_TYPE_STRING2 => self::SPREADSHEET_DATA_TYPE_STRING2,
		self::DATA_TYPE_STRING  => self::SPREADSHEET_DATA_TYPE_STRING,
		self::DATA_TYPE_FORMULA => self::SPREADSHEET_DATA_TYPE_FORMULA,
		self::DATA_TYPE_NUMERIC => self::SPREADSHEET_DATA_TYPE_NUMERIC,
		self::DATA_TYPE_BOOL    => self::SPREADSHEET_DATA_TYPE_BOOL,
		self::DATA_TYPE_NULL    => self::SPREADSHEET_DATA_TYPE_NULL,
		self::DATA_TYPE_INLINE  => self::SPREADSHEET_DATA_TYPE_INLINE,
		self::DATA_TYPE_ERROR   => self::SPREADSHEET_DATA_TYPE_ERROR,
	];

	public static function getSpoutCellType(string $cellType): int {
		return self::SPOUT_MAPPING[$cellType] ?? self::SPOUT_DATA_TYPE_STRING;
	}

	public static function getSpreadsheetCellType(string $cellType): string {
		return self::SPREADSHEET_MAPPING[$cellType] ?? self::SPREADSHEET_DATA_TYPE_STRING;
	}

}