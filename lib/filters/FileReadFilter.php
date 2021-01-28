<?php


namespace Citrus\Filters;

use PhpOffice\PhpSpreadsheet\Reader\IReadFilter;

/**
 * Фильтр для пошагового чтения файла
 *
 * Class FileReadFilter
 * @package Citrus\Filters
 */
class FileReadFilter implements IReadFilter
{

    private $maxColumn;
    private $maxRow;
    private $minRow;

    /**
     * Конструктор класса FileReadFilter
     *
     * FileReadFilter constructor.
     * @param $maxColumn - последний обрабатываемый на этом шаге столбец
     * @param $maxRow - последняя обрабатываемая на этом шаге строка
     * @param $minRow - первая обрабатываемая на этом шаге строка
     */
    public function __construct($maxColumn, $maxRow, $minRow)
    {
        $this->maxColumn = $maxColumn;
        $this->maxRow = $maxRow;
        $this->minRow = $minRow;
    }

    public function readCell($column, $row, $worksheetName = '')
    {
        if ($row <= $this->maxRow && $column <= $this->maxColumn && $row >= $this->minRow)
        {
            return true;
        }
        return false;
    }
}