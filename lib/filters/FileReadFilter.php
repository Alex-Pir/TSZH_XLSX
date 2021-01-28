<?php


namespace Citrus\Filters;

use PhpOffice\PhpSpreadsheet\Reader\IReadFilter;

/**
 * ������ ��� ���������� ������ �����
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
     * ����������� ������ FileReadFilter
     *
     * FileReadFilter constructor.
     * @param $maxColumn - ��������� �������������� �� ���� ���� �������
     * @param $maxRow - ��������� �������������� �� ���� ���� ������
     * @param $minRow - ������ �������������� �� ���� ���� ������
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