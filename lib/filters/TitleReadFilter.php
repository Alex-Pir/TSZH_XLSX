<?php


namespace Citrus\Filters;

use PhpOffice\PhpSpreadsheet\Reader\IReadFilter;

/**
 * ������ ��� ������ ������-��������� �����
 *
 * Class TitleReadFilter
 * @package Citrus\Filters
 */
class TitleReadFilter implements IReadFilter
{
    public function readCell($column, $row, $worksheetName = '')
    {
        // Read title row
        if ($row == 1) {
            return true;
        }
        return false;
    }
}