<?php

namespace Citrus\Xlsx;

// Composer
include_once __DIR__ . "/../vendor/autoload.php";

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;

trait XlsxHelper
{
    /**
     * ������ ����, ���� ���������� (���������� ��� ���������� ��������,
     * ��� ��� ���� ����� �� �������, �� ���� �������������)
     *
     * @param $fileName - �������� �����
     * @return Spreadsheet
     */
    protected function getActualWorkheet($fileName)
    {
        if (file_exists($fileName))
        {
            $reader = new Xlsx();
            $spreadsheet = $reader->load($fileName);
        }
        else
        {
            $spreadsheet = new Spreadsheet();
        }

        return $spreadsheet;
    }
}