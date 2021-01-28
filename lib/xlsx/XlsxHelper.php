<?php

namespace Citrus\Xlsx;

// Composer
include_once __DIR__ . "/../vendor/autoload.php";

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;

trait XlsxHelper
{
    /**
     * Читаем файл, если существует (необходимо для пошагового экспорта,
     * так как если этого не сделать, то файл перезапишется)
     *
     * @param $fileName - название файла
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