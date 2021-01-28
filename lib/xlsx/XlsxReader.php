<?php


namespace Citrus\Xlsx;

// Composer
include_once __DIR__ . "/../vendor/autoload.php";

use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use Citrus\Filters\TitleReadFilter;
use Citrus\Filters\FileReadFilter;

/**
 *  лас дл€ чтени€ из XLSX файла
 *
 * Class XlsxReader
 * @package Citrus
 */
class XlsxReader
{
    private $fileName;

    public function __construct($fileName)
    {
        $this->fileName = $fileName;
    }

    /**
     * „тение заголовка таблицы
     *
     * @param mixed ...$dateColumns - номера колонок с датой
     * @return mixed
     */
    public function readTitle(...$dateColumns)
    {
        $arResult = [];

        $reader = new Xlsx();

        if ($reader->canRead($this->fileName))
        {
            $reader->setReadFilter(new TitleReadFilter());

            $spreadsheet = $reader->load($this->fileName);
            $worksheet = $spreadsheet->getActiveSheet();

            foreach($dateColumns as $dateColumn)
            {
                $this->setDateFormat($worksheet, $dateColumn, 1, 1);
            }

            $arResult = $worksheet->toArray();
        }

        return $arResult[0];
    }

    /**
     * ѕошаговое чтение файла
     *
     * @param $maxColumn - номер последней читаемой колонки
     * @param $maxRow - номер последней читаемой строки
     * @param $minRow - номер первой читаемой строки
     * @param mixed ...$dateColumns - номера колонок с датой
     * @return array|mixed
     */
    public function readFileAsArray($maxColumn, $maxRow, $minRow, ...$dateColumns)
    {
        $arResult = [];

        $reader = new Xlsx();

        if ($reader->canRead($this->fileName))
        {
            $reader->setReadFilter(new FileReadFilter($maxColumn, $maxRow, $minRow));

            $spreadsheet = $reader->load($this->fileName);
            $worksheet = $spreadsheet->getActiveSheet();

            foreach($dateColumns as $dateColumn)
            {
                $this->setDateFormat($worksheet, $dateColumn, $minRow, $maxRow);
            }

            $arResult = $worksheet->toArray();

            $arResult = $this->unsetEmptyRows($arResult);

        }

        return $arResult;
    }


    /**
     * ”становка даты в формате DD.MM.YYYY
     *
     * @param $worksheet - рабоча€ книга
     * @param $column - столбец
     * @param $minRow - перва€ обрабатываема€ строка
     * @param $maxRow - последн€€ обрабатываема€ строка
     */
    private function setDateFormat(&$worksheet, $column, $minRow, $maxRow)
    {
        $worksheet->getStyleByColumnAndRow($column, $minRow, $column, $maxRow)->getNumberFormat()->setFormatCode("DD.MM.YYYY");
    }

    /**
     * ¬озвращает количество строк в файле
     *
     * @return int
     */
    public function getRowCountFromFile()
    {
        $count = 0;
        $reader = new Xlsx();

        if ($reader->canRead($this->fileName))
        {

            $spreadsheet = $reader->load($this->fileName);
            $worksheet = $spreadsheet->getActiveSheet();
            $count = $worksheet->getHighestDataRow();
        }

        return $count;
    }

    /**
     * ѕровер€ет, возможно ли прочитать файл по имени, добавленному при создании класса XlsxReader
     *
     * @return bool
     */
    public function canReadFile()
    {
        $result = false;

        if (file_exists($this->fileName))
        {
            $reader = new Xlsx();

            if ($reader->canRead($this->fileName))
            {
                $result = true;
            }

        }

        return $result;
    }

    /**
     * ”даление пустых строк из массива
     *
     * @param $array - фильтруемый массив
     * @return mixed
     */
    private function unsetEmptyRows($array)
    {
        foreach ($array as $key => $row)
        {
            if (empty(array_filter($row,
                function($value)
                {
                    if (trim($value) === '')
                    {
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                })
            ))
            {
                unset($array[$key]);
            }
        }

        return $array;
    }

}