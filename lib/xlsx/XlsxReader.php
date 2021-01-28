<?php


namespace Citrus\Xlsx;

// Composer
include_once __DIR__ . "/../vendor/autoload.php";

use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use Citrus\Filters\TitleReadFilter;
use Citrus\Filters\FileReadFilter;

/**
 * ���� ��� ������ �� XLSX �����
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
     * ������ ��������� �������
     *
     * @param mixed ...$dateColumns - ������ ������� � �����
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
     * ��������� ������ �����
     *
     * @param $maxColumn - ����� ��������� �������� �������
     * @param $maxRow - ����� ��������� �������� ������
     * @param $minRow - ����� ������ �������� ������
     * @param mixed ...$dateColumns - ������ ������� � �����
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
     * ��������� ���� � ������� DD.MM.YYYY
     *
     * @param $worksheet - ������� �����
     * @param $column - �������
     * @param $minRow - ������ �������������� ������
     * @param $maxRow - ��������� �������������� ������
     */
    private function setDateFormat(&$worksheet, $column, $minRow, $maxRow)
    {
        $worksheet->getStyleByColumnAndRow($column, $minRow, $column, $maxRow)->getNumberFormat()->setFormatCode("DD.MM.YYYY");
    }

    /**
     * ���������� ���������� ����� � �����
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
     * ���������, �������� �� ��������� ���� �� �����, ������������ ��� �������� ������ XlsxReader
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
     * �������� ������ ����� �� �������
     *
     * @param $array - ����������� ������
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