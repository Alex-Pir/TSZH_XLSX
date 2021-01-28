<?php
namespace Citrus\Xlsx;

// Composer
include_once __DIR__ . "/../vendor/autoload.php";

use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Border;

/**
 * Класс для записи в XLSX файл
 *
 * Class XlsxWriter
 * @package Citrus
 */
class XlsxWriter
{
    use XlsxHelper;

    private $spreadsheet;
    private $worksheet;
    private $rowXLSXIndex;

    /**
     * Конструктор класса XlsxWriter
     *
     * @param $fileName - название файла, в который осуществляется запись
     * @param int $rowXLSXIndex - строка, с которой начинаем запись в файл
     */
    public function __construct($fileName, $rowXLSXIndex = 1)
    {
        $this->spreadsheet = $this->getActualWorkheet($fileName);
        $this->worksheet = $this->spreadsheet->getActiveSheet();
        $this->rowXLSXIndex = $rowXLSXIndex;
    }

    /**
     * Устанавливает границы заголовка
     *
     * @param $columnStart - начальная колонка
     * @param $rowStart - начальная строка
     * @param $columnEnd - конечная колонка
     * @param $rowEnd - конечная строка
     */
    public function setHeaderStyle($columnStart, $rowStart, $columnEnd, $rowEnd)
    {
        $styleArray = array(
            'borders' => array(
                'outline' => array(
                    'borderStyle' => Border::BORDER_THICK,
                    'color' => array('argb' => '00000000'),
                ),
            ),
        );

        $this->worksheet->getStyleByColumnAndRow($columnStart, $rowStart, $columnEnd, $rowEnd)->applyFromArray($styleArray);
    }

    /**
     * Устанавливает границы внутри таблицы
     *
     * @param $columnStart - начальная колонка
     * @param $rowStart - начальная строка
     * @param $columnEnd - конечная колонка
     * @param $rowEnd - конечная строка
     */
    public function setTableStyle($columnStart, $rowStart, $columnEnd, $rowEnd)
    {
        $styleArray = array(
            'borders' => array(
                'allBorders' => array(
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => array('argb' => '00000000'),
                ),
            ),
        );

        $this->worksheet->getStyleByColumnAndRow($columnStart, $rowStart, $columnEnd, $rowEnd)->applyFromArray($styleArray);
    }

    /**
     * Запись данных в файл в формате XLSX
     *
     * @param $arValues - массив для записи
     */
    public function saveIntoXLSX($arValues)
    {
        $column = 1;

        foreach ($arValues as $value)
        {
            $this->worksheet->setCellValueByColumnAndRow($column++, $this->rowXLSXIndex, $value);
        }

        $this->rowXLSXIndex++;
    }

    /**
     * Задает необходимый размер каждому столбцу
     *
     * @param $lastColumn - последний столбец файла
     */
    public function setColumnDimension($lastColumn)
    {
        for ($column = 1; $column <= $lastColumn; $column++)
        {
            $this->worksheet->getColumnDimensionByColumn($column)->setAutoSize(true);
        }
    }

    /**
     * Сохраняет XLSX файл
     *
     * @param $fileName - название файла
     */
    public function saveFile($fileName)
    {
        $writer = new Xlsx($this->spreadsheet);
        $writer->save($fileName);
    }

    /**
     * Возвращает текущую строку в файле
     *
     * @return int
     */
    public function getCurrentRow()
    {
        return $this->rowXLSXIndex;
    }

}
?>