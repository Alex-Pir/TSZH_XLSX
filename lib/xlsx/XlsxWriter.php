<?php
namespace Citrus\Xlsx;

// Composer
include_once __DIR__ . "/../vendor/autoload.php";

use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Border;

/**
 * ����� ��� ������ � XLSX ����
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
     * ����������� ������ XlsxWriter
     *
     * @param $fileName - �������� �����, � ������� �������������� ������
     * @param int $rowXLSXIndex - ������, � ������� �������� ������ � ����
     */
    public function __construct($fileName, $rowXLSXIndex = 1)
    {
        $this->spreadsheet = $this->getActualWorkheet($fileName);
        $this->worksheet = $this->spreadsheet->getActiveSheet();
        $this->rowXLSXIndex = $rowXLSXIndex;
    }

    /**
     * ������������� ������� ���������
     *
     * @param $columnStart - ��������� �������
     * @param $rowStart - ��������� ������
     * @param $columnEnd - �������� �������
     * @param $rowEnd - �������� ������
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
     * ������������� ������� ������ �������
     *
     * @param $columnStart - ��������� �������
     * @param $rowStart - ��������� ������
     * @param $columnEnd - �������� �������
     * @param $rowEnd - �������� ������
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
     * ������ ������ � ���� � ������� XLSX
     *
     * @param $arValues - ������ ��� ������
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
     * ������ ����������� ������ ������� �������
     *
     * @param $lastColumn - ��������� ������� �����
     */
    public function setColumnDimension($lastColumn)
    {
        for ($column = 1; $column <= $lastColumn; $column++)
        {
            $this->worksheet->getColumnDimensionByColumn($column)->setAutoSize(true);
        }
    }

    /**
     * ��������� XLSX ����
     *
     * @param $fileName - �������� �����
     */
    public function saveFile($fileName)
    {
        $writer = new Xlsx($this->spreadsheet);
        $writer->save($fileName);
    }

    /**
     * ���������� ������� ������ � �����
     *
     * @return int
     */
    public function getCurrentRow()
    {
        return $this->rowXLSXIndex;
    }

}
?>