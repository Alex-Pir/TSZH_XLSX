<?php
namespace Citrus;

use PhpOffice\PhpSpreadsheet\Spreadsheet;

class XlsxWorksheetStorage
{
    private static $xlsxWriter;

    public static function init()
    {
        if (self::$xlsxWriter == null)
        {
            self::$xlsxWriter = new XlsxWriter(1);
        }

        return self::$xlsxWriter;
    }

}
?>