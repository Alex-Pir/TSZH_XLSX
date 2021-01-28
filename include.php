<?php

// поключение языковых файлов
IncludeModuleLangFile(__FILE__);

define("CITRUS_MODULE_XLSX_ID", 'citrus.xlsx');

CModule::AddAutoloadClasses(
    CITRUS_MODULE_XLSX_ID,
    array(
        "CTszhExportXLSX" => "classes/general/tszh_export_xlsx.php",
        "CTszhImportXLSX" => "classes/general/tszh_import_xlsx.php",
        "Citrus\\Xlsx\\XlsxHelper" => "lib/xlsx/XlsxHelper.php",
        "Citrus\\Xlsx\\XlsxWriter" => "lib/xlsx/XlsxWriter.php",
        "Citrus\\Xlsx\\XlsxReader" => "lib/xlsx/XlsxReader.php",
        "Citrus\\Filters\\FileReadFilter" => "lib/filters/FileReadFilter.php",
        "Citrus\\Filters\\TitleReadFilter" => "lib/filters/TitleReadFilter.php"
    )
);
?>