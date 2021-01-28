<?
IncludeModuleLangFile(__FILE__);

$module_rights = $APPLICATION->GetGroupRight("citrus.tszh");


if($module_rights > "D")
{

    if ($module_rights >= 'E') {
        $export_menu_items[] = Array(
            "text" => GetMessage("TSZH_MENU_EXPORT_XLSX_METERS"),
            "url" => "tszh_export_xlsx.php?lang=" . LANG,
            "more_url" => array(),
            "title" => GetMessage("TSZH_MENU_EXPORT_XLSX_TITLE")
        );
        $export_menu_items[] = array(
            "text" => GetMessage("TSZH_MENU_IMPORT_XLSX_METERS"),
            "url" => "tszh_import_xlsx.php?lang=" . LANG,
            "more_url" => array(),
            "title" => GetMessage("TSZH_MENU_IMPORT_XLSX_TITLE")
        );

        $aMenu = array(
            "parent_menu" => "global_menu_services",
            "section" => "citrus.xlsx",
            "sort" => 10,
            "text" => GetMessage("TSZH_EXPORT_MENU_SECT"),
            "title" => GetMessage("TSZH_EXPORT_MENU_SECT_TITLE"),
            "icon" => "tszh_menu_icon",
            "page_icon" => "tszh_page_icon",
            "items_id" => "menu_citrus_xlsx",
            "items" => $export_menu_items,
        );

        return $aMenu;
    }
}
return false;
?>