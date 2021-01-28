<?
/**
 * Модуль «ТСЖ»
 * Страница выгрузки лицевых счетов и показаний счетчиков с сайта в 1С
 * @package tszh
 */

// подключим все необходимые файлы:
require_once($_SERVER["DOCUMENT_ROOT"] . "/bitrix/modules/main/include/prolog_admin_before.php"); // первый общий пролог

use Bitrix\Main\Loader;

ob_start();
if (!CTszhFunctionalityController::CheckEdition())
{
    $c = ob_get_contents();
    ob_end_clean();
    require($_SERVER["DOCUMENT_ROOT"] . "/bitrix/modules/main/include/prolog_admin_after.php");
    echo $c;
    require($_SERVER["DOCUMENT_ROOT"] . "/bitrix/modules/main/include/epilog_admin.php");

    return;
}
$demoNotice = ob_get_contents();
ob_end_clean();

if (!Loader::includeModule("citrus.xlsx"))
{
    CAdminMessage::ShowMessage(GetMessage("TE_ERROR_MODULE_NOT_INSTALLED"));
}

require_once($_SERVER["DOCUMENT_ROOT"] . "/local/modules/citrus.xlsx/include.php"); // инициализация модуля
//require_once($_SERVER["DOCUMENT_ROOT"] . "/bitrix/modules/citrus.tszh/include.php"); // инициализация модуля ТСЖ
require_once($_SERVER["DOCUMENT_ROOT"] . "/local/modules/citrus.xlsx/prolog.php"); // пролог модуля

// подключим языковой файл
IncludeModuleLangFile(__FILE__);

set_time_limit(0);

$step = IntVal($_REQUEST['step']);
if ($step < 1 || $step > 2 || !check_bitrix_sessid())
{
    $step = 1;
}

// получим права доступа текущего пользователя на модуль ТСЖ
$POST_RIGHT = $APPLICATION->GetGroupRight("citrus.tszh");

// если нет прав - отправим к форме авторизации с сообщением об ошибке
if ($POST_RIGHT < "E")
    $APPLICATION->AuthForm(GetMessage("ACCESS_DENIED"));

$strError = false;

// удаление файла экспорта
if ($_REQUEST['delete_btn'] && check_bitrix_sessid())
{
    if (strlen($_SESSION['citrus.tszh.export.xlsx']['FILENAME']) > 0 && file_exists($_SERVER['DOCUMENT_ROOT'] . $_SESSION['citrus.tszh.export']['FILENAME']))
    {
        unlink($_SERVER['DOCUMENT_ROOT'] . $_SESSION['citrus.tszh.export.xlsx']['FILENAME']);
        unset($_SESSION['citrus.tszh.export.xlsx']);
    }
}

// проверка прав для портала
$arTszhFilter = $arTszhRight = array();
if (CModule::IncludeModule("vdgb.portaltszh") && TSZH_PORTAL === true && !$USER->IsAdmin())
{
    global $USER;
    $arGroups = $USER->GetUserGroupArray();

    $rsPerms = \Vdgb\PortalTszh\PermsTable::getList(
        array(
            "filter" => array("@GROUP_ID" => $arGroups)
        )
    );

    while ($arPerms = $rsPerms->fetch())
    {
        if ($arPerms["PERMS"] >= "R")
        {
            $arTszhRight[$arPerms["TSZH_ID"]] = $arPerms["PERMS"];
        }
    }

    $arTszhFilter["@ID"] = array_keys($arTszhRight);
}

// обработка и подготовка данных
if ($step > 1 && check_bitrix_sessid())
{

    if ($step == 2)
    {
        $rowXLSXIndex = IntVal($_REQUEST['rowXLSXIndex']) > 0 ? IntVal($_REQUEST['rowXLSXIndex']) : 1; //запись начинаем с первой строки xlsx-файла

        $lastID = isset($_REQUEST['lastID']) && IntVal($_REQUEST['lastID']) > 0 ? IntVal($_REQUEST['lastID']) : 0;

        if ($lastID <= 0)
        {
            $tszhExportFormat               = in_array($_REQUEST['tszhExportFormat'], Array(
                'xlsx'
            )) ? $_REQUEST['tszhExportFormat'] : 'xlsx';
            $_SESSION['citrus.tszh.export.xlsx'] = Array(
                'FILENAME' => "/upload/export_" . date('Y-m-d_His') . ".{$tszhExportFormat}",
                'TSZH_ID' => IntVal($_REQUEST["tszhID"]),
                'TSZH_OWNER' => trim($_REQUEST["tszhOwner"]),
                'STEP_TIME' => isset($_REQUEST['step_time']) && IntVal($_REQUEST['step_time']) > 0 ? IntVal($_REQUEST['step_time']) : 25,
                'TSZH' => CTszh::GetByID($_REQUEST["tszhID"]),
                'FORMAT' => $tszhExportFormat,
                "ONLY_OWNER_VALUES" => $_REQUEST['tszhExportOnlyOwnerValues'] == 'Y' ? true : false,
            );

            $dateFrom = trim($_REQUEST['tszhExportDateFrom']);
            $dateTo   = trim($_REQUEST['tszhExportDateTo']);
            if (strlen($dateFrom) > 0 || strlen($dateTo) > 0)
            {
                CheckFilterDates($dateFrom, $dateTo, $date1_wrong, $date2_wrong, $date2_less);
                if ($date1_wrong == "Y")
                    $strError .= GetMessage("TSZH_EXPORT_ERROR_DATE_1");
                if ($date2_wrong == "Y")
                    $strError .= GetMessage("TSZH_EXPORT_ERROR_DATE_2");
                if ($date2_less == "Y")
                    $strError .= GetMessage("TSZH_EXPORT_ERROR_DATE_3");

                if (strlen($strError) <= 0)
                {
                    if (strlen($dateFrom) > 0)
                        $_SESSION['citrus.tszh.export.xlsx']['DATE_FROM'] = $dateFrom;
                    if (strlen($dateTo) > 0)
                        $_SESSION['citrus.tszh.export.xlsx']['DATE_TO'] = $dateTo;
                }
            }
            else
            {
                $currentDate                               = new DateTime();
                $_SESSION['citrus.tszh.export.xlsx']['DATE_TO'] = $dateTo = $currentDate->format('Y-m-d H:i:s');
                $currentDate->modify('-1 month');
                $_SESSION['citrus.tszh.export.xlsx']['DATE_FROM'] = $dateFrom = $currentDate->format('Y-m-d H:i:s');
            }

            CUserOptions::SetOption('citrus.tszh', 'export.format.xlsx', 'xlsx');
            CUserOptions::SetOption('citrus.tszh', 'export.onlyOwnerValues', $_SESSION['citrus.tszh.export.xlsx']['ONLY_OWNER_VALUES']);
        }

        $filename         = $_SESSION['citrus.tszh.export.xlsx']['FILENAME'];
        $tszhID           = $_SESSION['citrus.tszh.export.xlsx']['TSZH_ID'];
        $tszhOwner        = $_SESSION['citrus.tszh.export.xlsx']['TSZH_OWNER'];
        $stepTime         = $_SESSION['citrus.tszh.export.xlsx']['STEP_TIME'];
        $arTszh           = $_SESSION['citrus.tszh.export.xlsx']['TSZH'];
        $tszhExportFormat = 'xlsx';

        $limit = CUserOptions::GetOption('citrus.tszh', 'export.MetterAccountsLimit', false);
        if (!$limit)
        {
            $limit = 2000;
            CUserOptions::SetOption('citrus.tszh', 'export.MetterAccountsLimit', $limit);
        }

        if (is_array($arTszh))
        {
            $orgName = $arTszh["NAME"];
            $orgINN  = $arTszh["INN"];
        }
        else
        {
            $strError .= GetMessage("TSZH_EXPORT_NO_TSZH_SELECTED");
        }

        if (strlen($strError) <= 0)
        {
            $obExport = new CTszhExportXLSX();

            //$obExport::$limit = $limit;

            $arValueFilter = Array();
            if ($_SESSION['citrus.tszh.export.xlsx']['ONLY_OWNER_VALUES'])
                $arValueFilter['MODIFIED_BY_OWNER'] = "Y";
            if ($_SESSION['citrus.tszh.export.xlsx']['DATE_FROM'])
                $arValueFilter[">=TIMESTAMP_X"] = $_SESSION['citrus.tszh.export.xlsx']['DATE_FROM'];
            if ($_SESSION['citrus.tszh.export.xlsx']['DATE_TO'])
            {
                $dateTo   = $_SESSION['citrus.tszh.export.xlsx']['DATE_TO'];
                $arDateTo = ParseDateTime($dateTo);
                if (!isset($arDateTo["HH"]) && !isset($arDateTo["H"]) && !isset($arDateTo["GG"]) && !isset($arDateTo["G"]))
                {
                    $dateTo = ConvertTimeStamp(
                        AddToTimeStamp(array("HH" => 23, "MI" => 59, "SS" => 59), MakeTimeStamp($dateTo)),
                        "FULL"
                    );
                }
                $arValueFilter["<=TIMESTAMP_X"] = $dateTo;
            }

            list($total, $last, $limit, $lastID, $rowXLSXIndex) = $obExport->DoExport($orgName, $orgINN, $_SERVER['DOCUMENT_ROOT'] . $filename, $tszhOwner, $rowXLSXIndex, Array("TSZH_ID" => $tszhID), $stepTime, $lastID, $arValueFilter);

            if ($lastID === false)
            {
                if ($ex = $APPLICATION->GetException())
                {
                    $strError = $ex->GetString() . '<br />';
                }
                else
                {
                    $strError = GetMessage("TE_ERROR_EXPORT") . "<br />";
                }
                $APPLICATION->ResetException();
            }
            elseif ($last > 0)
            {
                $href = $APPLICATION->GetCurPageParam("lastID=$lastID&step=$step&rowXLSXIndex=$rowXLSXIndex&" . bitrix_sessid_get(), Array(
                    'tszhID',
                    'tszhOwner',
                    "lastID",
                    'step',
                    'rowXLSXIndex',
                    'step_time',
                    'sessid'
                ));
                ?>
                <!doctype html>
                <html>
                <meta http-equiv="REFRESH" content="0;url=<?=$href?>">
                </html>
                <body>
                <?
                echo GetMessage("TSZH_EXPORT_PROGRESS", array('#CNT#' => $total, '#TOTAL#' => $last));
                ?>
                </body>
                <?
                return;
            }
        }

        if (strlen($strError) <= 0)
        {
            // запрещает ввод показаний счетчиков (будет разрешен снова после импорта)
            //COption::SetOptionString("citrus.tszh", "meters_block_edit", "Y");
        }

        if (strlen($strError) > 0)
            $step = 1;
    }
}
?>
<?

$APPLICATION->SetTitle(GetMessage("TE_PAGE_TITLE"));

require($_SERVER["DOCUMENT_ROOT"] . "/bitrix/modules/main/include/prolog_admin_after.php"); // второй общий пролог

echo $demoNotice;

//==========================================================================================
//==========================================================================================


CAdminMessage::ShowMessage($strError);

?>
    <form method="POST" action="<? echo $sDocPath ?>?lang=<? echo LANG ?>" enctype="multipart/form-data" name="dataload"
          id="dataload">
        <?=bitrix_sessid_post()?>
        <?

        // вывод страницы
        //==========================================================================================
        $aTabs = array(
            array(
                "DIV" => "edit1",
                "TAB" => GetMessage("TE_TAB1"),
                "ICON" => "iblock",
                "TITLE" => GetMessage("TE_TAB1_TITLE")
            ),
            array(
                "DIV" => "edit2",
                "TAB" => GetMessage("TE_TAB2"),
                "ICON" => "iblock",
                "TITLE" => GetMessage("TE_TAB2_TITLE")
            ),
        );

        $tabControl = new CAdminTabControl("tabControl", $aTabs, false);
        $tabControl->Begin();
        ?>

        <?
        $tabControl->BeginNextTab();

        if ($step == 1)
        { ?>
            <tr>
                <td></td>
                <td>
                    <? echo BeginNote(); ?>
                    <? echo GetMessage("TSZH_EXPORT_TEXT"); ?>
                    <? echo EndNote(); ?>
                </td>
            </tr>
            <tr>
                <td><? echo GetMessage("TSZH_EXPORT_TSZH"); ?>:</td>
                <td>
                    <select name="tszhID" id="tszhID">
                        <?
                        $dbTszh = CTszh::GetList(Array("NAME" => "ASC"), $arTszhFilter);
                        while ($arTszh = $dbTszh->Fetch())
                        {
                            ?>
                        <option value="<?=$arTszh["ID"]?>"<? if ($arTszh["ID"] == CUserOptions::GetOption('citrus.tszh.import', 'TszhID', false))
                            echo " selected"; ?>>[<?=htmlspecialcharsex($arTszh["ID"])?>]
                            &nbsp;<?=htmlspecialcharsex($arTszh["NAME"])?></option><?
                        }
                        ?>
                    </select>
                </td>
            </tr>
            <tr>
                <td><? echo GetMessage("TSZH_EXPORT_STEP_TIME"); ?>:</td>
                <td>
                    <input name="step_time" value="25" type="text"/>
                </td>
            </tr>
            <tr><?
                $dateFrom = trim($_REQUEST['tszhExportDateFrom']);
                $dateTo   = trim($_REQUEST['tszhExportDateTo']);
                ?>
                <td><label for="tszhExportDateFrom"> <?=GetMessage("TSZH_EXPORT_PERIOD")?></label></td>
                <td>
                    <?=CalendarPeriod("tszhExportDateFrom", $dateFrom, "tszhExportDateTo", $dateTo, "dataload", "N", 'id="tszhExportPeriod"')?>
                </td>
            </tr>
            <tr><?
                $tszhExportOnlyOwnerValues = CUserOptions::GetOption('citrus.tszh', 'export.onlyOwnerValues', false);
                ?>
                <td><label for="tszhExportOnlyOwnerValues"> <?=GetMessage("TSZH_EXPORT_VALUES_BY_OWNER")?></label></td>
                <td>
                    <input type="checkbox" value="Y"
                           name="tszhExportOnlyOwnerValues"<?=($tszhExportOnlyOwnerValues ? ' checked="checked"' : '')?>
                           id="tszhExportOnlyOwnerValues"/>
                </td>
            </tr>
            <tr>
                <td>
                    <?=GetMessage("TSZH_EXPORT_OWNER")?>
                </td>
                <td>
                    <select name="tszhOwner" id="tszhOwner">
                        <?
                        $dbRes = CIBlockElement::GetList(
                            array("NAME" => "ASC"),
                            array(
                                "IBLOCK_ID" => \Bitrix\Main\Config\Option::get("citrus.xlsx", "owners_iblock_id"),
                                "ACTIVE"=>"Y"
                            ),
                            false,
                            false,
                            array("ID", "NAME")
                        );

                        $ownersValue = array();
                        ?>
                        <option value="0">
                        &nbsp;<?=GetMessage("TSZH_EXPORT_OWNER_ALL")?>
                        </option>
                        <?
                        while ($arFields = $dbRes->GetNext())
                        {?>
                        <option value="<?=$arFields["ID"]?>">
                            [<?=htmlspecialcharsex($arFields["ID"])?>]
                            &nbsp;<?=htmlspecialcharsex($arFields["NAME"])?>
                            </option><?
                        }?>
                </td>
            </tr>
            <tr><?

                $tszhExportFormat = CUserOptions::GetOption('citrus.tszh', 'export.format.xlsx', 'xml');

                ?>
                <td>
                    <input type="hidden" name="tszhExportFormat" value="xlsx">
                </td>
            </tr>
            <?
        }

        $tabControl->EndTab();
        ?>


        <?
        $tabControl->BeginNextTab();

        if ($step == 2)
        { ?>
            <tr>
                <td>
                    <?=GetMessage("TE_EXPORT_DONE")?>.<br/>
                    <a href="<?=htmlspecialcharsbx($filename)?>" target="_blank"
                       download="<?=htmlspecialcharsbx(basename($filename))?>"><?=GetMessage("TE_DOWNLOAD_XML")?></a>
                </td>
            </tr>
            <?
        }

        $tabControl->EndTab();
        ?>

        <?
        $tabControl->Buttons();
        ?>
        <input type="hidden" name="step" value="<?=$step + 1?>"/>
        <?
        if ($step < 2)
        {
            ?>
            <input type="submit" name="submit_btn" value="<?=GetMessage("TE_NEXT_BTN")?>" class="adm-btn-save"/>
            <?
        }
        else
        {
            ?>
            <input type="submit" name="delete_btn" value="<?=GetMessage("TE_DELETE_FILE_BTN")?>"/>
            <?
        }
        ?>

        <?
        $tabControl->End();
        ?>
    </form>

    <script language="JavaScript">
        <!--
        <?if ($step < 2):?>
        tabControl.SelectTab("edit1");
        tabControl.DisableTab("edit2");
        <?elseif ($step >= 2):?>
        tabControl.SelectTab("edit2");
        tabControl.DisableTab("edit1");
        <?endif;?>
        //-->
    </script>


<? require($_SERVER["DOCUMENT_ROOT"] . "/bitrix/modules/main/include/epilog_admin.php"); ?>