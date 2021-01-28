<?php

require_once($_SERVER["DOCUMENT_ROOT"] . "/bitrix/modules/main/include/prolog_admin_before.php"); // первый общий пролог
require_once($_SERVER["DOCUMENT_ROOT"] . "/local/modules/citrus.xlsx/include.php"); // инициализация модуля
require_once($_SERVER["DOCUMENT_ROOT"] . "/local/modules/citrus.xlsx/prolog.php"); // пролог модуля

use Bitrix\Main\Loader;
use \Bitrix\Main\Web\Json;

\Bitrix\Main\UI\Extension::load("ui.progressbar");

IncludeModuleLangFile(__FILE__);

//Получим права доступа на модуль ЖКХ
$POST_RIGHT = $APPLICATION->GetGroupRight("citrus.tszh");

//если прав нет - отправим к форме авторизации с сообщением об ошибке
if ($POST_RIGHT < "U")
{
    $APPLICATION->AuthForm(GetMessage("ACCESS_DENIED"));
}

if (!Loader::includeModule("citrus.xlsx"))
{
    CAdminMessage::ShowMessage(GetMessage("TI_ERROR_MODULE_NOT_INSTALLED"));
}

$maxTime = intval(ini_get("max_execution_time"));


@set_time_limit(0);

$start_time = time();

$arErrors = [];
$arMessage = [];

$mayCreateTszhFlag = false;

$lastAccount = "";
$lastRow = 2; //начинаем обрабатывать файл со 2 строки, исключая заголовок

if ($_REQUEST["Import"] == "Y" && (array_key_exists("NS", $_REQUEST) || array_key_exists("bFirst", $_REQUEST)))
{
    require_once($_SERVER["DOCUMENT_ROOT"] . "/bitrix/modules/main/include/prolog_admin_js.php");

    CUtil::JSPostUnescape();

    if (array_key_exists("NS", $_REQUEST))
    {
        $NS = Json::decode($_REQUEST["NS"]);

    }
    else
    {
        $NS = array(
            "STEP" => 0,
            "ONLY_DEBT" => $_REQUEST['updateMode'] == "D",
            "UPDATE_MODE" => $_REQUEST['updateMode'] == "Y" || $_REQUEST['updateMode'] == "D",
            "URL_DATA_FILE" => $_REQUEST["URL_DATA_FILE"],
            "LAST_ROW" => $lastRow,
            "LAST_ACCOUNT" => "",
        );

        CUserOptions::SetOption('citrus.tszh.xlsx.import', 'UpdateMode', $NS["UPDATE_MODE"]);

    }

    //We have to strongly check all about file names at server side
    $ABS_FILE_NAME = false;
    $WORK_DIR_NAME = false;

    if (isset($NS["URL_DATA_FILE"]) && (strlen($NS["URL_DATA_FILE"]) > 0))
    {
        $filename = trim(str_replace("\\", "/", $NS["URL_DATA_FILE"]), "/");
        $FILE_NAME = rel2abs($_SERVER["DOCUMENT_ROOT"], "/" . $filename);

        if ((strlen($FILE_NAME) > 1) && ($FILE_NAME === "/" . $filename) && ($APPLICATION->GetFileAccessPermission($FILE_NAME) >= "W"))
        {
            $ABS_FILE_NAME = $_SERVER["DOCUMENT_ROOT"] . $FILE_NAME;
            $WORK_DIR_NAME = substr($ABS_FILE_NAME, 0, strrpos($ABS_FILE_NAME, "/") + 1);
        }

    }

    try
    {
        if ($NS["STEP"] <= 0)
        {
            $NS["STEP"] = 1;
        }

        if (file_exists($ABS_FILE_NAME) && is_file($ABS_FILE_NAME))
        {
            list($lastAccount, $fileName, $xlsxLastRow, $arErrors[], $finishFlag) = CTszhImportXLSX::doImport($NS["LAST_ACCOUNT"], $ABS_FILE_NAME, $NS["LAST_ROW"], $NS["STEP"], $NS["UPDATE_MODE"]);

            $resultArray = Json::encode(
                    array(
                        "LAST_ACCOUNT" => $lastAccount,
                        "FILE_NAME" => $NS["URL_DATA_FILE"],
                        "LAST_ROW" => $xlsxLastRow,
                        "ERROR_MESSAGE" => $arErrors,
                        "FINISH_FLAG" => $finishFlag,
                        "STEP" => ++$NS["STEP"], //увеличиваем шаг импорта
                        "ONLY_DEBT" => $NS["ONLY_DEBT"],
                        "UPDATE_MODE" => $NS["UPDATE_MODE"],
                        "URL_DATA_FILE" => $NS["URL_DATA_FILE"],
                        "PROGRESS" => CTszhImportXLSX::getProgress($ABS_FILE_NAME, $xlsxLastRow)
                        )
            );

            echo $resultArray;

            return;
        }
        else
        {

            $arErrors[] = GetMessage("TI_FILE_NOT_EXSISTS");
        }
    }
    catch(Exception $ex)
    {

        return;
    }

    if (is_array($arErrors) && $arErrors[0] !== "")
    {
        echo Json::encode(array("ERROR_MESSAGE" => $arErrors));
        return;
    }
}

$APPLICATION->SetTitle(GetMessage("TI_TITLE"));

require($_SERVER["DOCUMENT_ROOT"] . "/bitrix/modules/main/include/prolog_admin_after.php");

CUtil::InitJSCore(Array("ajax", "window"));
?>

    <div id="tszh_import_result_div"></div>

    </br>

    <div id="progress"></div>

<?
$aTabs = array(
    array(
        "DIV" => "edit1",
        "TAB" => GetMessage("TI_TAB"),
        "ICON" => "main_user_edit",
        "TITLE" => GetMessage("TI_TAB_TITLE"),
    ),
);
$tabControl = new CAdminTabControl("tabControl", $aTabs);
?>
<script>
    var running = false;
    var oldNS = '';

    function UpdateModeChanged() {
        var bUpdateMode = getCheckedValue(document.form1.updateMode).length > 0;
    }
    function getCheckedValue(radioObj) {
        if (!radioObj)
            return "";
        var radioLength = radioObj.length;
        if (radioLength == undefined)
            if (radioObj.checked)
                return radioObj.value;
            else
                return "";
        for (var i = 0; i < radioLength; i++) {
            if (radioObj[i].checked) {
                return radioObj[i].value;
            }
        }
        return "";
    }

    function DoNext(bFirst, progressBar, NS = []) {
        var queryData = {
            Import: 'Y',
            sessid: '<?=bitrix_sessid()?>'
        };

        if (bFirst) {

            queryData['bFirst'] = 1;
            queryData['URL_DATA_FILE'] = document.getElementById('URL_DATA_FILE').value;

            queryData['updateMode'] = getCheckedValue(document.form1.updateMode);

        } else {
            queryData['NS'] = JSON.stringify(NS);
        }

        var ajaxConfig = {
            method: 'POST', // request method: GET|POST
            dataType: 'json', // type of data loading: html|json|script
            cache: false, // whether NOT to add random addition to URL
            url: 'tszh_import_xlsx.php?lang=<?=LANG?>',
            data: queryData,
            onsuccess: function (result) {
                CloseWaitWindow();

                if (result["ERROR_MESSAGE"][0].trim() !== "")
                {
                    progressBar.setTextAfter(result["ERROR_MESSAGE"][0].trim());
                    progressBar.setColor(BX.UI.ProgressBar.Color.DANGER);

                    EndImport();
                }
                else if (result["FINISH_FLAG"])
                {
                    progressBar.update(result["PROGRESS"]);

                    progressBar.setTextAfter("<?=GetMessage('TI_IMPORT_COMPLETE')?>");
                    progressBar.setColor(BX.UI.ProgressBar.Color.SUCCESS);

                    EndImport();
                }
                else
                {
                    progressBar.update(result["PROGRESS"]);

                    var NS = result;
                    DoNext(false, progressBar, NS);
                }
            }
        };

        ShowWaitWindow();

        if (running)
        {
            if (!BX.ajax(ajaxConfig)) {
                alert("<?=GetMessage("TI_ERROR_AJAX_GET_ERROR")?>");
            }
        }
        else
        {
            progressBar.setTextAfter("<?=GetMessage('TI_IMPORT_STOP')?>");
            CloseWaitWindow();
            EndImport();
        }

    }

    function StartImport() {

        var progressBar = new BX.UI.ProgressBar({
            maxValue: 100,
            value: 0,
            statusType: BX.UI.ProgressBar.Status.PERCENT,
            size: BX.UI.ProgressBar.Size.LARGE,
            fill: true,
            column: true,
            textAfter: "<?=GetMessage('TI_EXECUTE_IMPORT')?>"
        });
        document.getElementById('tszh_import_result_div').innerHTML = "";
        document.getElementById("progress").innerHTML = "";
        document.getElementById("progress").append(progressBar.getContainer());

        running = true;
        document.getElementById('start_button').disabled = true;
        DoNext(true, progressBar);
    }

    function EndImport() {
        running = false;
        document.getElementById('start_button').disabled = false;
    }
</script>
    <form method="POST" action="<? echo $APPLICATION->GetCurPage() ?>?lang=<? echo htmlspecialcharsbx(LANG) ?>"
          name="form1" id="form1">
        <?
        $tabControl->Begin();
        $tabControl->BeginNextTab();
        ?>

        <tr valign="top">
            <td width="40%"><span class="required">*</span><? echo GetMessage("TI_URL_DATA_FILE") ?>:</td>
            <td width="60%">
                <input type="text" id="URL_DATA_FILE" name="URL_DATA_FILE" size="30"
                       value="<?=htmlspecialcharsbx($URL_DATA_FILE)?>">
                <input type="button" value="<? echo GetMessage("TI_OPEN") ?>" OnClick="BtnClick()">
                <?
                CAdminFileDialog::ShowScript
                (
                    Array(
                        "event" => "BtnClick",
                        "arResultDest" => array("FORM_NAME" => "form1", "FORM_ELEMENT_NAME" => "URL_DATA_FILE"),
                        "arPath" => array("SITE" => SITE_ID, "PATH" => "/upload"),
                        "select" => 'F',// F - file only, D - folder only
                        "operation" => 'O',
                        "showUploadTab" => true,
                        "showAddToMenuTab" => false,
                        "fileFilter" => 'xlsx',
                        "allowAllFiles" => true,
                        "SaveConfig" => true,
                    )
                );
                ?>
            </td>
        </tr>
        <tr valign="top">
            <td><?=GetMessage("TI_UPDATE_MODE")?>:</td>
            <td>

                <label><input type="radio" name="updateMode" value="N" onchange="UpdateModeChanged();"
                              onclick="UpdateModeChanged();"<?

                    if (!CUserOptions::GetOption('citrus.tszh.xlsx.import', 'UpdateMode', false))
                    {
                        echo ' checked="checked"';
                    }

                    ?> /><? echo GetMessage("TI_UPDATE_MODE_NO") ?></label><br/>

                <label title="<?=GetMessage("TI_NOTE_1")?>"><input type="radio" name="updateMode" value="Y"
                                                                   onchange="UpdateModeChanged();"
                                                                   onclick="UpdateModeChanged();"<?

                    if (CUserOptions::GetOption('citrus.tszh.xlsx.import', 'UpdateMode', false))
                    {
                        echo ' checked="checked"';
                    }

                    ?> /><? echo GetMessage("TI_UPDATE_MODE_TITLE") ?><span
                        class="required"><sup>1</sup></span></label><br/>

            </td>
        </tr>

        <? $tabControl->Buttons(); ?>
        <input type="button" id="start_button" value="<? echo GetMessage("TI_START_IMPORT") ?>"
               onclick="StartImport();"
               class="adm-btn-save"/>
        <input type="button" id="stop_button" value="<? echo GetMessage("TI_STOP_IMPORT") ?>" onclick="EndImport();"/>
        <? $tabControl->End(); ?>
    </form>

<? echo BeginNote(); ?>
    <span class="required"><sup>1</sup></span> <?=GetMessage('TI_NOTE_1')?><br/>
<? echo EndNote(); ?>

<?require($_SERVER["DOCUMENT_ROOT"] . "/bitrix/modules/main/include/epilog_admin.php");
?>