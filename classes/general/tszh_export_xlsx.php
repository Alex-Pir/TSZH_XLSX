<? if (!defined("B_PROLOG_INCLUDED") || B_PROLOG_INCLUDED !== true) die();

IncludeModuleLangFile(__FILE__);

use Citrus\Xlsx\XlsxWriter;
use Bitrix\Main\Loader;

/**
 * Реализует выгрузку показаний счетчиков в формате XLSX
 */
class CTszhExportXLSX
{
    const TSZH_ACCOUNT_ENTITY = "TSZH_ACCOUNT";
    const TSZH_ACCOUNT_CODE = "UF_OWN_TYPE";
    const TSZH_ACCOUNT_OWNER = "UF_OWNERS";

    static public $limit = 2000;
    /**
     * Форматирует число
     *
     * @param float $number
     * @return string
     */
    public static function num($number, $decPlaces)
    {
        return number_format($number, $decPlaces, ',', ' ');
    }

    /**
     * Выгрузка показаний счетчиков
     *
     * @param string $orgName Наименование организации (сохраняется в файле)
     * @param string $orgINN ИНН организации (сохраняется в файле)
     * @param string $filename Абсолютный путь к файлу, куда будут записаны данные
     * @param int $tszhOwner ID Собственника
     * @param int $rowXLSXIndex - номер строки, с которой начинается запись в файл
     * @param array $arFilter Поля фильтра для выборки счетчиков, попадающих в файл (@link CTszhMeter::GetList())
     * @param int $step_time Время работы одного в секундах. Если == 0, выгрузка пройдет за один шаг
     * @param int $lastID ID последней обработанной записи за предыдущий шаг (нужно будет передать при следующем вызове)
     * @param array $arValueFilter Поля фильтра по показаниям счетчиков (@link CTszhMeterValue::GetList())
     * @return bool|int true если выгрузка завершена полностью или ID последней обработанной записи (для передачи при вызове на следующем хите)
     */
    public static function DoExport($orgName, $orgINN, $filename, $tszhOwner, $rowXLSXIndex, $arFilter = Array(), $step_time = 0, $lastID = 0, $arValueFilter = Array())
    {
        global $APPLICATION;

        $startTime = microtime(1);

        @set_time_limit(0);
        @ignore_user_abort(true);

        if (!is_array($arFilter))
        {
            $arFilter = Array();
        }

        $xlsxWriter = new XlsxWriter($filename, $rowXLSXIndex);

        $arFields = Array(
            GetMessage("TSZH_EXPORT_OWN"),
            GetMessage("TSZH_EXPORT_ACCOUNT_OWN"),
            GetMessage("TSZH_EXPORT_MNAME"),
            GetMessage("TSZH_EXPORT_SERVICE"),
            GetMessage("TSZH_EXPORT_MDATE"),
            GetMessage("TSZH_EXPORT_MVAL"),
            GetMessage("TSZH_EXPORT_ROOM_NUMBER"),
            GetMessage("TSZH_EXPORT_AXIS"),
            GetMessage("TSZH_EXPORT_UNIT"),
            GetMessage("TSZH_EXPORT_COEFFICIENT"),
            GetMessage("TSZH_EXPORT_RENTER"),
            GetMessage("TSZH_EXPORT_TYPE"),
            GetMessage("TSZH_EXPORT_DATE_INSTALL"),
            GetMessage("TSZH_EXPORT_DATE_UNINSTALL")
        );

        //считаем количество столбцов для последующего изменения размеров столбцов
        $maxColumnNumber = count($arFields);

        // $bNextStep === false на первом шаге
        $bNextStep = $lastID > 0;
        $limit = self::$limit;

        /// Получаем аккаунты на этот шаг
        $accFilter = array_merge( $arFilter, Array(">ID" => $lastID) );
        $rsAccount = CTszhAccount::GetList(Array("ID" => "ASC"), $accFilter, false, Array( 'nPageSize' => $limit, 'iNumPage' => 1), Array("*"));

        $total = CTszhAccount::GetCount();
        $last = CTszhAccount::GetCount($accFilter) - $limit;
        $last = ($last < 0)?0:$last;

        if (!$bNextStep && $total == 0)
        {
            $APPLICATION->ThrowException(GetMessage("TSZH_ERROR_EXPORT_NO_DATA"), "TSZH_EXPORT_NO_DATA");
            return false;
        }

        // первая строка с именами столбцов
        if (!$bNextStep)
        {

            //Библиотека экспорта работает только с данными в формате UTF-8
            $arFields = self::setArrayUTF8Encoding($arFields);

            if (!empty($arFields))
            {
                $currentRow = $xlsxWriter->getCurrentRow();

                $xlsxWriter->setTableStyle(1, $currentRow, $maxColumnNumber, $currentRow);
                $xlsxWriter->setHeaderStyle(1, $currentRow, $maxColumnNumber, $currentRow);

                $xlsxWriter->saveIntoXLSX($arFields);

            }
        }

        /// Формируем массив аккаунтов
        while ($arAccount = $rsAccount->Fetch()){
            $lastID = $arAccount['ID'];
            $users[$arAccount['ID']] = $arAccount;

        }
        unset($rsAccount);

        //фильтруем лицевые счета по собственнику ($tszhOwner), если не выбран вариант 'Все собственники' ($tszhOwner == 0)
        if ($tszhOwner > 0)
        {
            $users = self::getAccountsByOwner(self::TSZH_ACCOUNT_ENTITY, $users,self::TSZH_ACCOUNT_OWNER, $tszhOwner);
        }

        /// Выбираем счетчики для выбранных аккаунтов
        $rsMeters = CTszhMeter::GetList(
            array(),
            array("@ACCOUNT_ID" => array_keys($users), "ACTIVE" => "Y"),
            false,
            false,
            array("ID", "SERVICE_NAME", "XML_ID", "NAME", "VALUES_COUNT", "ACCOUNT_ID", "DEC_PLACES")
        );
        $arMeters = array();

        /// Формируем массив счетчиков
        while ($arMeter = $rsMeters->getNext()) $arMeters[$arMeter["ID"]] = $arMeter;
        unset($rsMeters);

        ///Получаем все пользовательские поля, связанные с объектом TSZH
        $arCode = self::getCustomFieldsCode(CTszhMeter::USER_FIELD_ENTITY, self::TSZH_ACCOUNT_CODE);

        /// Выбираем значения счетчиков
        $rsMetersValues = CTszhMeterValue::GetList(
            array("TIMESTAMP_X" => "ASC", "ID" => "ASC"),
            array_merge($arValueFilter, Array("@METER_ID" => array_keys($arMeters))),
            false,
            false,
            array("ID", "VALUE1", "TIMESTAMP_X", "METER_ID")
        );

        /// формируем массив для выгрузки
        while ($arValue = $rsMetersValues->fetch())
        {
            if(!isset($arMeters[$arValue["METER_ID"]])) continue;
            $arMeter = $arMeters[$arValue["METER_ID"]];

            foreach ($arMeter["ACCOUNT_ID"] as $accountID)
            {
                $arMeterValues = array();
                if(!isset($users[$accountID])) continue;

                $arMeterValues[] = $users[$accountID]["XML_ID"];

                $arAccountCustomFieldsValues = self::getUserFields(self::TSZH_ACCOUNT_ENTITY, $accountID, array(self::TSZH_ACCOUNT_CODE));
                $arMeterValues = array_merge($arMeterValues, $arAccountCustomFieldsValues);

                $arMeterValues[] = $arMeter['NAME'];
                $arMeterValues[] = $arMeter['SERVICE_NAME'];
                $arMeterValues[] = $arValue['TIMESTAMP_X'];
                $arMeterValues[] = self::num($arValue["VALUE1"], $arValue["DEC_PLACES"]);

                $arMeterCustomFieldsValues = self::getUserFields(CTszhMeter::USER_FIELD_ENTITY, $arValue["METER_ID"], $arCode);
                $arMeterValues = array_merge($arMeterValues, $arMeterCustomFieldsValues);

                //Библиотека экспорта работает только с данными в формате UTF-8
                $arMeterValues = self::setArrayUTF8Encoding($arMeterValues);

                ///задаем границы ячеек
                $currentRow = $xlsxWriter->getCurrentRow();
                $xlsxWriter->setTableStyle(1, $currentRow, $maxColumnNumber, $currentRow);
                /// пакуем в файл.
                $xlsxWriter->saveIntoXLSX($arMeterValues);

            }
        }
        unset($rsMetersValues);

        //задаем размеры столбцов
        $xlsxWriter->setColumnDimension($maxColumnNumber);

        $xlsxWriter->saveFile($filename);

        $lastID = ($last > 0) ? $lastID : true;

        return array($total, $last, $limit, $lastID, $xlsxWriter->getCurrentRow());
    }


    /**
     * Получение значений пользовательских полей
     *
     * @param $userEntityObjectName - объект TSZH, для которого получаем значения
     * @param $objectId - ID записи
     * @param $arCode - массив с наименованиями полей
     * @return mixed
     */
    private static function getUserFields($userEntityObjectName, $objectId, $arCode)
    {

        global $USER_FIELD_MANAGER;

        foreach ($arCode as $code)
        {
            $arResult[] = $USER_FIELD_MANAGER->GetUserFieldValue($userEntityObjectName, $code, $objectId);
        }

        return $arResult;

    }


    /**
     * Получение символьных кодов полей для объекта
     *
     * @param $userEntityObjectName - объект, поля которого нужно получить
     * @return array
     */
    private static function getCustomFieldsCode($userEntityObjectName)
    {
        $arResult = [];
        $arTszhEntity = CUserTypeEntity::GetList(array(), array("ENTITY_ID" => $userEntityObjectName));

        while($arEntity = $arTszhEntity->Fetch())
        {

            $arResult[] = $arEntity["FIELD_NAME"];
        }

        return $arResult;
    }

    /**
     * Изменение кодироки входного массива на UTF-8
     *
     * @param $arValues - массив, кодировку которого необходимо изменить
     * @return array
     */
    private static function setArrayUTF8Encoding($arValues)
    {
        global $APPLICATION;

        $arResult = $arValues;

        if (SITE_CHARSET !== 'utf-8')
        {
            $arResult = $APPLICATION->ConvertCharsetArray($arValues, SITE_CHARSET, 'utf-8');
        }

        return $arResult;
    }

    /**
     * Получение списка лицевых счетов, у которых собственником является $tszhOwner
     *
     * @param $tszhObject - Объект TSZH, содержащий список лицевых счетов
     * @param $arAccounts - исходный массив с лицевыми счетами
     * @param $tszhField - пользовательское поле - собственники
     * @param $tszhOwner - id собственника
     * @return array
     */
    private static function getAccountsByOwner($tszhObject, $arAccounts, $tszhField, $tszhOwner)
    {
        $arResultAccounts = [];

        foreach ($arAccounts as $account)
        {
            $arOwnerId = self::getUserFields($tszhObject, $account["ID"], array($tszhField));

            if ($arOwnerId[0] == $tszhOwner)
            {
                $arResultAccounts[$account["ID"]] = $account;
            }
        }

        return $arResultAccounts;
    }

    /**
     * Получение имени собственника по его ID
     *
     * @param $ownId - ID собственника в ИБ "Собственники"
     * @return string
     */
    private static function getOwnNameById($ownId)
    {
        $name = "";

        if (Loader::includeModule("iblock"))
        {
            $dbOwn = CIBlockElement::GetByID($ownId)->GetNext();
            $name = $dbOwn["NAME"];
        }

        return $name;
    }
}

?>