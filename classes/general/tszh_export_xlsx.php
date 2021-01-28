<? if (!defined("B_PROLOG_INCLUDED") || B_PROLOG_INCLUDED !== true) die();

IncludeModuleLangFile(__FILE__);

use Citrus\Xlsx\XlsxWriter;
use Bitrix\Main\Loader;

/**
 * ��������� �������� ��������� ��������� � ������� XLSX
 */
class CTszhExportXLSX
{
    const TSZH_ACCOUNT_ENTITY = "TSZH_ACCOUNT";
    const TSZH_ACCOUNT_CODE = "UF_OWN_TYPE";
    const TSZH_ACCOUNT_OWNER = "UF_OWNERS";

    static public $limit = 2000;
    /**
     * ����������� �����
     *
     * @param float $number
     * @return string
     */
    public static function num($number, $decPlaces)
    {
        return number_format($number, $decPlaces, ',', ' ');
    }

    /**
     * �������� ��������� ���������
     *
     * @param string $orgName ������������ ����������� (����������� � �����)
     * @param string $orgINN ��� ����������� (����������� � �����)
     * @param string $filename ���������� ���� � �����, ���� ����� �������� ������
     * @param int $tszhOwner ID ������������
     * @param int $rowXLSXIndex - ����� ������, � ������� ���������� ������ � ����
     * @param array $arFilter ���� ������� ��� ������� ���������, ���������� � ���� (@link CTszhMeter::GetList())
     * @param int $step_time ����� ������ ������ � ��������. ���� == 0, �������� ������� �� ���� ���
     * @param int $lastID ID ��������� ������������ ������ �� ���������� ��� (����� ����� �������� ��� ��������� ������)
     * @param array $arValueFilter ���� ������� �� ���������� ��������� (@link CTszhMeterValue::GetList())
     * @return bool|int true ���� �������� ��������� ��������� ��� ID ��������� ������������ ������ (��� �������� ��� ������ �� ��������� ����)
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

        //������� ���������� �������� ��� ������������ ��������� �������� ��������
        $maxColumnNumber = count($arFields);

        // $bNextStep === false �� ������ ����
        $bNextStep = $lastID > 0;
        $limit = self::$limit;

        /// �������� �������� �� ���� ���
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

        // ������ ������ � ������� ��������
        if (!$bNextStep)
        {

            //���������� �������� �������� ������ � ������� � ������� UTF-8
            $arFields = self::setArrayUTF8Encoding($arFields);

            if (!empty($arFields))
            {
                $currentRow = $xlsxWriter->getCurrentRow();

                $xlsxWriter->setTableStyle(1, $currentRow, $maxColumnNumber, $currentRow);
                $xlsxWriter->setHeaderStyle(1, $currentRow, $maxColumnNumber, $currentRow);

                $xlsxWriter->saveIntoXLSX($arFields);

            }
        }

        /// ��������� ������ ���������
        while ($arAccount = $rsAccount->Fetch()){
            $lastID = $arAccount['ID'];
            $users[$arAccount['ID']] = $arAccount;

        }
        unset($rsAccount);

        //��������� ������� ����� �� ������������ ($tszhOwner), ���� �� ������ ������� '��� ������������' ($tszhOwner == 0)
        if ($tszhOwner > 0)
        {
            $users = self::getAccountsByOwner(self::TSZH_ACCOUNT_ENTITY, $users,self::TSZH_ACCOUNT_OWNER, $tszhOwner);
        }

        /// �������� �������� ��� ��������� ���������
        $rsMeters = CTszhMeter::GetList(
            array(),
            array("@ACCOUNT_ID" => array_keys($users), "ACTIVE" => "Y"),
            false,
            false,
            array("ID", "SERVICE_NAME", "XML_ID", "NAME", "VALUES_COUNT", "ACCOUNT_ID", "DEC_PLACES")
        );
        $arMeters = array();

        /// ��������� ������ ���������
        while ($arMeter = $rsMeters->getNext()) $arMeters[$arMeter["ID"]] = $arMeter;
        unset($rsMeters);

        ///�������� ��� ���������������� ����, ��������� � �������� TSZH
        $arCode = self::getCustomFieldsCode(CTszhMeter::USER_FIELD_ENTITY, self::TSZH_ACCOUNT_CODE);

        /// �������� �������� ���������
        $rsMetersValues = CTszhMeterValue::GetList(
            array("TIMESTAMP_X" => "ASC", "ID" => "ASC"),
            array_merge($arValueFilter, Array("@METER_ID" => array_keys($arMeters))),
            false,
            false,
            array("ID", "VALUE1", "TIMESTAMP_X", "METER_ID")
        );

        /// ��������� ������ ��� ��������
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

                //���������� �������� �������� ������ � ������� � ������� UTF-8
                $arMeterValues = self::setArrayUTF8Encoding($arMeterValues);

                ///������ ������� �����
                $currentRow = $xlsxWriter->getCurrentRow();
                $xlsxWriter->setTableStyle(1, $currentRow, $maxColumnNumber, $currentRow);
                /// ������ � ����.
                $xlsxWriter->saveIntoXLSX($arMeterValues);

            }
        }
        unset($rsMetersValues);

        //������ ������� ��������
        $xlsxWriter->setColumnDimension($maxColumnNumber);

        $xlsxWriter->saveFile($filename);

        $lastID = ($last > 0) ? $lastID : true;

        return array($total, $last, $limit, $lastID, $xlsxWriter->getCurrentRow());
    }


    /**
     * ��������� �������� ���������������� �����
     *
     * @param $userEntityObjectName - ������ TSZH, ��� �������� �������� ��������
     * @param $objectId - ID ������
     * @param $arCode - ������ � �������������� �����
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
     * ��������� ���������� ����� ����� ��� �������
     *
     * @param $userEntityObjectName - ������, ���� �������� ����� ��������
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
     * ��������� �������� �������� ������� �� UTF-8
     *
     * @param $arValues - ������, ��������� �������� ���������� ��������
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
     * ��������� ������ ������� ������, � ������� ������������� �������� $tszhOwner
     *
     * @param $tszhObject - ������ TSZH, ���������� ������ ������� ������
     * @param $arAccounts - �������� ������ � �������� �������
     * @param $tszhField - ���������������� ���� - ������������
     * @param $tszhOwner - id ������������
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
     * ��������� ����� ������������ �� ��� ID
     *
     * @param $ownId - ID ������������ � �� "������������"
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