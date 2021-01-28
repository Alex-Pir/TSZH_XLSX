<? if (!defined("B_PROLOG_INCLUDED") || B_PROLOG_INCLUDED !== true) die();

IncludeModuleLangFile(__FILE__);

use Citrus\Xlsx\XlsxReader;

class CTszhImportXLSX
{

    const LIMIT_COUNT = 500;
    const TSZH_ACCOUNT_ENTITY = "TSZH_ACCOUNT";
    const DATE_START_PERIOD = 8;
    const DATE_END_PERIOD = 9;

    /**
     * Импортирует счетчики из указанного файла (формат файла XLSX)
     *
     * @param $lastAccount - XML_ID последнего обрабатываемого аккаунта
     * @param $fileName - название файла
     * @param $xlsxLastRow - номер строки из предыдущего шага
     * @param $step - номер шага
     * @param false $updatePeriod - флаг, указывающий на то, необходимо ли обновлять показания счетчика (по умолчанию создает новые)
     * @return array
     */
    public static function doImport($lastAccount, $fileName, $xlsxLastRow, $step, $updatePeriod = false)
    {
        $finishFlag = false;

        $errorMessage = "";

        $xlsxReader = new XlsxReader($fileName);

        if (file_exists($fileName) && $xlsxReader->canReadFile())
        {
            try
            {
                $arPeriod = self::getPeriod($xlsxReader);

                $arHeaderIdFields = [
                    "NUMBER" => 0,
                    "OWN" => 1,
                    "OWN_TYPE" => 2,
                    "ROOM_NUMBER" => 3,
                    "AXIS" => 4,
                    "METER_NAME" => 5,
                    "SERVICE" => 6,
                    "UNIT" => 7,
                    "PERIOD_START" => 8,
                    "PERIOD_END" => 9,
                    "DIFFERENCE" => 10,
                    "COEFFICIENT" => 11,
                    "CONSUMPTION" => 12,
                    "RENTER" => 13,
                    "TYPE" => 14,
                    "DATE_INSTALL" => 15,
                    "DATE_UNINSTALL" => 16
                ];

                //получаем реальное количество строк в файле
                $fileRowCount = $xlsxReader->getRowCountFromFile();
                $stepMaxCount = self::LIMIT_COUNT * $step;

                //сравниваем количество строк в файле с максимально возможным на этом шаге
                if ($fileRowCount < $stepMaxCount)
                {
                    $arFile = $xlsxReader->readFileAsArray(count($arHeaderIdFields), $fileRowCount, $xlsxLastRow, $arHeaderIdFields["DATE_INSTALL"] + 1, $arHeaderIdFields["DATE_UNINSTALL"] + 1);
                    $xlsxLastRow = $fileRowCount;
                    $finishFlag = true;
                }
                else
                {
                    $arFile = $xlsxReader->readFileAsArray(count($arHeaderIdFields), $stepMaxCount, $xlsxLastRow, $arHeaderIdFields["DATE_INSTALL"] + 1, $arHeaderIdFields["DATE_UNINSTALL"] + 1);
                    $xlsxLastRow = $stepMaxCount + 1;
                }

                //построчная обработка файла
                foreach ($arFile as $row)
                {
                    //меняем кодировку массива в кодировку windows-1251
                    $row = self::setArrayWindows1251Encoding($row);

                    //проверяем, заполнены ли обязательные ячейки
                    if (self::issetRequiredFields($arHeaderIdFields, $row))
                    {
                        //в случае, если первый столбец не пустой, задаем новое значение для л/с, иначе
                        //оставляем старое
                        if (trim($row[$arHeaderIdFields["OWN"]]) !== '')
                        {
                            $lastAccount = trim($row[$arHeaderIdFields["OWN"]]);
                        }

                        $arAccountsID = self::getAccountByXmlId($lastAccount);


                        foreach ($arAccountsID as $accountId)
                        {
                            if (trim($row[$arHeaderIdFields["OWN_TYPE"]]) !== "")
                            {
                                $arAccountCustomFields = array(
                                    "UF_OWN_TYPE" => $row[$arHeaderIdFields["OWN_TYPE"]]
                                );
                                self::setValueInCustomField(self::TSZH_ACCOUNT_ENTITY, $accountId, $arAccountCustomFields);
                            }

                            $arMeters = self::getMetersByAccountsId($accountId);

                            $arMeterFromFile = self::getMeterByName($row[$arHeaderIdFields["METER_NAME"]], $arMeters);

                            $arMeterCustomFields = array(
                                "UF_ROOM_NUMBER" => $row[$arHeaderIdFields["ROOM_NUMBER"]],
                                "UF_AXIS" => $row[$arHeaderIdFields["AXIS"]],
                                "UF_UNIT" => $row[$arHeaderIdFields["UNIT"]],
                                "UF_COEFFICIENT" => $row[$arHeaderIdFields["COEFFICIENT"]],
                                "UF_RENTER" => $row[$arHeaderIdFields["RENTER"]],
                                "UF_TYPE" => $row[$arHeaderIdFields["TYPE"]],
                                "UF_DATE_INSTALL" => trim($row[$arHeaderIdFields["DATE_INSTALL"]]),
                                "UF_DATE_UNINSTALL" => trim($row[$arHeaderIdFields["DATE_UNINSTALL"]])
                            );

                            $arMeterFields = array(
                                "ACCOUNT_ID" => $accountId,
                                "SERVICE_NAME" => $row[$arHeaderIdFields["SERVICE"]],
                                "MODIFIED_BY" => 1,
                                "NAME" => $row[$arHeaderIdFields["METER_NAME"]]
                            );

                            $arValueStart = [
                                "VALUE1" => $row[$arHeaderIdFields["PERIOD_START"]],
                                "TIMESTAMP_X" => $arPeriod["PERIOD_START"]
                            ];

                            $arValueEnd = [
                                "VALUE1" => $row[$arHeaderIdFields["PERIOD_END"]],
                                "AMOUNT1" => $row[$arHeaderIdFields["CONSUMPTION"]],
                                "TIMESTAMP_X" => $arPeriod["PERIOD_END"]
                            ];

                            if (empty($arMeterFromFile))
                            {
                                $meterId = CTszhMeter::Add($arMeterFields);

                                self::setValueInCustomField(CTszhMeter::USER_FIELD_ENTITY, $meterId, $arMeterCustomFields);

                                $arValueStart["METER_ID"] = $meterId;
                                $arValueEnd["METER_ID"] = $meterId;

                                CTszhMeterValue::Add($arValueStart);
                                CTszhMeterValue::Add($arValueEnd);
                            }
                            else
                            {
                                //проверяем, есть ли хотя бы один счетчик
                                if (!empty($arMeterFromFile[0]))
                                {
                                    foreach ($arMeterFromFile as $meter)
                                    {
                                        $arValueStart["METER_ID"] = $meter["ID"];
                                        $arValueEnd["METER_ID"] = $meter["ID"];

                                        CTszhMeter::Update($meter["ID"], $arMeterFields);

                                        self::setValueInCustomField(CTszhMeter::USER_FIELD_ENTITY, $meter["ID"], $arMeterCustomFields);

                                        //в соответствии с установленной в админке радиокнопкой
                                        if ($updatePeriod)
                                        {
                                            self::updateMetersValues($arValueStart);
                                            self::updateMetersValues($arValueEnd);
                                        }
                                        else
                                        {
                                            CTszhMeterValue::Add($arValueStart);
                                            CTszhMeterValue::Add($arValueEnd);
                                        }

                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception $ex)
            {
                $errorMessage = $ex->getMessage();
                return array($lastAccount, $fileName, $xlsxLastRow, $errorMessage, $finishFlag);
            }

        }
        else
        {
            $errorMessage = GetMessage("CAN_NOT_READ_FILE");
        }

        return array($lastAccount, $fileName, $xlsxLastRow, $errorMessage, $finishFlag);
    }

    /**
     * Получение периода добавляемых показаний счетчика
     *
     * @param $xlsxReader - экземпляр класса XlsxReader (для чтения XLSX-файла)
     * @return array
     */
    private static function getPeriod($xlsxReader)
    {
        $arCurrentTitle = self::setArrayWindows1251Encoding($xlsxReader->readTitle(self::DATE_START_PERIOD + 1, self::DATE_END_PERIOD + 1));



        $arResult = [
            "PERIOD_START" => trim($arCurrentTitle[self::DATE_START_PERIOD]),
            "PERIOD_END" => trim($arCurrentTitle[self::DATE_END_PERIOD])
        ];

        return $arResult;
    }

    /**
     * Устанавливает кодировку Windows-1251 на входной массив
     *
     * @param $arValues - входной массив
     * @return array
     */
    private static function setArrayWindows1251Encoding($arValues)
    {
        global $APPLICATION;

        $arResult = $arValues;

        if (SITE_CHARSET === "windows-1251")
        {
            $arResult = $APPLICATION->ConvertCharsetArray($arValues, "utf-8", 'windows-1251');
        }

        return $arResult;
    }

    /**
     * Получение массива ID лицевых счетов по внешнему коду
     *
     * @param $xmlId - внешний код лицевого счета
     * @return array
     */
    private static function getAccountByXmlId($xmlId)
    {
        $arResult = [];

        $rsAccount = CTszhAccount::GetList(
            array(),
            array("XML_ID" => trim($xmlId), "ACTIVE" => "Y"),
            false,
            false,
            array("ID"));

        while ($arAccount = $rsAccount->Fetch())
        {
            $arResult[] = $arAccount["ID"];
        }

        return $arResult;
    }

    /**
     * Обновление значений пользовательских полей
     *
     * @param $userEntityObjectName - объект, к которому привязаны пользовательские поля
     * @param $id - id поля
     * @param $arFields - массив со значениями, добавляемыми в поле
     */
    private static function setValueInCustomField($userEntityObjectName, $id, $arFields)
    {
        global $USER_FIELD_MANAGER;

        $USER_FIELD_MANAGER->Update($userEntityObjectName, $id, $arFields);
    }

    /**
     * Получение массива ID счетчиков ID лицевого счета
     *
     * @param $accountsId - ID лицевого счета
     * @return array
     */
    private static function getMetersByAccountsId($accountsId)
    {
        $arResult = [];

        $rsMeters = CTszhMeter::GetList(
            array(),
            array("ACCOUNT_ID" => $accountsId),
            false,
            false,
            array("ID", "NAME")
        );

        while ($arMeters = $rsMeters->Fetch())
        {
            $arResult[] = $arMeters;
        }

        return $arResult;
    }

    /**
     * Получение счетчика по его названию
     *
     * @param $meterName - Название счетчика
     * @param $arMeters - Массив с данными счетчиков
     * @return array
     */
    private static function getMeterByName($meterName, $arMeters)
    {
        $arResult = [];

        foreach ($arMeters as $meter)
        {
            if (trim($meter["NAME"]) === trim($meterName))
            {
                $arResult[] = $meter;
            }
        }

        return $arResult;

    }

    /**
     * Обновление значений счетчиков
     * Если счетчик не найден, то он добавляется
     *
     * @param $arFields - значения обновляемых полей
     */
    private static function updateMetersValues($arFields)
    {
        $rsMeterValue = CTszhMeterValue::GetList(
            array(),
            array(
                "ACTIVE" => "Y",
                "METER_ID" => $arFields["METER_ID"],
                "TIMESTAMP_X" => $arFields["TIMESTAMP_X"]
                ),
            false,
            false,
            array("ID")
        );

        if ($arMeterValue = $rsMeterValue->Fetch())
        {
            CTszhMeterValue::Update($arMeterValue["ID"], $arFields);
        }
        else
        {
            CTszhMeterValue::Add($arFields);
        }
    }

    /**
     * Проверяет, заполнены ли обязательные ячейки в строке.
     * Обязательными ячейками являются: "№ Счетчика", "Вид услуги", "Единица измерения", Дата показаний на начало периода, Дата показаний на конец периода, "Коэффициент измерения"
     *
     * @param $arHeaderIdFields - ассоциативный массив, в котором значениями являются номера столбцов в получаемой строке $row
     * @param $row - проверяемая строка
     * @return bool
     */
    private static function issetRequiredFields($arHeaderIdFields, $row)
    {
        $result = false;

        if (
            !empty(trim($row[$arHeaderIdFields["METER_NAME"]])) &&
            !empty(trim($row[$arHeaderIdFields["SERVICE"]])) &&
            !empty(trim($row[$arHeaderIdFields["UNIT"]]))  &&
            !empty(trim($row[$arHeaderIdFields["PERIOD_START"]])) &&
            !empty(trim($row[$arHeaderIdFields["PERIOD_END"]]))
        )
        {
            $result = true;
        }

        return $result;
    }

    /**
     * Возвращает прогресс чтения файла относительно текущей строки
     *
     * @param $fileName - название файла
     * @param $currentRow - номер строки
     * @return int
     */
    public static function getProgress($fileName, $currentRow)
    {
        $result = 0;
        $xlsxReader = new XlsxReader($fileName);

        if (file_exists($fileName) && $xlsxReader->canReadFile())
        {
            $fileRowCount = $xlsxReader->getRowCountFromFile();
            $result = round($currentRow / $fileRowCount * 100);
        }

        return $result;
    }

}