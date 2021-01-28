<?

use Bitrix\Main\Application;
use Bitrix\Main\Localization\Loc;
use Bitrix\Main\Config\Option;
use Bitrix\Main\Loader;
use Bitrix\Main\ModuleManager;

Loc::LoadMessages(__FILE__);

class citrus_xlsx extends CModule
{
    public $MODULE_ID = "citrus.xlsx";
    public $MODULE_VERSION;
    public $MODULE_VERSION_DATE;
    public $MODULE_NAME;
    public $MODULE_DESCRIPTION;
    public $MODULE_CSS;

    const IBLOCK_TYPE = "owners";
    const TSZH_METTER_ENTITY = "TSZH_METER";
    const TSZH_ACCOUNT_ENTITY = "TSZH_ACCOUNT";

    public $docRoot = '';
    public $iblockID = 0;

    function citrus_xlsx()
    {
        $arModuleVersion = array();

        $this->docRoot = Application::getDocumentRoot();

        include(__DIR__ . "/version.php");

        $this->MODULE_VERSION = $arModuleVersion["VERSION"];
        $this->MODULE_VERSION_DATE = $arModuleVersion["VERSION_DATE"];

        $this->MODULE_NAME = Loc::getMessage("MODULE_NAME");
        $this->MODULE_DESCRIPTION = Loc::getMessage("MODULE_DESCRIPTION");

        $this->PARTNER_NAME = GetMessage("CITRUS_TSZH_MODULE_PARTNER_NAME");
        $this->PARTNER_URI = GetMessage("CITRUS_TSZH_MODULE_PARTNER_URI");


    }

    public function DoInstall()
    {
        global $APPLICATION;

        if(Loader::includeModule("citrus.tszh"))
        {
            $this->InstallDB();
            $this->InstallEvents();

            $this->InstallFiles();

            $APPLICATION->IncludeAdminFile(Loc::getMessage("MODULE_INSTALL_TITLE"), __DIR__ . '/step.php');
        }
        else
        {
            $APPLICATION->ThrowException(Loc::getMessage("INSTALL_ERROR"));
        }

    }

    public function DoUninstall()
    {
        global $APPLICATION;

        $this->UnInstallDB();
        $this->UnInstallEvents();


        $this->UnInstallFiles();

        \Bitrix\Main\ModuleManager::UnRegisterModule($this->MODULE_ID);

        $APPLICATION->IncludeAdminFile(Loc::getMessage("MODULE_UNINSTALL_TITLE"), __DIR__ . '/unstep.php');
    }

    public function InstallFiles()
    {
        CopyDirFiles(__DIR__ . "/admin", $this->docRoot . "/bitrix/admin", true, true);

        return true;
    }

    public function UnInstallFiles()
    {
        DeleteDirFiles(__DIR__ . "/admin", $this->docRoot . "/bitrix/admin");

        return true;
    }

    public function InstallDB()
    {
        ModuleManager::RegisterModule($this->MODULE_ID);
        Loader::includeModule($this->MODULE_ID);

        $this->createIBlock();
        $this->createCustomFields();

        return true;
    }

    public function UnInstallDB()
    {
        $this->deleteIBlock();

        return true;
    }

    public function InstallEvents()
    {

        return true;
    }

    public function UnInstallEvents()
    {
        return true;
    }


    private function createIBlock()
    {
        if (Loader::includeModule('iblock')) {

            $rsSites = \Bitrix\Main\SiteTable::getList();

            while($arSite = $rsSites->fetch())
            {
                $arSites[] = $arSite["LID"];
            }

            $obIBlockType =  new CIBlockType;
            $arFields = Array(
                "ID" => self::IBLOCK_TYPE,
                "SECTIONS" => "Y",
                "LANG" => Array(
                    "ru" => Array(
                        "NAME" => Loc::getMessage("MODULE_IBLOCK_TYPE_NAME"),
                    )
                )
            );
            $res = $obIBlockType->Add($arFields);
            if(!$res)
            {
                $error = $obIBlockType->LAST_ERROR;
            }
            else
            {
                $obIblock = new CIBlock;
                $arFields = Array(
                    "NAME" => Loc::getMessage("MODULE_IBLOCK_NAME"),
                    "CODE" => self::IBLOCK_TYPE,
                    "ACTIVE" => "Y",
                    "IBLOCK_TYPE_ID" => self::IBLOCK_TYPE,
                    "SITE_ID" => $arSites //Массив ID сайтов
                );
                $this->iblockID = $obIblock->Add($arFields);
                if ($this->iblockID)
                {
                    Option::set($this->MODULE_ID, "owners_iblock_id", $this->iblockID);
                }
            }
        }
    }

    private function deleteIBlock()
    {
        global $DB;

        if (Loader::includeModule("iblock"))
        {
            $DB->StartTransaction();

            if(!CIBlockType::Delete(self::IBLOCK_TYPE))
            {
                $DB->Rollback();
                echo 'Delete error!';
            }

            $DB->Commit();
        }

    }

    private function createCustomFields()
    {
        $userTypeEntity = new \CUserTypeEntity();

        $arFields[] = array(
            "ENTITY_ID" => self::TSZH_METTER_ENTITY,
            "FIELD_NAME" => "UF_ROOM_NUMBER",
            "USER_TYPE_ID" => "string",
            "SORT" => 100,
            "MULTIPLE" => "N",
            "MANDATORY" => "N",
            "SHOW_FILTER" => "I",
            "EDIT_FORM_LABEL" => array(
                "ru" => Loc::getMessage("EXPORT_ENTITY_ROOM"),
                "en" => "",
            ),
            "LIST_COLUMN_LABEL" => array(
                "ru" => Loc::getMessage("EXPORT_ENTITY_ROOM"),
                "en" => "",
            )
        );

        $arFields[] = array(
            "ENTITY_ID" => self::TSZH_METTER_ENTITY,
            "FIELD_NAME" => "UF_AXIS",
            "USER_TYPE_ID" => "string",
            "SORT" => 200,
            "MULTIPLE" => "N",
            "MANDATORY" => "N",
            "SHOW_FILTER" => "I",
            "EDIT_FORM_LABEL" => array(
                "ru" => Loc::getMessage("EXPORT_AXIS"),
                "en" => "",
            ),
            "LIST_COLUMN_LABEL" => array(
                "ru" => Loc::getMessage("EXPORT_AXIS"),
                "en" => "",
            )
        );

        $arFields[] = array(
            "ENTITY_ID" => self::TSZH_METTER_ENTITY,
            "FIELD_NAME" => "UF_UNIT",
            "USER_TYPE_ID" => "string",
            "SORT" => 300,
            "MULTIPLE" => "N",
            "MANDATORY" => "N",
            "SHOW_FILTER" => "I",
            "EDIT_FORM_LABEL" => array(
                "ru" => Loc::getMessage("EXPORT_UNIT"),
                "en" => "",
            ),
            "LIST_COLUMN_LABEL" => array(
                "ru" => Loc::getMessage("EXPORT_UNIT"),
                "en" => "",
            )
        );

        $arFields[] = array(
            "ENTITY_ID" => self::TSZH_METTER_ENTITY,
            "FIELD_NAME" => "UF_COEFFICIENT",
            "USER_TYPE_ID" => "string",
            "SORT" => 400,
            "MULTIPLE" => "N",
            "MANDATORY" => "N",
            "SHOW_FILTER" => "I",
            "EDIT_FORM_LABEL" => array(
                "ru" => Loc::getMessage("EXPORT_COEFFICIENT"),
                "en" => "",
            ),
            "LIST_COLUMN_LABEL" => array(
                "ru" => Loc::getMessage("EXPORT_COEFFICIENT"),
                "en" => "",
            )
        );

        $arFields[] = array(
            "ENTITY_ID" => self::TSZH_METTER_ENTITY,
            "FIELD_NAME" => "UF_RENTER",
            "USER_TYPE_ID" => "string",
            "SORT" => 500,
            "MULTIPLE" => "N",
            "MANDATORY" => "N",
            "SHOW_FILTER" => "I",
            "EDIT_FORM_LABEL" => array(
                "ru" => Loc::getMessage("EXPORT_RENTER"),
                "en" => "",
            ),
            "LIST_COLUMN_LABEL" => array(
                "ru" => Loc::getMessage("EXPORT_RENTER"),
                "en" => "",
            )
        );

        $arFields[] = array(
            "ENTITY_ID" => self::TSZH_METTER_ENTITY,
            "FIELD_NAME" => "UF_TYPE",
            "USER_TYPE_ID" => "string",
            "SORT" => 600,
            "MULTIPLE" => "N",
            "MANDATORY" => "N",
            "SHOW_FILTER" => "I",
            "EDIT_FORM_LABEL" => array(
                "ru" => Loc::getMessage("EXPORT_TYPE"),
                "en" => "",
            ),
            "LIST_COLUMN_LABEL" => array(
                "ru" => Loc::getMessage("EXPORT_TYPE"),
                "en" => "",
            )
        );

        $arFields[] = array(
            "ENTITY_ID" => self::TSZH_METTER_ENTITY,
            "FIELD_NAME" => "UF_DATE_INSTALL",
            "USER_TYPE_ID" => "date",
            "SORT" => 700,
            "MULTIPLE" => "N",
            "MANDATORY" => "N",
            "SHOW_FILTER" => "I",
            "EDIT_FORM_LABEL" => array(
                "ru" => Loc::getMessage("EXPORT_DATE_INSTALL"),
                "en" => "",
            ),
            "LIST_COLUMN_LABEL" => array(
                "ru" => Loc::getMessage("EXPORT_DATE_INSTALL"),
                "en" => "",
            )
        );

        $arFields[] = array(
            "ENTITY_ID" => self::TSZH_METTER_ENTITY,
            "FIELD_NAME" => "UF_DATE_UNINSTALL",
            "USER_TYPE_ID" => "date",
            "SORT" => 800,
            "MULTIPLE" => "N",
            "MANDATORY" => "N",
            "SHOW_FILTER" => "I",
            "EDIT_FORM_LABEL" => array(
                "ru" => Loc::getMessage("EXPORT_DATE_UNINSTALL"),
                "en" => "",
            ),
            "LIST_COLUMN_LABEL" => array(
                "ru" => Loc::getMessage("EXPORT_DATE_UNINSTALL"),
                "en" => "",
            )
        );

        $arFields[] = array(
            "ENTITY_ID" => self::TSZH_ACCOUNT_ENTITY,
            "FIELD_NAME" => "UF_OWN_TYPE",
            "USER_TYPE_ID" => "string",
            "SORT" => 100,
            "MULTIPLE" => "N",
            "MANDATORY" => "N",
            "SHOW_FILTER" => "I",
            "EDIT_FORM_LABEL" => array(
                "ru" => Loc::getMessage("EXPORT_OWN_TYPE"),
                "en" => "",
            ),
            "LIST_COLUMN_LABEL" => array(
                "ru" => Loc::getMessage("EXPORT_OWN_TYPE"),
                "en" => "",
            )
        );

        $arFields[] = array(
            "ENTITY_ID" => self::TSZH_ACCOUNT_ENTITY,
            "FIELD_NAME" => "UF_OWNERS",
            "USER_TYPE_ID" => "iblock_element",
            "SORT" => 200,
            "MULTIPLE" => "N",
            "MANDATORY" => "N",
            "SHOW_FILTER" => "I",
            "SETTINGS" => array(
                "IBLOCK_TYPE" => self::IBLOCK_TYPE,
                "IBLOCK_ID" => $this->iblockID,
            ),
            "EDIT_FORM_LABEL" => array(
                "ru" => Loc::getMessage("EXPORT_OWNERS"),
                "en" => "",
            ),
            "LIST_COLUMN_LABEL" => array(
                "ru" => Loc::getMessage("EXPORT_OWNERS"),
                "en" => "",
            )
        );

        foreach ($arFields as $field)
        {
            $userTypeEntity->Add($field);
        }
    }

}
?>

