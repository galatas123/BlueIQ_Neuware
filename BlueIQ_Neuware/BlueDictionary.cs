﻿using System.Collections.ObjectModel;
using System.Security;

namespace BlueIQ_Neuware
{
    internal class BlueDictionary
    {
        public const int WEIGHT = 1;
        public const string ASSET = "NA";
        public const string loading = "//*[@id=\"ctl00_ctl00_imgWorking\"]";

        public static readonly Dictionary<string, string> LINKS = new()
        {
            { "LOGIN", "https://blueiq.cloudblue.com/" },
            { "AUDIT", "https://blueiq.cloudblue.com/Audit/PreSort/PreSortInfoCapture.aspx" },
            { "Q&ORDERS", "https://blueiq.cloudblue.com/Remarketing/QuoteOrder.aspx" },
            { "ORDER_FULLFILLMENT", "https://blueiq.cloudblue.com/Remarketing/OrderFulfillment.aspx" },
            { "ORDER_DETAIL", "https://blueiq.cloudblue.com/Remarketing/OrderDetail.aspx?oid=" },
            { "REPAIR_DETAIL", "https://blueiq.cloudblue.com/Audit/RepairDetailLOB.aspx?Scanid=" },
            { "JOBS", "https://blueiq.cloudblue.com/Job1/Jobs.aspx?searchby=Job&searchvalue=" },
            { "LOAD", "https://blueiq.cloudblue.com/Receiving/LoadPalletDetails.aspx?loadid="},
            { "RECEIVING", "https://blueiq.cloudblue.com/Receiving/LoadPalletDetails.aspx" },
            { "MASS_MOVE", "https://blueiq.cloudblue.com/Audit/MassMove.aspx" },
            { "QUARANTINE", "https://blueiq.cloudblue.com/Inventory/Quarantine.aspx" }
        };

        public static readonly Dictionary<string, string> LOGIN_PAGE = new()
        {
            { "USERNAME", "//*[@id=\"ctl00_MainContent_Login1_UserName\"]" },
            { "PASSWORD", "//*[@id=\"ctl00_MainContent_Login1_Password\"]" },
            { "BUTTON", "//*[@id=\"ctl00_MainContent_Login1_LoginButton\"]" },
            { "LOADING", "//*[@id=\"ctl00_imgWorking\"]" },
            { "LOGIN_ERROR", "/html/body/form/div[3]/div[4]/div/div/div/div/div[1]/table/tbody/tr/td/div[1]" },
            { "LOCATION_POPUP", "//*[@id=\"ctl00_MainContent_ASPxPopupControlErrors_btnSubmit\"]" }
        };

        public static readonly Dictionary<string, string> AUDIT_PAGE = new()
        {
            { "LOADING", "//*[@id=\"ctl00_ctl00_imgWorking\"]" },
            { "PALLET_ID", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_txtPalletID\"]" },
            { "UPDATE_COMP", "ctl00_ctl00_MainContent_PageMainContent_lnkUpdate" }, //By ID
            { "DUNNAGE_YES", "ctl00_ctl00_MainContent_PageMainContent_btnCreateDunnageScanYes"}, //By ID close second tab after
            { "SORTED_YES", "ctl00_ctl00_MainContent_PageMainContent_btnLoadPreSortOk"}, //By ID
            { "LOCK_PALLET", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_chkLockPallet\"]" },
            { "PART_NUMBER", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_ddlPartNumberSerialized_I\"]" },
            { "PART_NUMBER_BTN", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_ddlPartNumberSerialized_B-1\"]" },
            { "PART_TABLE", "ctl00_ctl00_MainContent_PageMainContent_ddlPartNumberSerialized_I" }, //ID,
            { "SERIAL#", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_txtSerialNumber_I\"]" },
            { "ASSET", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_txtAssetNumber\"]" },
            { "WEIGHT", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_txtWeight\"]" },
            { "LOCATION", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_txtLocation\"]" },
            { "LOCK_LOCATION", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_chkLockLocation\"]" },
            { "WARRANTY", "/html/body/form/div[3]/div[6]/div/div[2]/div/div/div[1]/div[3]/div[8]/div/div[6]/div[2]/select/option[2]" },
            { "APPROVED", "/html/body/form/div[3]/div[6]/div/div[2]/div/div/div[1]/div[3]/div[13]/div/div/table/tbody/tr/td/table/tbody/tr[4]/td[2]/div/div/select/option[4]" },
            { "SERVICE_TYPE_CREDIT", "/html/body/form/div[3]/div[6]/div/div[2]/div/div/div[1]/div[3]/div[13]/div/div/table/tbody/tr/td/table/tbody/tr[2]/td[2]/div/div[1]/select/option[4]" },
            { "SERVICE_TYPE_SWAP", "/html/body/form/div[3]/div[6]/div/div[2]/div/div/div[1]/div[3]/div[13]/div/div/table/tbody/tr/td/table/tbody/tr[2]/td[2]/div/div[1]/select/option[3]" },
            { "NEW_IN_BOX", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_chkNewinBox\"]" },
            { "PONO", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_dynamicUFDFields_dlDynamicCtrls_ctl03_txt9730\"]" },
            { "SAVE", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_btnSaveNew\"]" },
            { "SEARCH_ASSETS", "//*[@id=\"ctl00_ctl00_contentLeftWrapper_contentLeftMenu_txtSearchText\"]" },
            { "SERIAL#_RADIO", "//*[@id=\"ctl00_ctl00_contentLeftWrapper_contentLeftMenu_rdoSerialNumber\"]" },
            { "NEW_STOCK", "/html/body/form/div[3]/div[6]/div/div[2]/div/div/div[1]/div[3]/div[13]/div/div/table/tbody/tr/td/table/tbody/tr[2]/td[2]/div/div[1]/select/option[8]" }
        };

        public static readonly Dictionary<string, string> Q_ORDERS_PAGE = new()
        {
            { "LOADING", "//*[@id=\"ctl00_ctl00_imgWorking\"]" },
            { "SEARCH", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_txtSearch\"]" },
            { "SEARCH_BUTTON", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_btnSearch\"]" },
            { "NAME", "/html/body/form/div[3]/div[6]/div/div[2]/div/div/div/contenttemplate/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[4]/td/table/tbody/tr/td/div[1]/table/tbody/tr[4]/td[3]" },
            { "PO", "/html/body/form/div[3]/div[6]/div/div[2]/div/div/div/contenttemplate/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[4]/td/table/tbody/tr/td/div[1]/table/tbody/tr[4]/td[9]" },
            { "SO_DATE", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_gvQuotes_tccell0_3\"]" },
            { "EXPORT_EXCEL", "ctl00_ctl00_MainContent_PageMainContent_ImgExcel" } //ID btn
        };

        public static readonly Dictionary<string, string> ORDER_DETAILS_PAGE = new()
        {
            { "RECLAIM", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_pcCustomer_T10T\"]" },
            { "TRACKING", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_pcCustomer_T9\"]" },
            { "PART_NO", "/html/body/form/div[3]/div[6]/div/div[2]/div/div/div[1]/div/contenttemplate/table/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr[2]/td/table[2]/tbody/tr[1]/td/table/tbody/tr[3]/td/div[1]/div/div[11]/div/table/tbody/tr/td/table[1]/tbody/tr[2]/td[2]" },
            { "SERIAL_NO", "/html/body/form/div[3]/div[6]/div/div[2]/div/div/div[1]/div/contenttemplate/table/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr[2]/td/table[2]/tbody/tr[1]/td/table/tbody/tr[3]/td/div[1]/div/div[11]/div/table/tbody/tr/td/table[1]/tbody/tr[2]/td[3]" },
            { "TRACKING_IN", "/html/body/form/div[3]/div[6]/div/div[2]/div/div/div[1]/div/contenttemplate/table/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr[2]/td/table[2]/tbody/tr[1]/td/table/tbody/tr[3]/td/div[1]/div/div[11]/div/table/tbody/tr/td/table[1]/tbody/tr[2]/td[4]" },
            { "SHIP_METHOD_IN", "/html/body/form/div[3]/div[6]/div/div[2]/div/div/div[1]/div/contenttemplate/table/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr[2]/td/table[2]/tbody/tr[1]/td/table/tbody/tr[3]/td/div[1]/div/div[11]/div/table/tbody/tr/td/table[1]/tbody/tr[2]/td[6]" },
            { "CARRIER_IN", "/html/body/form/div[3]/div[6]/div/div[2]/div/div/div[1]/div/contenttemplate/table/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr[2]/td/table[2]/tbody/tr[1]/td/table/tbody/tr[3]/td/div[1]/div/div[11]/div/table/tbody/tr/td/table[1]/tbody/tr[2]/td[7]" },
            { "TRACKING_OUT", "/html/body/form/div[3]/div[6]/div/div[2]/div/div/div[1]/div/contenttemplate/table/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr[2]/td/table[2]/tbody/tr[1]/td/table/tbody/tr[3]/td/div[1]/div/div[10]/div[1]/table/tbody/tr/td/table[1]/tbody/tr[2]/td[8]" },
            { "SHIPPED_OUT", "/html/body/form/div[3]/div[6]/div/div[2]/div/div/div[1]/div/contenttemplate/table/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr[2]/td/table[2]/tbody/tr[1]/td/table/tbody/tr[3]/td/div[1]/div/div[10]/div[1]/table/tbody/tr/td/table[1]/tbody/tr[2]/td[10]" },
            { "CARRIER_OUT", "/html/body/form/div[3]/div[6]/div/div[2]/div/div/div[1]/div/contenttemplate/table/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr[2]/td/table[2]/tbody/tr[1]/td/table/tbody/tr[3]/td/div[1]/div/div[10]/div[1]/table/tbody/tr/td/table[1]/tbody/tr[2]/td[14]" },
            { "COPY_ORDER", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_lnkCopyOrder\"]" },
            { "CONFIRM1_YES", "ctl00_ctl00_MainContent_PageMainContent_btnCopyYes" },  // id
            { "NEW_ORDER_ID", "ctl00_ctl00_MainContent_PageMainContent_hdnNewOrderId" },  // id
            { "NAVIGATE_NO", "ctl00_ctl00_MainContent_PageMainContent_btnNewOrderNo" },  // id
            { "SHIP_TO", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_pcCustomer_imgShippingANDBilling1\"]" },
            { "COMPANY_DROPDOWN", "ctl00_ctl00_MainContent_PageMainContent_popupAddressEdit_ucASPxPanel1_CustomerShiipingBilling_ddlCustSource_I" },
            { "FLENSBURG_OPTION", "ctl00_ctl00_MainContent_PageMainContent_popupAddressEdit_ucASPxPanel1_CustomerShiipingBilling_ddlCustSource_DDD_L_LBI0T0" },
            { "LOADING_SHIPTO", "ctl00_ctl00_imgWorking" },  // id
            { "LOADING_SHIPTO2", "//*[@id=\"ctl00_ctl00_imgWorking\"]" },
            { "SHIP_SAVE", "ctl00_ctl00_MainContent_PageMainContent_popupAddressEdit_ucASPxPanel1_btnShippingSave_CD" },  // ID
            { "FRAME_ADDRESS", "ctl00_ctl00_MainContent_PageMainContent_popupAddress_CIF-1" },
            { "ADDRESS_CHECKBOX", "GrdValidatedAddressDetails_header1_chkUserAddressAll" },  // ID
            { "SERVICE_TYPE", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_ddlServiceType_B-1Img\"]" },
            { "SWAP", "ctl00_ctl00_MainContent_PageMainContent_ddlServiceType_DDD_L_LBI2T0" },
            { "RETURN_TO_SOURCE", "/html/body/form/div[3]/div[6]/div/div[2]/div/div/div[1]/div/contenttemplate/table/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr[2]/td/table[2]/tbody/tr[1]/td/table/tbody/tr[1]/td/table/tbody[1]/tr[3]/td[2]/select/option[25]" },
            { "INCLUDE_LABEL_CHECKBOX", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_chkIncludeReturnlabel\"]" },
            { "SUBMIT_ORDER", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_lnkSubmitOrder\"]" },
            { "CONFIRM_SUBMIT", "ctl00_ctl00_MainContent_PageMainContent_popupMessageBoxForOrderSubmit_ASPxPanel6_btnSubmitYes" }, // ID
            { "SAVE_ORDER", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_btnUpdate\"]" },
            { "SWAP_ORDER_REF", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_pcCustomer_txtSwapOrderRef\"]" },
            { "ADD_SCAN", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_lnkAddScanID\"]" },
            { "INPUT_SCAN", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_txtAddScanID\"]" },
            { "ADD", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_btnAddScanID\"]" },
            { "CANCEL", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_btnAddScanIDCancel\"]" },
            { "MESSAGE", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_lblAddScanIDErrorMessage\"]" },
            { "RECLAIM_LABEL", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_gvDetails_DXSelBtn1_D\"]" },
            { "DELETE", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_lnkDelete\"]" },
            { "CONFIRM_DEL", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_btnDeletsSelectedItems\"]" },
            { "ORDER_DATE", "ctl00_ctl00_MainContent_PageMainContent_lblSODate"} //ID text
        };

        public static readonly Dictionary<string, string> REPAIR_DETAIL_PAGE = new()
        {
            { "SERIAL_NO", "//*[@id=\"lblSerialNumber\"]" },
            { "SERIAL_NO_TEXT", "//*[@id=\"txtSerialNumber\"]" },
            { "SCAN_ID", "//*[@id=\"lblCurrentScanID\"]" },
            { "RETURN_INFORMATION", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_btnViewReturnInformation\"]" },
            { "REPAIR_INPUT", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_txtNotes\"]" },
            { "REPAIR_NOTES", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_txtRepairNotesHistory_I\"]" },
            { "SITE", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_tabSource_PCAssetDetails_ddlSite\"]" },
            { "ORDER_ID", "/html/body/form/div[3]/div[2]/div/div/div/div/div[2]/div[5]/table/tbody/tr/td/div[2]/table/tbody/tr[2]/td[1]" },
            { "CLOSE_POPUP", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_popupViewReturnInfo_HCB-1\"]" },
            { "LOADING", "//*[@id=\"ctl00_ctl00_imgWorking\"]" },
            { "DIAG_TABLE", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_tabSource_PCAssetDetails_ddlRepairCodes_B-1\"]" },
            { "DIAG", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_tabSource_PCAssetDetails_ddlRepairCodes_DDD_DDTC_lstRepairCodes_28137_D\"]" },
            { "SAVE", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_btnSave\"]" }
        };

        public static readonly Dictionary<string, string> RECEIVING_PAGE = new()
        {
            { "JOB_ID", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_txtJobID\"]" },
            { "PONO", "ctl00_ctl00_MainContent_PageMainContent_txtCustomerPO"}, //ID
            { "SEARCH_JOB", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_txtJobID\"]" },
            { "SEARCH_BTN", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_lnkShowJobs\"]" },
            { "PALLET_ID", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_gvLoadPallet_ctl02_lblPallet\"]" },
            { "LOAD_ID", "ctl00_ctl00_MainContent_PageMainContent_lblCurrentLoadID" }, //ID
            { "OTHER", "/html/body/form/div[3]/div[6]/div/div[2]/div/div/div[1]/div[1]/div[2]/div[2]/div/div/div[1]/div[2]/select/option[6]" },
            { "LOADING", "//*[@id=\"ctl00_ctl00_UpdateProgress1\"]" },
            { "BOL", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_txtBOL\"]" },
            { "QTY_PALLET", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_txtPalletUnits\"]" },
            { "DATE", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_txtDocDate\"]" },
            { "ARRIVAL_TIME", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_txtDocTime\"]" },
            { "SAVE&EXIT", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_btnSaveNxt\"]" },
            { "PENCIL", "//*[@id=\"ctl00_ctl00_MainContent_PageMainContent_gvLoadPallet_ctl02_imgEdit\"]" },
            { "WEIGHT", "//*[@id=\"txtWeight\"]" },
            { "DUNNAGE_PALLET", "/html/body/form/div[4]/div/div[2]/div[1]/div[2]/div[2]/select/option[2]" },
            { "AUDIO", "ddlClass" }, //ID
            { "MISCELLANEOUS", "ddlCategory" }, //ID
            { "QTY_DEVICES", "//*[@id=\"txtQuantity\"]" },
            { "SORT", "ddlMoveToSite" }, //ID
            { "LOCATION", "//*[@id=\"txtMoveToLocation\"]" },
            { "SAVE", "/html/body/form/div[4]/div/div[1]/div[2]/div[2]/div[3]/div/a/img" },
            { "PRINT", "//*[@id=\"btnPrint\"]" }
        };

        //all by id from here on
        public static readonly Dictionary<string, string> JOBS_PAGE = new()
        {
            { "ATTACHMENT_TAB", "__tab_ctl00_ctl00_MainContent_PageMainContent_tabJobs_tabAttachments" },
            { "LOAD_TAB", "__tab_ctl00_ctl00_MainContent_PageMainContent_tabJobs_tabLoadDetails"},
            { "SCHEDULES_TAB", "__tab_ctl00_ctl00_MainContent_PageMainContent_tabJobs_tabSchedules"},
            { "JOB_INFO_TAB", "__tab_ctl00_ctl00_MainContent_PageMainContent_tabJobs_tabSource"},
            { "JOB_STATUS", "ctl00_ctl00_MainContent_PageMainContent_tabJobs_tabSource_ddlJobStatus"}, //SELECT Processing Completed
            { "DOCUMENT", "ctl00_ctl00_MainContent_PageMainContent_tabJobs_tabAttachments_ucDocuments_gvDocumentsList_ctl02_lnk_downloadDocument" },
            { "CARRIER", "ctl00_ctl00_MainContent_PageMainContent_tabJobs_tabSchedules_cbpCustomCarrier_ddlCarrier"}, //SELECT "Other"
            { "SAVE", "ctl00_ctl00_MainContent_PageMainContent_btnSave"},
            { "REASON_NOTE", "TxtsaveJobReason"},
            { "REASON", "Updating to processing completed"},
            { "REASON_OK", "ctl00_ctl00_MainContent_PageMainContent_ASPxPopupNotesReason_btnSaveJobreason"}
            
        };

        public static readonly Dictionary<string, string> LOAD_PAGE = new()
        {
            { "LOAD_STATUS", "ctl00_ctl00_contentLeftWrapper_contentLeftMenu_lnkLoadStatus"},
            { "AUDIT_COMPLETED", "ctl00_ctl00_MainContent_PageMainContent_chkAuditCompleted"}, //Checkbox
            { "PRESORT_COMPLETED", "ctl00_ctl00_MainContent_PageMainContent_chkSortCompleted"}, //Checkbox
            { "OPS_COMPLETED", "ctl00_ctl00_MainContent_PageMainContent_chkOpsProcessingCompleteDate"}, //Checkbox
            { "SAVE", "ctl00_ctl00_MainContent_PageMainContent_btnSave"}
        };

        public static readonly Dictionary<string, string> MASS_MOVE_PAGE = new()
        {
            { "FROM_LOCATION", "ctl00_ctl00_MainContent_PageMainContent_txtFromLocation" },
            { "TO_LOCATION", "ctl00_ctl00_MainContent_PageMainContent_txtMoveToLocation" },
            { "FROM_SITE_SEL", "ctl00_ctl00_MainContent_PageMainContent_ddlFromSite" },
            { "TO_SITE_SEL", "ctl00_ctl00_MainContent_PageMainContent_ddlMoveToSite" },
            { "MOVE_BTN", "ctl00_ctl00_MainContent_PageMainContent_btnMove" },
            { "SUBMIT", "ctl00_ctl00_MainContent_PageMainContent_btnOk" },
            { "QUAR_REASON_SEL", "ctl00_ctl00_MainContent_PageMainContent_popupQuarantine_ASPxPanel2_ddlQuarentineReason_B-1Img" },
            { "QUAR_REASON_OTHER", "ctl00_ctl00_MainContent_PageMainContent_popupQuarantine_ASPxPanel2_ddlQuarentineReason_DDD_L_LBI23T0" },
            { "QUAR_COMMENT", "ctl00_ctl00_MainContent_PageMainContent_popupQuarantine_ASPxPanel2_txtQuarantineReason_I" },
            { "MOVE_BTN_POP", "ctl00_ctl00_MainContent_PageMainContent_popupQuarantine_ASPxPanel2_btnQuarantineReason_CD" },
            { "LOADING", "ctl00_ctl00_imgWorking" }
        };

        public static readonly Dictionary<string, string> QUARANTINE_PAGE = new()
        {
            { "JOB#", "ctl00_ctl00_MainContent_PageMainContent_txtjob" },
            { "SEARCH_BTN", "ctl00_ctl00_MainContent_PageMainContent_btnSearch" },
            { "DATA_ROW_1", "ctl00_ctl00_MainContent_PageMainContent_gvQuarantineAssets_DXDataRow0" },
            { "SELECT_ALL", "ctl00_ctl00_MainContent_PageMainContent_gvQuarantineAssets_header0_SelectAllCheckBox_S" },
            { "RELEASE_BTN", "ctl00_ctl00_MainContent_PageMainContent_btnReleaseQassets" },
            { "RELEASE_REASON", "ctl00_ctl00_MainContent_PageMainContent_popupRelease_ASPxPanel2_txtQuarantineReason_I" },
            { "RELEASE_YES", "ctl00_ctl00_MainContent_PageMainContent_popupRelease_ASPxPanel2_btnReleaseYes" }
        };

        public static readonly Dictionary<string, string> LOCATIONS = new()
        {
            {"QUARANTINE", "L1-68.01" }
        };

    }
}