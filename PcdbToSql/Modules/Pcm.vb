Module Pcm

    ' Pcdb Constants Module

#Region " Edit Masks "

    Public Const MaskSeparator As String = "|"

#End Region

#Region " Message Constants "

    Public Const SC_CLOSE As Integer = &HF060&
    Public Const WM_SYSCOMMAND As Integer = &H112

#End Region

#Region " Dates and Date Ranges "

    Public Const dtFirstDate As Date = #1/1/1980#
    Public Const NoDate As Date = #12:00:00 AM#

    Public Const dfShortDay As String = "d"
    Public Const dfShortDayFullMonth As String = dfShortDay & " " & dfFullMonth
    Public Const dfShortDayFullMonthFullYear As String = dfShortDay & " " & dfFullMonth & " " & dfFourDigitYear
    Public Const dfFourDigitYear As String = "yyyy"
    Public Const dfFullDate As String = Pcm.dfFullMonth & " dd, " & Pcm.dfFourDigitYear
    Public Const dfFullMonth As String = "MMMM"
    Public Const dfFullMonthYear As String = dfFullMonth & " " & dfFourDigitYear
    Public Const dfShortDate As String = "MM/dd/yyyy"
    Public Const dfShortMonthName As String = "MMM"
    Public Const dfShortMonthNameFullYear As String = dfShortMonthName & " " & dfFourDigitYear

    ' date range indexes
    Public Const driNotSet As Integer = -1
    Public Const driAllDates As Integer = 0
    Public Const driLastMonth As Integer = 1
    Public Const driLast2Months As Integer = 2
    Public Const driLast3Months As Integer = 3
    Public Const driLast6Months As Integer = 4
    Public Const driLastYear As Integer = 5
    Public Const driLast12Months As Integer = 6
    Public Const driSpecificMonth As Integer = 7
    Public Const driThisMonth As Integer = 8
    Public Const driThisYear As Integer = 9
    Public Const driCustom As Integer = 10

    Public Const driMinimum As Integer = -1
    Public Const driMaximum As Integer = 10

    ' date grouping indexes
    Public Const dgiDate As Integer = 0
    Public Const dgiMonth As Integer = 1
    Public Const dgiYear As Integer = 2
    Public Const dgiDateRange As Integer = 3

    ' date grouping (with week) indexes
    Public Const dgwiDate As Integer = 0
    Public Const dgwiWeek As Integer = 1
    Public Const dgwiMonth As Integer = 2
    Public Const dgwiYear As Integer = 3
    Public Const dgwiDateRange As Integer = 4

#End Region

#Region " Decimal (Money) / Rounding "

    Public Const NoMoney As Decimal = -1D
    Public Const ZeroMoney As Decimal = 0D

    Public Const RoundToCents As Integer = 2
    Public Const FfPercentRoundTo As Integer = 7

    Public Const DollarFormat As String = "$#,##0.00;($#,##0.00)"

#End Region

#Region " Tool Bar Button Indexes "

    Public Const biNotSet As Integer = -1

    Public Const biProjects As Integer = 0
    Public Const biCompanies As Integer = 1
    Public Const biContacts As Integer = 2

    Public Const biJobs As Integer = 4
    Public Const biInvoices As Integer = 5
    Public Const biPayments As Integer = 6
    Public Const biWorkDone As Integer = 7
    Public Const biTimeCard As Integer = 8

    Public Const biDelivery As Integer = 10
    Public Const biTravel As Integer = 11

    Public Const biPrint As Integer = 13
    Public Const biPreview As Integer = 14

    Public Const biUndo As Integer = 16
    Public Const biCut As Integer = 17
    Public Const biCopy As Integer = 18
    Public Const biPaste As Integer = 19

    Public Const biCopyAddress As Integer = 21
    Public Const biPhone As Integer = 22
    Public Const biAdvFilter As Integer = 23

    Public Const biMerge As Integer = 25

    Public Const biExit As Integer = 27
#End Region

#Region " Edit Tool Bar "

    Public Const EditCut As Boolean = True
    Public Const EditCopy As Boolean = False

#End Region

#Region " Control Navigator Button Indexes/Tag Values "

    'Public Const cnEdit As Integer = 0
    'Public Const cnEndEdit As Integer = 1
    'Public Const cnCancelEdit As Integer = 2
    'Public Const cnRefresh As Integer = 3
    'Public Const cnAppend As Integer = 4
    'Public Const cnRemove As Integer = 5
    'Public Const cnReviseAll As Integer = 6
    'Public Const cnMessage As Integer = 7
    'Public Const cnUpdateInv As Integer = 8
    'Public Const cnView As Integer = 9

#End Region

#Region " Data Selection Label Strings "

    Public Const dsParams As String = "Data Selection Parameters"
    Public Const dsEdited As String = "Data has been edited, selection not valid"

#End Region

    '****************
    ' ERROR VALUES
    '****************

#Region " Error Values "

#Region " General Errors "

    Public Const emInvalidEntry As String = "Invalid Entry"

    Public Const NoErrors As Integer = 0
    Public Const OtherExErr As Integer = -4000

#End Region

#Region " Table Errors -1 "

    ' all table errors are < 0 because table fills/updates return
    ' the number of rows filled/updates
    Public Const teNoErrorsRowHandle As Integer = -1
    Public Const teNoTableErr As Integer = -2
    Public Const teNoOleDbErr As Integer = -3
    Public Const teNoDataSetErr As Integer = -4
    Public Const teLoadErr As Integer = -5
    Public Const teConcurrencyErr As Integer = -6
    Public Const teUpdateErr As Integer = -7
    Public Const teDeleteErr As Integer = -8
    Public Const teConstraintErr As Integer = -9
    Public Const teUserCancelErr As Integer = -10
    Public Const teInAnotherDataSetErr As Integer = -11
    Public Const teCouldNotCopyChildRows As Integer = -12
    Public Const teNoDataInTable As Integer = -13
    Public Const teConnectionNotClosed As Integer = -14
    Public Const teSetLinkValueErr As Integer = -15
    Public Const teNoBindingManagerBase As Integer = -16
    Public Const teDataErr As Integer = -17
    Public Const teEmptyTableErr As Integer = -18
    Public Const teWrongRowCountErr As Integer = -19
    Public Const teAppendTableErr As Integer = -20
    Public Const teSelectErr As Integer = -21
    Public Const teSqlError As Integer = -22
    Public Const teParametersErr As Integer = -23
    Public Const tePrimaryKeyErr As Integer = -24
    Public Const teNoColumnErr As Integer = -25
    Public Const teCouldNotClearHasChildRowsErr As Integer = -26
    Public Const teNoDataPathErr As Integer = -27
    Public Const teDataPathNoFoundErr As Integer = -28
    Public Const teNoConnectionStringErr As Integer = -29
    Public Const teConnectionErr As Integer = -30
    Public Const teDatabaseNotFoundErr As Integer = -31
    Public Const teInvalidTableName As Integer = -32
    Public Const teRowNotFound As Integer = -33
    Public Const teNoForm As Integer = -34
    Public Const teCouldNotClearErr As Integer = -35
    Public Const teNoIdErr As Integer = -36
    Public Const teConnectionOpen As Integer = -37
    Public Const teCanNotDelete As Integer = -38
    Public Const teTableHasErrors As Integer = -39
    Public Const teColumnExists As Integer = -40
    Public Const teMakeTableErr As Integer = -41
    Public Const teNoRowErr As Integer = -42
    Public Const teNoDataView As Integer = -43
    Public Const teNoLocalDatabase As Integer = -44
    Public Const teNoDataRow As Integer = -45
    Public Const teCannotReload As Integer = -46
    Public Const teKeyValueAlreadyExists As Integer = -47
    Public Const teDataSetErr As Integer = -48
    Public Const teNoMasterRow As Integer = -49
    Public Const teMakeColumnErr As Integer = -50
    Public Const teNoViewErr As Integer = -51
    Public Const teActualColumnsErr As Integer = -52
    Public Const teRowCountErr As Integer = -53
    Public Const teNeedsToBeUpdated As Integer = -54
    Public Const teNoPrimaryKey As Integer = -55
    Public Const teMergeRowErr As Integer = -56
    Public Const teRowsDontMatchErr As Integer = -57
    Public Const teOtherErr As Integer = -59

#End Region

#Region " Create Database Errors "

    Public Const deCouldNotCreate As Integer = -60
    Public Const deDatabaseNotFound As Integer = -61
    Public Const deNoColumnName As Integer = -62
    Public Const deNoTextLength As Integer = -63
    Public Const deInvalidColumnType As Integer = -64
    Public Const deCreateTable As Integer = -65
    Public Const deCreatePrimaryKey As Integer = -66
    Public Const deCreateIndexes As Integer = -67
    Public Const deVerifyTableStruct As Integer = -68
    Public Const deNoDataAdapt As Integer = -69
    Public Const deVerifyTableData As Integer = -70
    Public Const deCreateQuery As Integer = -71
    Public Const deSortQuery As Integer = -72
    Public Const deVerifyQuery As Integer = -73
    Public Const deVerifyTableCount As Integer = -74
    Public Const deNoQueryName As Integer = -75
    Public Const deNoQueryText As Integer = -76
    Public Const deNoDaoDbEngine As Integer = -77
    Public Const deNoDaoWorkspace As Integer = -78
    Public Const deNoCheckbox As Integer = -79

#End Region

#Region " Non Link Errors -80 "

    Public Const nlNoParentRelsErr As Integer = -80
    Public Const nlMultiParentRelsErr As Integer = -81
    Public Const nlNoLinkedColumnsErr As Integer = -82
    Public Const nlMultiLinkedColumnsErr As Integer = -83

#End Region

#Region " Form Errors -90 "

    Public Const feNoFormErr As Integer = -90
    Public Const feChildHasError As Integer = -91

#End Region

#Region " KeyWord Error -99 "

    Public Const deKeyWord As Integer = -99

#End Region

#Region " Id Errors "

    Public Const idMinimumId As Integer = 1
    Public Const idNotSet As Integer = 0
    Public Const idNoNewId As Integer = -1
    Public Const idRetryForNew As Integer = -2
    Public Const idCouldNotInsert As Integer = -3
    Public Const idInvalidId As Integer = -4
    Public Const idAutoIncErr As Integer = -5
    Public Const idUserCancelErr As Integer = -10   ' save value as table error user cancel (teUserCancelErr)
    Public Const idAllIds As Integer = -99
    Public Const idNoEwJobNo As String = ""

    Public Const idMaxIdTries As Integer = 10

#End Region

#Region " Invoice Calculation Errors -101 "

    Public Const icUserCancel As Integer = -101
    Public Const icBillCalcStatusErr As Integer = -102
    Public Const icNoRowsUpdated As Integer = -103
    Public Const icMultiRowsUpdated As Integer = -104
    Public Const icGetFeeErr As Integer = -105
    Public Const icGetFixedFeeErr As Integer = -106
    Public Const icGetHourlyErr As Integer = -107
    Public Const icGetPriorFeeErr As Integer = -108
    Public Const icGetReimErr As Integer = -109
    Public Const icGetDelivErr As Integer = -110
    Public Const icGetTravelErr As Integer = -111
    Public Const icGetMarkupErr As Integer = -112
    Public Const icGetReimTotalErr As Integer = -113
    Public Const icNotLoggedInErr As Integer = -114
    Public Const icCannotSave As Integer = -115
    Public Const icNoFeeInvs As Integer = -116
    Public Const icGetMessageErr As Integer = -117
    Public Const icSaveErr As Integer = -118
    Public Const icUserAbourt As Integer = -119
    Public Const icEditingJob As Integer = -120
    Public Const icPromoJob As Integer = -121
    Public Const icOtherErr As Integer = -125

    Public Const icPaidSumErr As Integer = -150
    Public Const icAdjSumErr As Integer = -151
    Public Const icBalanceCalcErr As Integer = -152

#End Region

#Region " Invalid Data -200 "

    Public Const NoIdErr As Integer = -200
    Public Const NoValueErr As Integer = -201
    Public Const DuplicateErr As Integer = -202
    Public Const IdMismatchErr As Integer = -203
    Public Const NoMatchingChildErr As Integer = -204
    Public Const ChildDataErr As Integer = -205
    Public Const InvalidGrid As Integer = -206

#End Region

#Region " Value Errors -250 "

    Public Const veInteger As Integer = -250
    Public Const veString As Integer = -251
    Public Const veDecimal As Integer = -252
    Public Const veDate As Integer = -253
    Public Const veDouble As Integer = -254
    Public Const veBoolean As Integer = -255
    Public Const veNoData As Integer = -256

#End Region

#Region " Registry Errors -300 "

    Public Const reNoKey As Integer = -300
    Public Const reArgument As Integer = -301
    Public Const reArgumentNull As Integer = -302
    Public Const reIo As Integer = -303
    Public Const reNull As Integer = -304
    Public Const reObjectDisposed As Integer = -305
    Public Const reSecurity As Integer = -306
    Public Const reUnAuth As Integer = -307
    Public Const reOther As Integer = -399

#End Region

#Region " Io Errors -400 "

    Public Const ioErr As Integer = -400
    Public Const ioArgumentErr As Integer = -401
    Public Const ioArgumentNullErr As Integer = -402
    Public Const ioPathTooLongErr As Integer = -403
    Public Const ioNotSupportedErr As Integer = -404
    Public Const ioUnAuthErr As Integer = -405
    Public Const ioSecurityErr As Integer = -406
    Public Const ioFileNotFoundErr As Integer = -407
    Public Const ioOpCanceledErr As Integer = -408
    Public Const ioFolderNotFoundErr As Integer = -409
    Public Const ioInvalidParamErr As Integer = -498
    Public Const ioOtherErr As Integer = -499

#End Region

#Region " SQL Errors -500 "

    Public Const sqlErr As Integer = -500
    Public Const sqlNoConnection As Integer = -501
    Public Const sqlInvalidConnStr As Integer = -502
    Public Const sqlNoConvertToStr As Integer = -503
    Public Const sqlDeleteErr As Integer = -504
    Public Const sqlDropErr As Integer = -505
    Public Const sqlCreateErr As Integer = -506
    Public Const sqlAlterErr As Integer = -507
    Public Const sqlInvalidAlter As Integer = -508
    Public Const sqlNoDatabase As Integer = -509
    Public Const sqlNoTableName As Integer = -510
    Public Const sqlNoKeyName As Integer = -511
    Public Const sqlNoIndexName As Integer = -512
    Public Const sqlNoColumn As Integer = -513
    Public Const sqlInsertErr As Integer = -514
    Public Const sqlUpdateErr As Integer = -515
    Public Const sqlMakeTableErr As Integer = -516
    Public Const sqlRenameTableErr As Integer = -517
    Public Const sqlCreateIndexErr As Integer = -518
    Public Const sqlRenameColumnErr As Integer = -519
    Public Const sqlNoQueryName As Integer = -520
    Public Const sqlNoProcName As Integer = -521
    Public Const sqlNoViewName As Integer = -522
    Public Const sqlTableNotFound As Integer = -523

#End Region

#Region " Projects Data Errors -1000 "

    Public Const pjeChildHasErr As Integer = -1000

#End Region

#Region " Companies Data Errors -1100 "

    Public Const cmeChildHasErr As Integer = -1100

#End Region

#Region " Contacts Data Errors -1200 "

    Public Const cneChildHasErr As Integer = -1200

#End Region

#Region " Travel Data Errors -1300 "

    Public Const tveChildHasErr As Integer = -1300

#End Region

#Region " Jobs Data Errors -1400 "

    Public Const jeChildHasErr As Integer = -1400
    Public Const jeNoJobNumber As Integer = -1401
    Public Const jeNoJobName As Integer = -1402
    Public Const jeNoJobCity As Integer = -1403
    Public Const jeNoOpenedBy As Integer = -1404
    Public Const jeNoOpenedOn As Integer = -1405
    Public Const jeNoClientName As Integer = -1406
    Public Const jeNoAttn As Integer = -1407
    Public Const jeNoFixedFeeData As Integer = -1408
    Public Const jeInvalidFixedFeeErr As Integer = -1409
    Public Const jeNoFixedFeeTotalErr As Integer = -1410
    Public Const jeFixedFeeBilledTooHighErr As Integer = -1411
    Public Const jeFixedFeeBilledNoMatchErr As Integer = -1412
    Public Const jeNoRate As Integer = -1413
    Public Const jeZeroOrLessBillRate As Integer = -1414
    Public Const jeNegativeFeeMax As Integer = -1415
    Public Const jeFeeMaxLessThanRate As Integer = -1416
    Public Const jeFeeMaxLessThanBilled As Integer = -1417
    Public Const jeNoClientAddress As Integer = -1418
    Public Const jeJobNumberMismatch As Integer = -1419
    Public Const jeNoDepartment As Integer = -1420
    Public Const jeNoLinkProj As Integer = -1421
    Public Const jeNoBillTo As Integer = -1422
    Public Const jeOtherFixedFeeErr As Integer = -1423
    Public Const jeCouldNotAddProject As Integer = -1430
    Public Const jeNoProjectId As Integer = -1431
    Public Const jeEditingProject As Integer = -1432
    Public Const jeJobNoNotFound As Integer = -1433
    Public Const jeJobNoAlreadyUsed As Integer = -1434

    Public Const jeJobNumberAlreadyUsed As Integer = -1450
    Public Const jeInvalidJobNumber As Integer = -1451
    Public Const jeJobNumberNotChanged As Integer = -1452

#End Region

#Region " Invoice Data Errors -1500 "

    Public Const invChildHasErr As Integer = -1500
    Public Const invFfPhaseNotFound As Integer = -1501
    Public Const invVersionErr As Integer = -1502

#End Region

#Region " Payment Data Errors -1600 "

    Public Const peChildHasErr As Integer = -1600
    Public Const peNotBalanced As Integer = -1601
    Public Const peTotalFeeCalcErr As Integer = -1602
    Public Const peInsertPaymentErr As Integer = -1603
    Public Const peInvoicePaymentErr As Integer = -1604
    Public Const peMiscPaymentsErr As Integer = -1605
    Public Const peCreditPaymentErr As Integer = -1606
    Public Const peSaveErr As Integer = -1607
    Public Const peInvNotFound As Integer = -1608
    Public Const peEmpNoFound As Integer = -1609
    Public Const peLoadPaymentErr As Integer = -1610
    Public Const peDeleteErr As Integer = -1611
    Public Const peOverpaymentErr As Integer = -1612
    Public Const peEditPaymentErr As Integer = -1613
    Public Const peDataErr As Integer = -1614
    Public Const peCreditIdErr As Integer = -1615
    Public Const peNewPaymentErr As Integer = -1616

    Public Const peUserCancelErr As Integer = -1698
    Public Const peOtherErr As Integer = -1699

#End Region

#Region " Timecard Data Errors -1700 "

    Public Const tcChildHasErr As Integer = -1700
    Public Const tcNoJobs As Integer = -1701
    Public Const tcNoActiveJobs As Integer = -1702
    Public Const tcNoEwEmployees As Integer = -1703
    Public Const tcNoActiveEwEmployees As Integer = -1704

#End Region

#Region " Work Done Data Errors -1800 "

    Public Const wdChildHasErr As Integer = -1800
    Public Const wdNoJobs As Integer = -1801
    Public Const wdNoActiveJobs As Integer = -1802
    Public Const wdNoEwEmployees As Integer = -1803
    Public Const wdNoActiveEwEmployees As Integer = -1804
    Public Const wdWrongDataType As Integer = -1805

#End Region

#Region " Delivery Data Errors -1900 "

    Public Const dvChildHasErr As Integer = -1900

#End Region

#Region " Cloud Computing Errors -2000 "

    Public Const cceAddColErr As Integer = -2000
    Public Const cceInitPopErr As Integer = -2001
    Public Const cceCouldNotSet As Integer = -2002
    Public Const cceInvalidEditType As Integer = -2003
    Public Const cceChildDeletesErr As Integer = -2004
    Public Const cceMultiFieldKeyErr As Integer = -2005
    Public Const cceNoKeyFieldErr As Integer = -2006
    Public Const cceCloudEditKeyErr As Integer = -2007
    Public Const cceChildRelinkErr As Integer = -2008
    Public Const cceNewNetworkRowErr As Integer = -2009
    Public Const cceMultiFieldLinkErr As Integer = -2010
    Public Const cceNoLinkFieldErr As Integer = -2011
    Public Const cceNoIdColumnErr As Integer = -2012
    Public Const cceColumnIsNull As Integer = -2013
    Public Const cceDataLocIsNull As Integer = -2014
    Public Const cceUpdatedWhenIsNull As Integer = -2015
    Public Const cceUpdatedByIsNull As Integer = -2016
    Public Const cceOlsRelinkError As Integer = -2017

    Public Const cceSetCloudColsErr As Integer = -2020

    Public Const cceNoTableName As Integer = -2050
    Public Const cceNoKeyValue As Integer = -2051
    Public Const cceNoTable As Integer = -2052
    Public Const cceNoBindingManagerBase As Integer = -2053
    Public Const cceSqlSelectErr As Integer = -2054
    Public Const cceNoLocalRow As Integer = -2055
    Public Const cceBackupErr As Integer = -2056

#End Region

#Region " Outlook Errors "

    Public Const oeOpenWindows As Integer = -2100
    Public Const oeCouldNotConvert As Integer = -2101
    Public Const oeCouldNotAdd As Integer = -2102
    Public Const oeNoContacts As Integer = -2103
    Public Const oeCouldNotSetSyncRow As Integer = -2104

#End Region

#Region " Tapi Errors - 3000 "

    Public Const tapiNoModem As Integer = -3001
    Public Const tapiNoPhoneNumber As Integer = -3002
    Public Const tapiPhoneNumberToShort As Integer = -3003
    Public Const tapiInvalidPhoneNumberChar As Integer = -3004
    Public Const tapiNoTapi As Integer = -3005
    Public Const tapiNoAddress As Integer = -3006
    Public Const tapiOtherErr As Integer = -3010


#End Region

#Region " OleDb Errors "

    Public Const oleAlreadyInUse As Integer = -2146825243
    Public Const oleCannotOpen As Integer = -2147467259
    Public Const oleCannotModify As Integer = -2147217911
    Public Const oleNoColumn As Integer = -2147217904

    Public Const oleMaxTries As Integer = 10
    Public Const oleWaitMiliSeconds As Integer = 1000

#End Region

#End Region

    '****************
    ' ERROR TITLES
    '****************

#Region " Error Titles "

    Public Const esShowErrMessage As Boolean = True
    Public Const esNoErrorMessage As Boolean = False
    Public Const esAskToFixIfErr As Boolean = True
    Public Const esNoAskToFixIfErr As Boolean = False

    ' when a error messge box is show, theses are the titles for the error windows

    Public Const etGetDataSetErr As String = "Get DataSet Error"
    Public Const etSetupErr As String = "Setup Error"

    Public Const etActivateWindowErr As String = "Activating Window Error"
    Public Const etCancelEditErr As String = "Cancel Edit Error"
    Public Const etCancelErr As String = "Cancel Error"
    Public Const etCancelInsertErr As String = "Cancel Insert Error"
    Public Const etCannotAdvFilter As String = "Cannot use Advanced Filter"
    Public Const etCannotDelete As String = "Cannot Delete"
    Public Const etCannotEdit As String = "Cannot Edit"
    Public Const etCannotPrint As String = "Cannot Print"
    Public Const etCannotReload As String = "Cannot Reload"
    Public Const etClearTableErr As String = "Clear Table Error"
    Public Const etConcurrencyErr As String = "Concurrency Error"
    Public Const etConcurrencyRes As String = "Concurrency Resolution"
    Public Const etConfirmationErr As String = "Confirmation Error"
    Public Const etConfirmFfLogErr As String = "Confirm Fixed Fee Log Error"
    Public Const etCreateWindowErr As String = "Creating Window Error"
    Public Const etDataBindingErr As String = "Data Binding Error"
    Public Const etDataEntryErr As String = "Data Entry Error"
    Public Const etDataErr As String = "Data Error"
    Public Const etDataNotFound As String = "Data Not Found"
    Public Const etDataValidationErr As String = "Data Validation Error"
    Public Const etDialogErr As String = "Dialog Window Error"
    Public Const etEMailErr As String = "EMail Error"
    Public Const etEndEditErr As String = "End Edit Error"
    Public Const etErrorFindingErr As String = "Error Finding Error"
    Public Const etFillInErr As String = "Fill In Error"
    Public Const etFixingConstraintErr As String = "Fixing Constraint Error"    
    Public Const etImportRowErr As String = "Import Data Row Error"
    Public Const etInternalErr As String = "Internal Error"
    Public Const etInsertErr As String = "Insert Error"
    Public Const etKeyColumnErr As String = "Key Column Error"
    Public Const etLoadDataErr As String = "Load Data Error"
    Public Const etMySqlErr As String = "MySql Error"
    Public Const etNoDataLoaded As String = "No Data Loaded"
    Public Const etNoDataSetErr As String = "No DataSet Error"
    Public Const etNotInList As String = "Not in List"
    Public Const etOtherErr As String = "Other Error"
    Public Const etSaveErr As String = "Save Error"
    Public Const etSwitchViewsErr As String = "Error Switching Grid Views"
    Public Const etSwitchWindowsErr As String = "Switching Windows Error"
    Public Const etUndoErr As String = "Undo Error"
    Public Const etUnlinkErr As String = "Unlink Error"
    Public Const etUpdateErr As String = "Update Error"
    Public Const etUserCancel As String = "User Canceled"
    Public Const etValidateRowErr As String = "Validate Row Error"
    Public Const etWWWErr As String = "WWW Error"

    Public Const etEditJobErr As String = "Edit Job Error"
    Public Const etNewJobErr As String = "New Job Error"
    Public Const etEwJobNoError As String = "EW Job # Error"
    Public Const etViewJobErr As String = "View Job Error"

    Public Const etEditProjectErr As String = "Edit Project Error"
    Public Const etUpdateProjectHistory As String = "Update Project History Error"
    Public Const etKeyContactErr As String = "Save Key Contact Error"
    Public Const etViewProjectErr As String = "View Project Error"

    Public Const etNewIdError As String = "Get New ID Number Error"
    Public Const etDeleteRow As String = "Delete Row Error"
    Public Const etValueNotSetError As String = "Value not set Error"

    Public Const etCompAcctErr As String = "Company Account Error"
    Public Const etViewCompanyErr As String = "View Company Error"

    Public Const etNewContactErr As String = "New Contact Error"
    Public Const etEditContactErr As String = "Edit Contact Error"
    Public Const etViewContactErr As String = "View Contact Error"

    Public Const etCredit As String = "Credit Error"
    Public Const etNoCredits As String = "No Credits"

    Public Const etPrintGrid As String = "Print Grid Error"
    Public Const etPrint As String = "Print Error"
    Public Const etNoDataToPrint As String = "No Data to Print"

    Public Const etCannotAddInvoice As String = "Cannot Add Manual Invoice"
    Public Const etCannotReviseInvoice As String = "Cannot Revise Invoice Error"
    Public Const etCannotSavePrintInvoice As String = "Cannot Save and Print Invoices"
    Public Const etCalcInvoices As String = "Calculate Invoices Error"
    Public Const etEmpHourErr As String = "Employee Hourly Error"
    Public Const etFixedFeeErr As String = "Fixed Fee Error"
    Public Const etGetManualInvoice As String = "Get Manual Invoice Error"
    Public Const etInsertInvoices As String = "Insert Invoices Error"
    Public Const etInvoiceVersion As String = "Invoice Version Error"
    Public Const etInviceMessageErr As String = "Invoice Message Error"
    Public Const etLoadInvoice As String = "Load Invoice Error"
    Public Const etPreBillMessageErr As String = "Pre Billing Error"
    Public Const etPrintMonthlyInvoices As String = "Print Monthly Invoices Error"
    Public Const etPrintInvoice As String = "Print Invoice Error"
    Public Const etRevisonInvoice As String = "Invoice Revision Error"
    Public Const etSaveFfHistory As String = "Save Fixed Fee History Error"
    Public Const etSaveManualInvoice As String = "Save Manual Invoice Error"
    Public Const etSaveNewInvoices As String = "Save New Invoices Error"
    Public Const etUpdateInvoiceBalance As String = "Updating Invoice Balance Error"
    Public Const etUpdateInvoiceErr As String = "Update Invoice Error"

    Public Const etDeletePayment As String = "Delete Payment Error"
    Public Const etDisbursePayments As String = "Payment Disburse Error"
    Public Const etEditPayment As String = "Edit Payment Error"
    Public Const etLoadPayments As String = "Load Payments Error"
    Public Const etNoInvsForPayment As String = "No Outstaning Invoices"
    Public Const etPaymentData As String = "Payment Error"
    Public Const etPaymentsDialog As String = "Payments Dialog Error"
    Public Const etSavePayment As String = "Save Payment"
    Public Const etUnuseCreditPayment As String = "Unuse Credit Payment Error"

    Public Const etLoadDelivery As String = "Load Delivery Error"
    Public Const etSetDeliveryValue As String = "Set Delivery Values Error."

    Public Const etLoadTimecard As String = "Load Timecard Error"
    Public Const etSetTimecardValue As String = "Set Timecard Values Error."

    Public Const etLoadTravel As String = "Load Travel Expense Error"
    Public Const etSetTravelValue As String = "Set Travel Expense Values Error."
    Public Const etTravelErr As String = "Travel Expense Error"
    Public Const etCannotDeleteDay As String = "Cannot Delete Day"

    Public Const etLoadWorkDone As String = "Load Work Done Error"
    Public Const etSetWorkDoneValue As String = "Set Work Done Values Error"

    Public Const etDataPath As String = "Error Loading Database"
    Public Const etOpenConnection As String = "Connection is Open Error"
    Public Const etLoadTable As String = "Load Table Error"
    Public Const etGetRowErr As String = "Get Row Error"
    Public Const etNoTable As String = "Table Not Found Error"
    Public Const etGetConnection As String = "Get Database Connection Error"
    Public Const etFolderNotFound As String = "Folder not Found"
    Public Const etDatabaseNotFound As String = "Database not Found"
    Public Const etNoColumn As String = "Column not in Table"
    Public Const etNoDataView As String = "No DataView Error"

    Public Const etCannotMerge As String = "Cannot Merge"
    Public Const etMerge As String = "Merge Error"

    Public Const etInsertJobErr As String = "Insert Job Error"
    Public Const etJobValidErr As String = "Job Validation Error"
    Public Const etFixedFeeLogErr As String = "Fixed Fee Log Error"
    Public Const etUndisbErr As String = "Undisburse Error"
    Public Const etCannotCreateNewJob As String = "Cannot Create New Job"

    Public Const etRedisbureErr As String = "Redisburse Error"
    Public Const etReviseErr As String = "Revise Error"
    Public Const etReviseFfErr As String = "Revised Fixed Fee Error"

    Public Const etInvalidPhoneNumber As String = "Phone Number Error"
    Public Const etTapi As String = "TAPI/Modem Communications Error"

    Public Const etSqlAddIndexErr As String = "Add Index Error"
    Public Const etSqlAlterErr As String = "Alter Table Error"
    Public Const etSqlCanNotCloseErr As String = "Cannot Close Connection Error"
    Public Const etSqlColumnExistsErr As String = "SQL Column Exists Error"
    Public Const etSqlCreateErr As String = "Create Table Error"
    Public Const etSqlCreateIndexErr As String = "Create Index Error"
    Public Const etSqlCreateQueryErr As String = "Create Query Error"
    Public Const etSqlDropErr As String = "Drop Table Error"
    Public Const etSqlDropIndexErr As String = "Drop Index Error"
    Public Const etSqlDropProcErr As String = "Drop Procedure Error"
    Public Const etSqlDropViewErr As String = "Drop View Error"
    Public Const etSqlError As String = "SQL Error"
    Public Const etSqlInnerJoinErr As String = "Inner Join Error"
    Public Const etSqlLoadErr As String = "Load via SQL Error"
    Public Const etSqlRenameColumnErr As String = "Rename Column Error"
    Public Const etSqlRenameTableErr As String = "Rename Table Error"

    Public Const etNoBillingAddr As String = "No Billing Address"

    Public Const etFindErrorErr As String = Pcm.etErrorFindingErr

    Public Const etRegistryErr As String = "Registry Error"
    Public Const etIoErr As String = "I/O Error"
    Public Const etConfirmOverwrite As String = "Confirm Overwrite"
    Public Const etConfirm As String = "Confirm"

    Public Const etEmployeeDataErr As String = "EW Employee Data Error"
    Public Const etEmployeeEditErr As String = "Edit Employee Error"

    Public Const etCloudComputingErr As String = "Cloud Computing Error"
    Public Const etSyncErr As String = "Synchronize Error"
    Public Const etCannotWithLocalData As String = "Function not avalible when using local data"
    Public Const etBackupLocalDataErr As String = "Backup Local Database Error"

    Public Const etSyncWithOutlook As String = "Error in Sync with Outlook"
    Public Const etWindowOpenErr As String = "Window is Open Error"

#End Region

#Region " LookupEdit Values "

    Public Const lueCanAdd As Boolean = True
    Public Const lueNoAdd As Boolean = False

#End Region

#Region " Controls Spacing Values "

    Public Const CtrlHeightGap As Integer = 2

#End Region

#Region " Goto Grid Values "

    Public Const gtgShowGrid As Boolean = True
    Public Const gtgHideGrid As Boolean = False
    Public Const gtgEdit As Boolean = True
    Public Const gtgView As Boolean = False

#End Region

#Region " Search For Hints "

    Public Const sfConciseToolTip As String = "Concise Grid"
    Public Const sfDetailToolTip As String = "Detail Grid"

    Public Const sfConciseIndex As Integer = 0
    Public Const sfDetailIndex As Integer = 1

#End Region

#Region " Job # Values "

    Public Const jbMinJobNo As String = "1000"
    Public Const jbPromoDigit As String = "9"
    Public Const jbPromoCode As String = ".90"
    Public Const jbMaxPromoCode As String = ".99"
    Public Const jbSpecialJobStart As String = "8000"
    Public Const jbInternalJobStart As String = "9990"
    Public Const jbProfDevJobStart As String = "9998"
    Public Const jbMaxJobNo As String = "9999.99"

    Public Const jbBaseLength As Integer = 4                            ' ####
    Public Const jbFirstPromoDigitIndex As Integer = jbBaseLength + 1
    Public Const jbExtLength As Integer = 3                             ' .""
    Public Const jbIntStartLength As Integer = 3                        ' 999
    Public Const jbFullLength As Integer = jbBaseLength + jbExtLength   ' ####.##
    Public Const jbPromoDigitPos As Integer = jbBaseLength + 1          ' ####.9#

    Public Const jbExtIncrement As Double = 0.01
    Public Const jbPromoDbl As Double = 0.9
    Public Const jbMaxPromoDbl As Double = 0.99

    Public Const jbNewWhole As Integer = 0
    Public Const jbNewExt As Integer = 1

    Public Const jbNormal As Integer = 0
    Public Const jbSpecial As Integer = 1
    Public Const jbInternal As Integer = 2

    Public Const ptNotSet As Integer = -1
    Public Const ptNewProject As Integer = 0
    Public Const ptSubProject As Integer = 1
    Public Const ptFromCurrent As Integer = 2

#End Region

#Region " Bill Types/Address Types and Invoice Calculation Status "

    Public Const btEmployee As String = "Employee"
    Public Const btFixedFee As String = "Fixed Fee"
    Public Const btHourly As String = "Hourly"
    Public Const btContact As String = "Contact"
    Public Const btDepartment As String = "Department"

    Public Const MonthFfAmntColName As String = "colFfAmount"

    Public Const baBillAddrNotFound As String = "Billing Address not found"

    Public Const atRegular As String = "Regular"
    Public Const atBilling As String = "Billing"

#End Region

#Region " Fixed Fee Percentages/Fixed Fee Log Types "

    Public Const PercentFormat As String = "##0.##\%"

    Public Const ffMinPer As Double = 0.0
    Public Const ffMaxPer As Double = 100.0
    Public Const ffNearToFull As Double = 0.005

    Public Const fltNew As Integer = 1
    Public Const fltInvoice As Integer = 2
    Public Const fltFixedFeeRev As Integer = 3
    Public Const fltInvoiceRev As Integer = 4
    Public Const fltImported As Integer = 5
    Public Const fltReset As Integer = 6
    Public Const fltRedisburse As Integer = 7
    Public Const fltSystemAdjust As Integer = 8
    Public Const fltHourlyConvert As Integer = 9
    Public Const fltRevAllFf As Integer = 10
    Public Const fltRevAllReset As Integer = 11

    Public Const flfNewFormat As String = "New Fixed Fee Phase. Amount: {0:C}"
    Public Const flfInvFormat As String = "New Invoice. Billed {0:0.00}%. {1:C} * {0:0.00}% = {2:C}"
    Public Const flfFFRevFormat As String = "Revised Fixed Fee Phase Amount, from {0:C} to {1:C}. {2:C} / {1:C} = {3:0.00}%.  {3:0.00}% - {4:0.00}% = {5:0.00}%"
    Public Const flfInvRevFormat As String = "Revised Invoice, Fixed Fee Phase to {0:0.00}% Billed, {1:C}.  {1:C} / {2:C} = {0:0.00}%.  {0:0.00}% - {3:0.00}% = {4:0.00}%.  {4:0.00}% * {2:C} = {5:C}"
    Public Const flfImportFormat As String = "Imported Fixed Fee Phase. Amount: {0:C}; Billed: {1:C}; Billed: {2:0.00}%"
    Public Const flfResetFormat As String = "Reset Fixed Fee Phase for Redisbursal.  Amount: {0:C}; Billed: $0.00; Billed: 0.00%"
    Public Const flfRedisburseFormat As String = "Redisburse Fixed Fee Phase.  {0:0.00}% Billed.  {1:C} * {0:0.00}% = {2:C}"
    Public Const flfSystemAdjustFormat As String = "System Adjustment to: Amount: {0:C}; Billed: {1:C}; Billed: {2:0.00}%.  {1:C} - {3:C} = {4:C}.  {2:0.00}% - {5:0.00}% = {6:0.00}%"
    Public Const flfHourlyConvertFormat As String = "Converted from Hourly Job.  Amount: {0:C}; Billed: {1:C}; Billed: {2:0.00}%.  {1:C} / {0:C} = {2:0.00}%"
    Public Const flfRevAllFfFormat As String = "Revised All Fixed Fee Invoices. Billed {0:0.00}%. {1:C} * {0:0.00}% = {2:C}"
    Public Const flfRevAllResetFormat As String = "Reset Fixed Fee Phase for Revise All.  Amount: {0:C}; Billed: {1:C}; Billed: {2:0.00}%"

#End Region

#Region " Invoice Types, Invoice # Range, Print Invoice Version "

    Public Const ieInvNoIndex As Integer = 0
    Public Const ieEwJobIdIndex As Integer = 1
    Public Const ieWidth As Integer = 1             ' 2 array cells wide

    Public Const itAll As String = "All"
    Public Const itCredit As String = "Credit"
    Public Const itFee As String = "Fee"
    Public Const itInvoice As String = "Invoice"
    Public Const itMisc As String = "Miscellaneous"
    Public Const itMonth As String = "Month"
    Public Const itReim As String = "Reimbursable"
    Public Const itUsed As String = "Used Credit"
    Public Const itUserSpecified As String = "User Specified"

    Public Const InvNoMin As Integer = 10000
    Public Const InvNoMax As Integer = 2147483647

    Public Const ivNotSet As Integer = -1
    Public Const ivOriginal As Integer = 0

    Public Const pivCurrent As Boolean = True
    Public Const pivOriginal As Boolean = False

#End Region

#Region " Work Done Data Types "

    Public Const wdShowingAllEmps As Integer = 0
    Public Const wdShowingOneEmp As Integer = 1

#End Region

#Region " Make Child Form Values "

    Public Const cfMade As Integer = 0
    Public Const cfAlreadyExists As Integer = 1
    Public Const cfMakeErr As Integer = -1

#End Region

#Region " Update Type Values "

    Public Const utPostEdit As Integer = 0
    Public Const utPostDelete As Integer = 1

#End Region

#Region " Child Form Names/Titles "

    Public Const cfnCompanies As String = "CompaniesForm"
    Public Const cfnContacts As String = "ContactsForm"
    Public Const cfnDelivForm As String = "DelivForm"
    Public Const cfnJobs As String = "JobsForm"
    Public Const cfnInvoicesForm As String = "InvoicesForm"
    Public Const cfnProjects As String = "ProjectsForm"
    Public Const cfnPayments As String = "PaymentsForm"
    Public Const cfnTimecard As String = "TimecardForm"
    Public Const cfnTravel As String = "TravelForm"
    Public Const cfnWorkDone As String = "WorkDoneForm"

    Public Const cftCompanies As String = "Companies"
    Public Const cftContacts As String = "Contacts"
    Public Const cftDelivForm As String = "Delivieries"
    Public Const cftJobs As String = "Jobs"
    Public Const cftInvoicesForm As String = "Invoices"
    Public Const cftProjects As String = "Projects"
    Public Const cftPayments As String = "Payments"
    Public Const cftTimecard As String = "Timecard"
    Public Const cftTravel As String = "Travelt"
    Public Const cftWorkDone As String = "WorkDone"

#End Region

    '****************
    ' TABLE NAMES
    '****************

#Region " Table Names "

    Public Const AdjFfHistTableName As String = "AdjFfHist"
    Public Const AdjFfInvTableName As String = "AdjFfInv"
    Public Const AffilTableName As String = "Affiliations"    
    Public Const AutoIncTableName As String = "AutoInc"
    Public Const BillingAddrTableName As String = "BillingAddr"
    Public Const BldgTypesTableName As String = "BuildingTypes"
    Public Const CalcInvsTableName As String = "CalcInvs"
    Public Const CitiesTableName As String = "Cities"
    Public Const CloudEditsTableName As String = "CloudEdits"
    Public Const CloudDelsTableName As String = "CloudDels"
    Public Const CodesTableName As String = "Codes"
    Public Const CompAcctTableName As String = "CompAcct"
    Public Const CompAcctPayTableName As String = "CompAcctPay"    
    Public Const CompaniesTableName As String = "Companies"
    Public Const CompanyRolesTableName As String = "CompanyRoles"
    Public Const CompContsTableName As String = "CompConts"
    Public Const CompJobsTableName As String = "CompJobs"
    Public Const CompPayInvsTableName As String = "CompPayingInvs"
    Public Const CompProjsTableName As String = "CompProjs"
    Public Const ContactsTableName As String = "Contacts"
    Public Const ContactAffilTableName As String = "ContactAffiliations"
    Public Const ContJobsTableName As String = "ContJobs"
    Public Const ContProjsTableName As String = "ContProjs"
    Public Const CountriesTableName As String = "Countries"
    Public Const CountTableName As String = "CountTable"
    Public Const CreditsTableName As String = "Credits"
    Public Const CreditUsedTableName As String = "CreditUsed"
    Public Const DelivTableName As String = "Deliveries"
    Public Const DelivToBillTableName As String = "DelivToBill"
    Public Const EditInvPayTableName As String = "EditInvPay"
    Public Const EmpHourHistTableName As String = "EmpHourHistory"
    Public Const EmpHourToBillTableName As String = "EmpHourToBill"
    Public Const ElevCntrsTableName As String = "ElevContractors"
    Public Const EwEmpsTableName As String = "EwEmployees"
    Public Const ExpDisbTableName As String = "ExpDisb"
    Public Const FeeToBillTableName As String = "FeeToBill"
    Public Const FfBilledTableName As String = "FfBilled"
    Public Const FfForJobTableName As String = "FfForJob"
    Public Const FfHistCheckTableName As String = "FfHistCheck"
    Public Const FfHistoryTableName As String = "FfHistory"
    Public Const FfIdPreInvTableName As String = "FfIdPreInv"
    Public Const FfLogCheckTableName As String = "FfLogCheck"
    Public Const FfLogResetMaxIdTableName As String = "FfLogResetMaxId"
    Public Const FfLogTableName As String = "FixedFeeLog"
    Public Const FfPiTableName As String = "FfPreInv"    
    Public Const FixedFeeTableName As String = "FixedFee"
    Public Const GenericTableName As String = "Generic"
    Public Const HourlyHistTableName As String = "HourlyHistory"
    Public Const HourlyToBillTableName As String = "HourlyToBill"
    Public Const ImpFeeToBillTableName As String = "ImpFeeToBill"
    Public Const ImpReimToBillTableName As String = "ImpReimToBill"
    Public Const IndvRolesTableName As String = "IndvRoles"
    Public Const InsEhInvTableName As String = "InsEhInvoice"
    Public Const InsFfInvTableName As String = "InsFfInvoice"
    Public Const InsInvFfPhaseTableName As String = "InsInvFfPhases"
    Public Const InvAddrTableName As String = "InvAddr"
    Public Const InvAdjSumTableName As String = "InvAdjSum"
    Public Const InvAmntHistTableName As String = "InvAmntHistory"
    Public Const InvAllTableName As String = "InvAll"
    Public Const InvCalcTableName As String = "InvCalcStatus"
    Public Const InvCanPayTableName As String = "InvCanPay"
    Public Const InvCurValuesTableName As String = "InvCurValues"
    Public Const InvEmpHistTableName As String = "InvEmpHourHistory"
    Public Const InvImportTableName As String = "InvImport"
    Public Const InvFfHistTableName As String = "InvFfHistory"
    Public Const InvFfRevsTableName As String = "InvFfRevs"
    Public Const InvJobCompAttnTableName As String = "InvJobCompAttn"
    Public Const InvHistBillAddrTableName As String = "InvHistBillAddr"    
    Public Const InvoicesTableName As String = "Invoices"
    Public Const InvoiceRevsTableName As String = "InvoiceRevs"
    Public Const InvPastDueTableName As String = "InvPastDue"
    Public Const InvPaymentsTableName As String = "InvPayments"
    Public Const InvPaySumTableName As String = "InvPaySum"
    Public Const InvPayTableName As String = "InvPay"
    Public Const InvSomeTableName As String = "InvSome"
    Public Const JobDelivTableName As String = "JobDeliv"
    Public Const JobInvTempTableName As String = "JobInvTemp"
    Public Const JobInvsTableName As String = "JobInvs"
    Public Const JobInvSumTableName As String = "JobInvSum"
    Public Const JobLastInvSumTableName As String = "JobLastInvSum"
    Public Const JobMpTableName As String = "JobMp"
    Public Const JobsTableName As String = "Jobs"
    Public Const JobsNoContTableName As String = "JobsNoCont"
    Public Const JobTcTableName As String = "JobTc"
    Public Const JobTravTableName As String = "JobTrav"
    Public Const JobWdTableName As String = "JobWd"
    Public Const MaxFfLogIdTableName As String = "MaxFfLogId"
    Public Const MergePlayersTableName As String = "MergePlayers"
    Public Const MiscPayTableName As String = "MiscPayments"
    Public Const MonthEmpHoursTableName As String = "MonthEmpHours"
    Public Const MonthFfHistTableName As String = "MonthFfHist"
    Public Const MonthFfTableName As String = "MonthFf"    
    Public Const MonthMsgTableName As String = "MonthMsg"
    Public Const MsgHistoryTableName As String = "MsgHistory"
    Public Const MsgToBillTableName As String = "MsgToBill"
    Public Const NamePrefixesTableName As String = "NamePrefixes"
    Public Const NewIfrTableName As String = "NewInvFfRevs"
    Public Const NonInvMaxFfLogIdTableName As String = "NonInvMaxFfLogId"
    Public Const NonLinkedAdded As String = "NL"
    Public Const NonLinkedTableName As String = "NonLinked"
    Public Const NotUsedIfrTableName As String = "NotUsedIfr"
    Public Const OlContsTableName As String = "OutlookContacts"
    Public Const OlSyncStatTableName As String = "OlSyncStatus"
    Public Const PastDueHistTableName As String = "PastDueHistory"
    Public Const PayFormCreditTableName As String = "PayFormCredit"
    Public Const PayFormMiscTableName As String = "PayFormMisc"
    Public Const PayFormTableName As String = "PayForm"
    Public Const PayFormUsedTableName As String = "PayFormUsed"
    Public Const PayingInvsTableName As String = "PayingInvs"
    Public Const PaymentsTableName As String = "Payments"
    Public Const PaymentsToAddTableName As String = "PaymentsToAdd"
    Public Const PbDelivTableName As String = "PbDeliveries"
    Public Const PbEmpHourTableName As String = "PbEmpHour"
    Public Const PbFixedFeeTableName As String = "PbFixedFee"
    Public Const PbHourlyEmpsTableName As String = "PbHourlyEmps"
    Public Const PbHourlyJobsTableName As String = "PbHourlyJobs"
    Public Const PbHourlyJobsInvSumTableName As String = "PbHourlyJobsInvSum"
    Public Const PbRptFixedFeeTableName As String = "PbRptFixedFee"
    Public Const PbServMsgTableName As String = "PbServMsg"
    Public Const PbTravelTableName As String = "PbTravel"
    Public Const PlayersTableName As String = "Players"
    Public Const PreBillJobsTableName As String = "PreBillJobs"
    Public Const PreBillTableName As String = "PreBill"
    Public Const PriorFeeBillTableName As String = "PriorFeeBilling"
    Public Const ProjBldgsTableName As String = "ProjBldgs"
    Public Const ProjCodesTableName As String = "ProjCodes"
    Public Const ProjDrawLogTableName As String = "ProjDrawLog"
    Public Const ProjectsTableName As String = "Projects"
    Public Const ProjElevsTableName As String = "ProjElevs"
    Public Const ProjEscsTableName As String = "ProjEscs"
    Public Const ProjHistPlayersTableName As String = "ProjHistPlayers"
    Public Const ProjHistRolesTableName As String = "ProjHistRoles"
    Public Const ProjJobsTableName As String = "ProjJobs"
    Public Const ProjMaintTableName As String = "ProjMaint"
    Public Const ProjMWalksTableName As String = "ProjMWalks"
    Public Const ProjRolesTableName As String = "ProjRoles"
    Public Const ProjScopeTableName As String = "ProjScope"
    Public Const ProjSpecFeatTableName As String = "ProjSpecFeat"
    Public Const ReimbSumTableName As String = "ReimbSum"
    Public Const ReimHistoryTableName As String = "ReimHistory"
    Public Const ReimToBillTableName As String = "ReimToBill"
    Public Const RevFixedFeeTableName As String = "RevFixedFee"
    Public Const RolesTableName As String = "Roles"
    Public Const RptInvCompsTableName As String = "RptInvComps"
    Public Const RptInvContsTableName As String = "RptInvConts"
    Public Const RptInvEmpHoursTableName As String = "RptInvEmpHours"
    Public Const RptInvFfTableName As String = "RptInvFf"
    Public Const RptInvHoursTableName As String = "RptInvHours"
    Public Const RptInvMsgTableName As String = "RptInvMsg"
    Public Const RptInvoicesTableName As String = "RptInvoices"
    Public Const RptInvReimTableName As String = "RptInvReim"
    Public Const RptOutInvTableName As String = "RptOutInv"
    Public Const RptTravCabsTableName As String = "RptTravelCabs"
    Public Const RptTravDailyTableName As String = "RptTravelDaily"
    Public Const RptTravEntTableName As String = "RptTravelEnt"
    Public Const RptTravTableName As String = "RptTravel"
    Public Const SpecFeatTableName As String = "SpecialFeatures"
    Public Const StatesTableName As String = "States"
    Public Const SuffixTableName As String = "Suffix"
    Public Const TcPreBillTableName As String = "TcPreBill"
    Public Const TcFormTableName As String = "TcForm"
    Public Const TempCompaniesTableName As String = "TempCompanies"
    Public Const TempCreditsTableName As String = "TempCredits"
    Public Const TempEmpHourHistTableName As String = "TempEmpHourHist"
    Public Const TempEmpsTableName As String = "TempEmps"
    Public Const TempFfHistTableName As String = "TempFfHist"
    Public Const TempFfTableName As String = "TempFixedFee"    
    Public Const TempHourlyHistTableName As String = "TempHourlyHist"
    Public Const TempInvAmntHistTableName As String = "TempInvAmntHist"
    Public Const TempInvTableName As String = "TempInv"
    Public Const TempInvPayTableName As String = "TempInvPay"
    Public Const TempJobsTableName As String = "TempJobs"
    Public Const TempJobWdTableName As String = "TempJobWd"
    Public Const TempMiscPayTableName As String = "TempMiscPay"
    Public Const TempPaymentsTableName As String = "TempPayments"
    Public Const TempReimHistTableName As String = "TempReimHist"
    Public Const TimecardTableName As String = "Timecard"
    Public Const TitlesTableName As String = "Titles"
    Public Const TravCabsTableName As String = "TravelCabs"
    Public Const TravDailyTableName As String = "TravelDaily"
    Public Const TravDailyTotTableName As String = "TravelDailyTot"
    Public Const TravelTableName As String = "Travel"
    Public Const TravEntTableName As String = "TravelEnt"
    Public Const TravForInvTableName As String = "TravForInv"
    Public Const TravPerMileTableName As String = "TravelPerMile"
    Public Const TravToBillTableName As String = "TravelToBill"
    Public Const UseCreditsTableName As String = "UseCredits"
    Public Const ViewCreditsTableName As String = "ViewCredits"
    Public Const WdEmpHourToBillTableName As String = "WdEmpHourToBill"
    Public Const WdHourlyToBillTableName As String = "WdHourlyToBill"
    Public Const WdPbServMsgTableName As String = "WdPbServMsg"
    Public Const WdPreBillTableName As String = "WdPreBill"
    Public Const WdTempTableName As String = "WdTemp"
    Public Const WorkDoneTableName As String = "WorkDone"
    Public Const WorkTypesTableName As String = "WorkTypes"

    Public Const TempTableAdd As String = "Temp"

    Public Const AlTableName As String = "AlTest"
    Public Const ThomTableName As String = "ThomTest"
    Public Const FfProbsTableName As String = "FixedFeeProbs"

#End Region

    '****************
    ' VIEW NAMES
    '****************

#Region " View Names "

    Public Const AllInvCurValsViewName As String = "AllInvCurValues"
    Public Const FixedFeeSumViewName As String = "FixedFeeSum"
    Public Const JobInvoiceSumViewName As String = "JobInvoiceSummary"
    Public Const SumOfChangesViewName As String = "SumOfChanges"
    Public Const SumOfPaymentsViewName As String = "SumOfPayments"

#End Region

    '****************
    ' RELATION NAMES
    '****************

#Region " Relation Names "

#Region " Company Relation Names "

    Public Const CompaniesCompanyRolesRelName As String = Pcm.CompaniesTableName & ".CompaniesCompanyRoles"
    Public Const CompaniesContactsRelName As String = Pcm.CompaniesTableName & ".CompaniesContacts"
    Public Const CompaniesCompContsRelName As String = Pcm.CompaniesTableName & ".CompaniesCompConts"
    Public Const CompBillAddrRelName As String = "CompaniesBillingAddr"

    Public Const CompaniesProjRolesFkRelName As String = "CompaniesProjRolesFk"

#End Region

#Region " Contact Relation Names "

    Public Const ContactsAffilRelName As String = _
            ContactsTableName & "." & ContactsTableName & ContactAffilTableName
    Public Const ContactsPlayersRelName As String = _
            ContactsTableName & "." & ContactsTableName & PlayersTableName

    Public Const ContactsPlayersFkRelName As String = ContactsTableName & "_" & PlayersTableName
    Public Const ContactsJobsFkRelName As String = ContactsTableName & "_" & JobsTableName

#End Region

#Region " Job Relation Names "

    Public Const JobsFixedFeeRelName As String = JobsTableName & ".JobsFixedFee"
    Public Const JobsInvoicesRelName As String = JobsTableName & ".JobsInvoices"

#End Region

#Region " Invoice Relation Names "

    Public Const InvoicePaymentsRelName As String = InvoicesTableName & ".InvoicesPayments"
    Public Const InvoiceInvRevsRelName As String = InvoicesTableName & ".InvoicesInvoiceRevs"

    Public Const RptInvEmpRelName As String = RptInvoicesTableName & ".RptInvoicesRptInvEmpHour"
    Public Const RptInvFfRelName As String = RptInvoicesTableName & ".RptInvoicesRptInvFf"
    Public Const RptInvHourlyRelName As String = RptInvoicesTableName & ".RptInvoicesRptInvHours"
    Public Const RptInvMsgRelName As String = RptInvoicesTableName & ".RptInvoicesRptInvMsg"
    Public Const RptInvReimRelName As String = RptInvoicesTableName & ".RptInvoicesRptInvReim"

    Public Const TempInvFfHistRelName As String = "TempInvTempInvFfHist"

    Public Const AdjFfInvRelName As String = "AdjFfInvAdjFfHist"

#End Region

#Region " PreBill Releation Names "

    Public Const PreBillFFeeRelName As String = Pcm.PreBillTableName & ".PreBillPbFixedFee"
    Public Const PreBillRptFFeeRelName As String = Pcm.PreBillTableName & ".PreBillPbRptFixedFee"
    Public Const PreBillHourlJobsRelName As String = Pcm.PbHourlyJobsTableName & "." & Pcm.PbHourlyJobsTableName & Pcm.PbHourlyEmpsTableName

#End Region

#Region " Project Relation Names "

    Public Const ProjectsProjBldgsRelName As String = ProjectsTableName & ".ProjectsProjBldgs"
    Public Const ProjectsProjCodesRelName As String = ProjectsTableName & ".ProjectsProjCodes"
    Public Const ProjectsProjDrawLogRelName As String = ProjectsTableName & ".ProjectsProjDrawLog"
    Public Const ProjectsProjElevsRelName As String = ProjectsTableName & ".ProjectsProjElevs"
    Public Const ProjectsProjEscsRelName As String = ProjectsTableName & ".ProjectsProjEscs"
    Public Const ProjectsProjHistRolesRelName As String = ProjectsTableName & ".ProjectsProjHistRoles"
    Public Const ProjectsProjMaintRelName As String = ProjectsTableName & ".ProjectsProjMaint"
    Public Const ProjectsProjMWalksRelName As String = ProjectsTableName & ".ProjectsProjMWalks"
    Public Const ProjectsProjRolesRelName As String = ProjectsTableName & ".ProjectsProjRoles"
    Public Const ProjectsProjScopeRelName As String = ProjectsTableName & ".ProjectsProjScope"
    Public Const ProjectsProjSpecFeatRelName As String = ProjectsTableName & ".ProjectsProjSpecFeat"

    Public Const ProjRolesPlayersRelName As String = "ProjRolesPlayers"
    Public Const ProjHistRolesPlayersRelName As String = "ProjHistRolesProjHistPlayers"

    ' grand child relations
    Public Const ProjectsProjRolesPlayersRelName As String = ProjectsProjRolesRelName & "." & ProjRolesPlayersRelName
    Public Const ProjectsProjHistRolesPlayersRelName As String = ProjectsProjHistRolesRelName & "." & ProjHistRolesPlayersRelName


    ' foreign key constrains
    Public Const CompaniesProjRolesFKeyRelName As String = "CompaniesProjRoles"
    Public Const ContactsPlayersFKeyRelName As String = "ContactsPlayers"

#End Region

#Region " Travel Relation Names "

    Public Const TravDailyRelName As String = "TravelTravelDaily"
    Public Const TravTravDailyRelName As String = Pcm.TravelTableName & "." & TravDailyRelName

    ' travel report relations
    Public Const RptTravDailyRelName As String = "RptTravelRptTravelDaily"
    Public Const RptTravTravDailyRelName As String = Pcm.RptTravTableName & "." & Pcm.RptTravDailyRelName

    ' RptTravel.RptTravelRptTravelDaily

    ' travel report grand child relations
    Public Const RptTravDailyCabsRelName As String = RptTravTravDailyRelName & ".RptTravelDaily_RptTravelCabs"
    Public Const RptTravDailyEntRelName As String = RptTravTravDailyRelName & ".RptTravelDaily_RptTravelEnt"

#End Region

#Region " Payments Relation Names "

    Public Const PaymentsCreditRelName As String = "PaymentsCredits"
    Public Const PaymentsInvRelName As String = "PaymentsInvPayments"
    Public Const PaymentsMiscRelName As String = "PaymentsMiscPayments"
    Public Const PaymentsSplitRelName As String = "PaymentsPayCompSplits"

#End Region

#End Region

    '****************
    ' COLUMN NAMES
    '****************

#Region " Column Names "

    Public Const NumMatched As String = "NUM_MATCHED"

#Region " AdjFfInv Column Names "

    Public Const AfiNewAmntColName As String = "NewAmount"
    ' other column names same as invoice

#End Region

#Region " Affiliation Column Names "

    Public Const AffilIdColName As String = "AffilId"
    Public Const AffilNameColName As String = "Name"

#End Region

#Region " AllInvCurValues "

    Public Const AicvCurAmountColName As String = "CurAmount"
    Public Const AicvCurBalColName As String = "CurBalance"
    Public Const AicvCurPaidColName As String = "CurPaid"
    Public Const AicvInvNoColName As String = Pcm.InvInvNoColName
    Public Const AicvOrigAmountColName As String = Pcm.InvOrigAmountColName

#End Region

#Region " AutoInc Column Names "

    Public Const AutoIdColName As String = "Id"
    Public Const AutoTblNameColName As String = "TableName"

#End Region

#Region " BillingAddr Column Names "

    Public Const BillAddrIdColName As String = "BaId"
    ' all other column names same as address column names in Companies

#End Region

#Region " BuildingTypes Column Names "

    Public Const BldgTypeIdColName As String = "BldgTypeId"
    Public Const BldgTypeColName As String = "Types"

#End Region

#Region " Cities Column Names "

    Public Const CitiesIdColName As String = "CityId"
    Public Const CitiesNameColName As String = "Name"

#End Region

#Region " CloudEdits Column Names "

    Public Const ceTblNameColName As String = "TableName"
    Public Const ceKeyValColName As String = "KeyValue"
    Public Const ceEditTypeColName As String = "EditType"
    Public Const ceSyncResultColName As String = "SyncResult"
    Public Const ceSyncMessageColName As String = "SyncMessage"

#End Region

#Region " Codes Column Names "

    Public Const CodesDescColName As String = "Description"
    Public Const CodesIdColName As String = "CodesId"
    Public Const CodesNameColName As String = "Name"

#End Region

#Region " CompAcct Column Names "

    Public Const CaAcctDateColName As String = "AcctDate"
    Public Const CaAmountColName As String = "Amount"
    Public Const CaBalanceColName As String = "Balance"
    Public Const CaInvChkNoColName As String = "InvChkNo"
    Public Const CaItemColName As String = "AcctItem"
    Public Const CaItemIdColName As String = "ItemId"
    Public Const CaTypeName As String = "Type"

#End Region

#Region " Companies Column Names "

    Public Const CompActiveColName As String = "Active"
    Public Const CompAddr1ColName As String = "Address1"
    Public Const CompAddr2ColName As String = "Address2"
    Public Const CompCADSoftColName As String = "CADSoftware"
    Public Const CompCityColName As String = "City"
    Public Const CompCommColName As String = "Comments"
    Public Const CompCountryColName As String = "Country"
    Public Const CompEwCustColName As String = "Customer"
    Public Const CompEMailColName As String = "EMail"
    Public Const CompFaxColName As String = "Fax"
    Public Const CompFocusIntColName As String = "FocusInternational"
    Public Const CompFocusLocalColName As String = "FocusLocal"
    Public Const CompFocusNatColName As String = "FocusNational"
    Public Const CompFormalNameColName As String = "FormalName"
    Public Const CompIdColName As String = "CompId"
    Public Const CompNameColName As String = "Name"
    Public Const CompPdoxColName As String = "PDox35"
    Public Const CompPhoneColName As String = "Phone"
    Public Const CompPostCodeColName As String = "PostalCode"
    Public Const CompProspectColName As String = "Prospect"
    Public Const CompStateColName As String = "StateRegion"
    Public Const CompWebPageColName As String = "WebPage"
    Public Const CompWpSoftColName As String = "WpSoftware"

#End Region

#Region " CompanyProjectRoles Column Names "

    Public Const CompPRCityColName As String = "City"
    Public Const CompPRCompIdColName As String = "CompId"
    Public Const CompPRNameColName As String = "Name"
    Public Const CompPRProjIdColName As String = "ProjId"

#End Region

#Region " CompanyRoles Column Names "

    Public Const CompRoleCompIdColName As String = "CompId"
    Public Const CompRoleIdColName As String = "CrId"
    Public Const CompRoleRoleColName As String = "Role"

#End Region

#Region " Contacts Column Names "

    Public Const ContActiveColName As String = "Active"
    Public Const ContCellPhoneColName As String = "CellPhone"
    Public Const ContCommentsColName As String = "Comments"
    Public Const ContCompIdColName As String = "CompId"
    Public Const ContEMailColName As String = "EMailAddress"
    Public Const ContFaxColName As String = "Fax"
    Public Const ContFirstNameColName As String = "FirstName"
    Public Const ContHomePhoneColName As String = "HomePhone"
    Public Const ContIdColName As String = "ContId"
    Public Const ContLastNameColName As String = "LastName"
    Public Const ContMiddleColName As String = "Middle"
    Public Const ContPdoxColName As String = "PDox35"
    Public Const ContPagerColName As String = "Pager"
    Public Const ContPrefixColName As String = "Prefix"
    Public Const ContPrimaryRoleColName As String = "PrimaryRole"
    Public Const ContSalutColName As String = "Salutation"
    Public Const ContSuffixColName As String = "Suffix"
    Public Const ContTitleColName As String = "Title"
    Public Const ContWorkPhoneColName As String = "WorkPhone"

#End Region

#Region " ContactAffiliation Column Names "

    Public Const ConAflContIdColName As String = "ContId"
    Public Const ConAflIdColName As String = "CaId"
    Public Const ConAflNameColName As String = "Name"

#End Region

#Region " ContactProjects Column Names "

    Public Const ContProjCityColName As String = "City"
    Public Const ContProjContIdColName As String = "ContId"
    Public Const ContProjNameColName As String = "Name"
    Public Const ContProjPrimColName As String = "PrimaryContact"

#End Region

#Region " Countries Column Names "

    Public Const CountryCapitalColName As String = "Capital"
    Public Const CountryCodeColName As String = "Code"
    Public Const CountryContinentColName As String = "Continent"
    Public Const CountryIdColName As String = "CountryId"
    Public Const CountryNameColName As String = "Name"
    Public Const CountryUnitsColName As String = "UnitsOfMeasure"

#End Region

#Region " CountTable Column Names "

    Public Const CtRowCountColName As String = "RowCount"

#End Region

#Region " Credits Column Names "

    Public Const CredAmountColName As String = "Amount"
    Public Const CredCreditIdColName As String = "CreditId"
    Public Const CredCreditNoColName As String = "CreditNo"
    Public Const CredCompIdColName As String = "CompId"
    Public Const CredCommentsColName As String = "Comments"
    Public Const CredCurBalColName As String = "CurrentBalance"
    Public Const CredEwJobIdColName As String = "EwJobId"
    Public Const CredPayIdColName As String = "PayId"

#End Region

#Region " Deliveries Column Names "

    Public Const DelivCarrierColName As String = "Carrier"
    Public Const DelivCostColName As String = "Cost"
    Public Const DelivDateColName As String = "DelivDate"
    Public Const DelivDoNoBillColName As String = "DoNoBill"
    Public Const DelivIdColName As String = "DelivId"
    Public Const DelivEwJobIdColName As String = "EwJobId"
    Public Const DelivEwJobId2ColName As String = "EwJobId2"
    Public Const DelivInvNoColName As String = "InvoiceNo"
    Public Const DelivJobNameColName As String = "JobName"
    Public Const DelivRecipColName As String = "Recipient"

#End Region

#Region " EmpHourHistory Column Names "

    Public Const EhhEmpNoColName As String = "EmpNo"
    Public Const EhhHoursColName As String = "Hours"
    Public Const EhhInvoiceNoColName As String = "InvoiceNo"
    Public Const EhhInvVerColName As String = "InvVersion"

#End Region

#Region " ElevContractors Column Names "

    Public Const ElevCntrsIdColName As String = "ElevCntrsId"
    Public Const ElevCntrsNameColName As String = "Name"

#End Region

#Region " ElevTypes Column Names "

    Public Const ElevTypeIdColName As String = "ElevTypeId"
    Public Const ElevTypeColName As String = "Type"

#End Region

#Region " EwEmployees Column Names "

    Public Const EwEmpActiveColName As String = "Active"
    Public Const EwEmpEndColName As String = "EndDate"
    Public Const EwEmpFirstNameColName As String = "FirstName"
    Public Const EwEmpInitColName As String = "Init"
    Public Const EwEmpJobTitleColName As String = "JobTitle"
    Public Const EwEmpLastNameColName As String = "LastName"
    Public Const EwEmpNoColName As String = "EmpNo"
    Public Const EwEmpPrincipleColName As String = "Principle"
    Public Const EwEmpStartColName As String = "StartDate"

#End Region

#Region " ExpDisp Column Names "

    Public Const ExpDAmountColName As String = "Amount"
    Public Const ExpDCompColName As String = "Company"
    Public Const ExpDDisColName As String = "Dis"
    Public Const ExpDEdIdColName As String = "EdId"
    Public Const ExpDEmpColName As String = "Employee"
    Public Const ExpDFieldColName As String = "FieldName"

#End Region

#Region " FeeToBill Column Names "

    Public Const Fee2BAmountColName As String = "Amount"
    Public Const Fee2BRateColName As String = "BillRate"
    Public Const Fee2BEwJobIdColName As String = "EwJobId"
    Public Const Fee2BFeeMaxColName As String = "FeeMax"
    Public Const Fee2BHoursColName As String = "Hours"
    Public Const Fee2BInvNoColName As String = "InvoiceNo"
    Public Const Fee2BOverMaxColName As String = "OverMax"
    Public Const Fee2BPriorAmntColName As String = "PriorAmount"
    Public Const Fee2BTotAmntColName As String = "TotalAmount"

#End Region

#Region " FfHistory Column Names "

    Public Const FfHAmountColName As String = "Amount"
    Public Const FfHAmtThisBillColName As String = "AmountThisBill"
    Public Const FfHAmtToDateColName As String = "AmountToDate"
    Public Const FfHIdColName As String = "FfId"
    Public Const FfHInvNoColName As String = Pcm.InvInvNoColName
    Public Const FfHInvVerColName As String = Pcm.InvVerColName
    Public Const FfHNameColName As String = "FfName"
    Public Const FfHOrderColName As String = "FfOrder"
    Public Const FfHPerThisBillColName As String = "PercentThisBill"
    Public Const FfHPerToDateColName As String = "PercentToDate"

#End Region

#Region " FfPreInv Column Names "

    Public Const PiAmntToDate As String = "PiAmntToDate"
    Public Const PiPerToDate As String = "PiPerToDate"

#End Region

#Region " FixedFee Column Names "

    Public Const FFeeAmountColName As String = "Amount"
    Public Const FFeeAmountBilled As String = "AmountBilled"
    Public Const FFeeAmtToDateColName As String = "AmountToDate"
    Public Const FFeeAmtThisBillColName As String = "AmountThisBill"
    Public Const FFeeEwJobIdColName As String = "EwJobId"
    Public Const FFeeIdColName As String = "FfId"
    Public Const FFeeNameColName As String = "FfName"
    Public Const FFeeOrderColName As String = "FfOrder"
    Public Const FFeePercentBilled As String = "PercentBilled"
    Public Const FFeePerToDateColName As String = "PercentToDate"
    Public Const FFeePerThisBillColName As String = "PercentThisBill"
    Public Const FFeeFflPerToDateColName As String = "FflPerToDate"
    Public Const FFeeFflAmntToDateColName As String = "FflAmntToDate"

#End Region

#Region " Fixed Fee Log Column Names "

    Public Const FflAmountColName As String = "FfAmount"
    Public Const FflAmntThisItemColName As String = "AmntThisItem"
    Public Const FflDateColName As String = "FflDate"
    Public Const FflDescripColName As String = "Description"
    Public Const FflFixedFeeIdColName As String = "FfId"
    Public Const FflIdColName As String = "FfLogId"
    Public Const FflInvNoColName As String = "InvoiceNo"
    Public Const FflPerThisItemColName As String = "PerThisItem"
    Public Const FflTypeIdColName As String = "TypeId"

    Public Const FflTypeTextColName As String = "TypeText"

#End Region

#Region " HourlyHistory Column Names "

    Public Const HhHoursColName As String = "Hours"
    Public Const HhInvoiceNoColName As String = Pcm.InvInvNoColName
    Public Const HhInvVerColName As String = Pcm.InvVerColName
    Public Const HhRateColName As String = "Rate"

#End Region

#Region " HourlyToBill Column Names "

    Public Const Hour2BEwJobIdColName As String = "EwJobId"
    Public Const Hour2BRateColName As String = "BillRate"
    Public Const Hour2BHoursColName As String = "Hours"

#End Region

#Region " IndvRoles Column Names "

    Public Const IndvRolesIdColName As String = "IrId"
    Public Const IndvRolesRoleColName As String = "Role"

#End Region

#Region " InvAddr Column Names "

    Public Const IaIdColName As String = "IaId"
    Public Const IaLine1ColName As String = "Line1"
    Public Const IaLine2ColName As String = "Line2"
    Public Const IaLine3ColName As String = "Line3"
    Public Const IaLine4ColName As String = "Line4"

#End Region

#Region " InvAmntHistory Column Names "

    Public Const IahAmountColName As String = "Amount"
    Public Const IahInvNoColName As String = Pcm.InvInvNoColName
    Public Const IahInvVerColName As String = Pcm.InvVerColName
    Public Const IahJcaVerColName As String = Pcm.InvJcaVerColName

#End Region

#Region " InvCalcStatus Column Names "

    Public Const InvCalcStatusColName As String = "Status"
    Public Const InvCalcEmpNoColName As String = "EmpNo"
    Public Const InvCalcFomColName As String = "FirstOfMonth"

#End Region

#Region " InvCanPay Column Names "

    Public Const IcpBalanceColName As String = Pcm.InvBalanceColName
    Public Const IcpBilledFromColName As String = Pcm.InvBilledFromColName
    Public Const IcpCompNameColName As String = "CompanyName"
    Public Const IcpCompCityColName As String = Pcm.CompCityColName
    Public Const IcpCompIdColName As String = Pcm.JobsCompIdColName
    Public Const IcpEwJobIdColName As String = Pcm.JobsEwJobIdColName
    Public Const IcpJobNameColName As String = Pcm.JobsJobNameColName
    Public Const IcpJobNoColName As String = Pcm.JobsJobNoColName
    Public Const IcpInvNoColName As String = Pcm.InvInvNoColName
    Public Const IcpOpenedByColName As String = Pcm.JobsOpenByColName
    Public Const IcpPaidColName As String = "Paid"
    Public Const IcpProjIdColName As String = Pcm.JobsProjIdColName
    Public Const IcpTypeColName As String = Pcm.InvTypeColName

#End Region

#Region " InvCurValues Column Names "

    Public Const IcvInvNoColName As String = Pcm.InvInvNoColName
    Public Const IcvAmountColName As String = "Amount"
    Public Const IcvBalanceColName As String = "Balance"
    Public Const IcvPaidColName As String = "Paid"

#End Region

#Region " InvFfRevs Column Names "

    Public Const IfrAmntThisBillColName As String = "AmountThisBill"
    Public Const IfrAmntToDateColName As String = "AmountToDate"
    Public Const IfrAmountColName As String = "Amount"
    Public Const IfrFfIdColName As String = "FfId"
    Public Const IfrFfNameColName As String = "FfName"
    Public Const IfrFfOrderColName As String = "FfOrder"
    Public Const IfrNewAmntColName As String = "NewAmount"
    Public Const IfrNewPerToDateColName As String = "NewPercentToDate"
    Public Const IfrPerToDateColName As String = "PercentToDate"
    Public Const IfrPerThisBillColName As String = "PercentThisBill"
    Public Const IfrRevAmntColName As String = "RevAmount"
    Public Const IfrRevIdColName As String = "RevisionId"

    Public Const IfrFflAmntToDateColName As String = "FflAmntToDate"
    Public Const IfrFflPerToDateColName As String = "FflPerToDate"

#End Region

#Region " InvJobCompAttn Column Names "

    Public Const IjAddr1ColName As String = "Address1"
    Public Const IjAddr2ColName As String = "Address2"
    Public Const IjCityColName As String = "City"
    Public Const IjClientJobNoColName As String = "ClientJobNo"
    Public Const IjCompNameColName As String = "CompanyName"
    Public Const IjCountryColName As String = "Country"
    Public Const IjFirstName As String = "FirstName"
    Public Const IjInvNoColName As String = Pcm.InvInvNoColName
    Public Const IjJcaVerColName As String = Pcm.InvJcaVerColName
    Public Const IjJobNameColName As String = "JobName"
    Public Const IjJobNoColName As String = "JobNo"
    Public Const IjLastName As String = "LastName"
    Public Const IjPostCodeColName As String = "PostalCode"
    Public Const IjPrefixColName As String = "Prefix"
    Public Const IjStateColName As String = "StateRegion"
    Public Const IjSuffixColName As String = "Suffix"

#End Region

#Region " Invoices Column Names "

    'Public Const InvAdjAmountColName As String = "AdjAmount"
    'Public Const InvAdjColName As String = "Adjustments"
    'Public Const InvAmountColName As String = "Amount"
    Public Const InvBalanceColName As String = "Balance"
    Public Const InvBilledFromColName As String = "BilledFrom"
    Public Const InvBilledToColName As String = "BilledTo"
    Public Const InvBillTypeColName As String = "BillType"
    Public Const InvDateBilledColName As String = "DateBilled"
    Public Const InvEwJobIdColName As String = "EwJobId"
    Public Const InvFomBilledColName As String = "FomBilled"
    Public Const InvHourlyTypeColName As String = "HourlyType"
    Public Const InvInvNoColName As String = "InvoiceNo"
    Public Const InvJcaVerColName As String = "JcaVersion"
    Public Const InvOrigAmountColName As String = "OrigAmount"
    'Public Const InvPaidColName As String = "Paid"
    Public Const InvTypeColName As String = "Type"
    Public Const InvVerColName As String = "InvVersion"

    Public Const InvAdjSumColName As String = "AdjustSum"
    Public Const InvAmountSumColName As String = "AmountSum"
    Public Const InvPaidSumColName As String = "PaidSum"

    Public Const InvAttnContIdColName As String = Pcm.JobsAttnContIdColName
    Public Const InvAcFirstNameColName As String = Pcm.ContFirstNameColName
    Public Const InvAcLastNameColName As String = Pcm.ContLastNameColName
    Public Const InvBaseJobNo As String = "BaseJobNo"
    Public Const InvClientJobNoColName As String = Pcm.JobsClientJobNoColName
    Public Const InvCompCityColName As String = Pcm.CompCityColName
    Public Const InvCompIdColName As String = Pcm.CompIdColName
    Public Const InvCompNameColName As String = "CompanyName"
    Public Const InvJobNameColName As String = Pcm.JobsJobNameColName
    Public Const InvJobNoColName As String = Pcm.JobsJobNoColName
    Public Const InvObFirstNameColName As String = "EmpFirstName"
    Public Const InvObLastNameColName As String = "EmpLastName"
    Public Const InvOpenedByColName As String = Pcm.JobsOpenByColName    
    Public Const InvProjIdColName As String = Pcm.JobsProjIdColName
    Public Const InvProjNameColName As String = "ProjectName"

#End Region

#Region " InvoiceRevs Column Names "

    Public Const InvRevAmountColName As String = "Amount"
    Public Const InvRevChangeColName As String = "Change"
    Public Const InvRevCommColName As String = "Comments"
    Public Const InvRevDateColName As String = "RevisionDate"
    Public Const InvRevInvNoColName As String = Pcm.InvInvNoColName
    Public Const InvRevInvVerColName As String = Pcm.InvVerColName
    Public Const InvRevIdColName As String = "RevisionId"    
    Public Const InvRevNewAmtColName As String = "NewAmount"
    Public Const InvRevObcColName As String = "OBC"

#End Region

#Region " InvPayments Column Names "

    Public Const InvPayAmountColName As String = "Amount"
    Public Const InvPayInvPayIdColName As String = "InvPayId"
    Public Const InvPayInvNoColName As String = "InvoiceNo"
    Public Const InvPayPayIdNoColName As String = "PayId"

#End Region

#Region " Jobs Column Names "

    Public Const JobsActiveColName As String = "Active"
    Public Const JobsAttnContIdColName As String = "AttnContId"
    Public Const JobsBillableColName As String = "Billable"
    Public Const JobsBillAddrTypeColName As String = "BillingAddrType"
    Public Const JobsBillRateColName As String = "BillRate"
    Public Const JobsBillToColName As String = "BillTo"
    Public Const JobsBillTypeColName As String = "BillType"
    Public Const JobsCityColName As String = "City"
    Public Const JobsClientJobNoColName As String = "ClientJobNo"
    Public Const JobsCommColName As String = "Comments"
    Public Const JobsCompIdColName As String = "CompId"
    Public Const JobsDeptColName As String = "Department"
    Public Const JobsEwJobIdColName As String = "EwJobId"
    Public Const JobsFeeColName As String = "FeeMax"
    Public Const JobsFilePathColName As String = "FilePath"
    Public Const JobsHourlyTypeColName As String = "HourlyType"
    Public Const JobsJobNameColName As String = "JobName"
    Public Const JobsJobNoColName As String = "JobNo"    
    Public Const JobsOpenByColName As String = "OpenedBy"
    Public Const JobsOpenOnColName As String = "OpenedOn"
    Public Const JobsProjIdColName As String = "ProjId"
    Public Const JobsRBkUpColName As String = "ReimbBackup"
    Public Const JobsRMkUpColName As String = "ReimbMkup"
    Public Const JobsServMsgColName As String = "ServiceMsg"
    Public Const JobsUseSmColName As String = "UseServiceMsg"

    Public Const JobsCompNameColName As String = "CompanyName"
    Public Const JobsCompAddColName As String = "CompanyAddress"
    Public Const JobsCompBillCityColName As String = "BillingCity"
    Public Const JobsCompCityColName As String = "CompanyCity"
    Public Const JobsCompStateColName As String = "CompanyState"
    Public Const JobsCompPostColName As String = "CompanyPostCode"
    Public Const JobsCountryColName As String = Pcm.CompCountryColName
    Public Const JobsAttnFirstColName As String = "AttnFirst"
    Public Const JobsAttnLastColName As String = "AttnLast"
    Public Const JobsOpByFirstColName As String = "OpenedByFirst"
    Public Const JobsOpByLastColName As String = "OpenedByLast"
    Public Const JobsProjNameColName As String = "ProjectName"
    Public Const JobsLastInvMonthColName As String = "LastInvMonth"
    Public Const JobsInvsAmountColName As String = "InvoicesAmount"
    Public Const JobsInvsBalanceColName As String = "InvoicesBalance"

#End Region

#Region " JobInvoiceSummary "

    Public Const JisMaxOfBilledFromColName As String = "MaxOf" & InvBilledFromColName
    Public Const JisSumOfCurAmountColName As String = "SumOf" & AicvCurAmountColName
    Public Const JisSumOfCurBalanceColName As String = "SumOf" & AicvCurBalColName

#End Region

#Region " JobInvSum Column Names "

    Public Const JobInvSumColName As String = "InvoiceSum"

#End Region

#Region " JobLastInvSum Column Names "

    Public Const JlisEwJobIdColName As String = JobsEwJobIdColName
    Public Const JlisMaxBilledFromColName As String = JobsLastInvMonthColName
    Public Const JlisSumAmountColName As String = JobsInvsAmountColName
    Public Const JlisSumBalanceColName As String = JobsInvsBalanceColName

#End Region

#Region " MiscPayments Column Names "

    Public Const MpAmountColName As String = "Amount"
    Public Const MpCompIdColName As String = "CompId"
    Public Const MpDescColName As String = "Description"
    Public Const MpEwJobIdColName As String = "EwJobId"
    Public Const MpIdColName As String = "MiscPayId"
    Public Const MpPayIdColName As String = "PayId"

#End Region

#Region " MsgHistory Column Names "

    Public Const MsgHistInvNoColName As String = "InvoiceNo"
    Public Const MsgHistInvVerColName As String = "InvVersion"
    Public Const MsgHistServMsgColName As String = "ServiceMsg"

#End Region

#Region " NewIfrTableName Column Names "

    ' use InvFfRevs column names

#End Region

#Region " NamePrefixes Column Names "

    Public Const NamePrefixIdColName As String = "NpId"
    Public Const NamePrefixNameColName As String = "Name"

#End Region

#Region " NotUsedIfrTableName Column Names "

    ' use InvFfRevs and Invoice column names

#End Region

#Region " OlSyncStatus Column Names  "

    Public Const olsEmpNoColName As String = Pcm.EwEmpNoColName
    Public Const olsLastSyncColName As String = "LastSync"

#End Region

#Region " OutlookContacts Column Names "

    Public Const OlContAddress1ColName As String = "Address1"
    Public Const OlContAddress2ColName As String = "Address2"
    Public Const OlContMobilePhoneColName As String = "MobilePhone"
    Public Const OlContCityColName As String = "City"
    Public Const OlContCompanyColName As String = "Company"
    Public Const OlContCompPhoneColName As String = "CompanyPhone"
    Public Const OlContCountryColName As String = "Country"
    Public Const OlContEMail1ColName As String = "EMail1Address"
    Public Const OlContEMail2ColName As String = "EMail2Address"
    Public Const OlContEMail3ColName As String = "EMail3Address"
    Public Const OlContEMailNoColName As String = "EMailNo"
    Public Const OlContFaxColName As String = "Fax"
    Public Const OlContFirstNameColName As String = "FirstName"
    Public Const OlContHomePhoneColName As String = "HomePhone"
    Public Const OlContIdColName As String = "Id"
    Public Const OlContJobTitleColName As String = "JobTitle"
    Public Const OlContLastNameColName As String = "LastName"
    Public Const OlContMiddleColName As String = "Middle"
    Public Const OlContNickNameColName As String = "NickName"
    Public Const OlContPagerColName As String = "Pager"
    Public Const OlContPostCodeColName As String = "PostalCode"
    Public Const OlContProfColName As String = "Profession"
    Public Const OlContStRegionColName As String = "StateRegion"
    Public Const OlContSuffixColName As String = "Suffix"
    Public Const OlContTitleColName As String = "Title"
    Public Const OlContWebPageColName As String = "WebPage"
    Public Const OlContWorkPhoneColName As String = "WorkPhone"

#End Region

#Region " PastDueHistory Column Names "

    Public Const PastDue30ColName As String = "ThirtyDays"
    Public Const PastDue60ColName As String = "SixtyDays"
    Public Const PastDue90ColName As String = "NinetyPlus"
    Public Const PastDueInvNoColName As String = "InvoiceNo"
    Public Const PastDueUnder30ColName As String = "UnderThirty"

#End Region

#Region " PayCompSplits Column Names "

    Public Const PcsAmountColName As String = "Amount"
    Public Const PcsCheckNoColName As String = Pcm.PayCheckNoColName
    Public Const PcsCityColName As String = Pcm.CompCityColName
    Public Const PcsCompIdColName As String = "CompId"
    Public Const PcsCreditColName As String = Pcm.PayCreditColName
    Public Const PcsDDateColName As String = Pcm.PayDDateColName
    Public Const PcsFeeColName As String = Pcm.PayFeeColName
    Public Const PcsIdColName As String = "PcsId"
    Public Const PcsMiscColName As String = Pcm.PayMiscColName
    Public Const PcsNameColName As String = Pcm.CompNameColName
    Public Const PcsPayIdColName As String = "PayId"
    Public Const PcsReimColName As String = Pcm.PayReimColName

#End Region

#Region " PayForm Column Names "

    ' only unique column names are here
    Public Const PayFJobNameColName As String = "JobName"
    Public Const PayFCompNameColName As String = "CompanyName"

#End Region

#Region " Payment Column Names "

    Public Const PayAmountColName As String = "Amount"
    Public Const PayCheckNoColName As String = "CheckNo"
    Public Const PayCompIdColName As String = "CompId"
    Public Const PayCreditColName As String = "Credit"
    Public Const PayCreditIdColName As String = "CreditId"
    Public Const PayDDateColName As String = "DepositDate"
    Public Const PayFeeColName As String = "Fee"
    Public Const PayIdColName As String = "PayId"
    Public Const PayMiscColName As String = "Misc"
    Public Const PayReimColName As String = "Reimbursable"

#End Region

#Region " PbDeliveies Column Names "

    Public Const PbDelivEwJobIdColName As String = Pcm.DelivEwJobIdColName
    Public Const PbDelivSumCostColName As String = "SumOf" & Pcm.DelivCostColName

#End Region

#Region " PbEmpHour Column Names "

    Public Const PehEwJobIdColName As String = "EwJobId"
    Public Const PehEmpNoColName As String = "EmpNo"
    Public Const PehHoursColName As String = "Hours"
    Public Const PehRateColName As String = "Rate"

#End Region

#Region " PbFixedFee Column Names "

    Public Const PbFFeeAmountColName As String = "Amount"
    Public Const PbFFeeAmtThisBillColName As String = "AmountThisBill"
    Public Const PbFFeeAmtToDateColName As String = "AmountToDate"
    Public Const PbFFeeEwJobIdColName As String = "EwJobId"
    Public Const PbFFeeIdColName As String = "FfId"
    Public Const PbFFeeNameColName As String = "FfName"
    Public Const PbFFeeOrderColName As String = "FfOrder"
    Public Const PbFFeePerThisBillColName As String = "PercentThisBill"
    Public Const PbFFeePerToDateColName As String = "PercentToDate"

#End Region

#Region " PbHourly Column Names "

    'Public Const PbhEwJobIdColName As String = "EwJobId"
    Public Const PbhHoursSumColName As String = "HoursSum"
    'Public Const PbhRateColName As String = "Rate"
    Public Const PbhOkToBillColName As String = "OkToBill"
    Public Const PbhToBillColName As String = "ToBill"

#End Region

#Region " PbServMsg Column Names "

    Public Const PbSmEwJobIdColName As String = Pcm.TcEwJobIdColName
    Public Const PbSmMsgColName As String = "Message"
    Public Const PbSmJobNoColName As String = Pcm.JobsJobNoColName
    Public Const PbSmProjNameColName As String = Pcm.ProjNameColName
    Public Const PbSmJobNameColName As String = Pcm.JobsJobNameColName

#End Region

#Region " PbTravel Column Names "

    Public Const PbTravEwJobIdColName As String = Pcm.TravEwJobIdColName
    Public Const PbTravSumTotalColName As String = "SumOf" & Pcm.TravTotalColName

#End Region

#Region " PeriodPastDue Column Names "

    Public Const PpdBalanceColName As String = "Balance"
    Public Const PpdEwJobNoColName As String = "EwJobId"

#End Region

#Region " Players Column Names "

    Public Const PlayersIdColName As String = "PlayerId"
    Public Const PlayersRoleIdColName As String = "RoleId"
    Public Const PlayersContIdColName As String = "ContId"
    Public Const PlayersPrimColName As String = "PrimaryContact"

    Public Const PlayersFirstNameUbColName As String = "FirstNameGridColumn"

#End Region

#Region " PreBill Column Names "

    Public Const PreBillEwJobIdColName As String = Pcm.TcEwJobIdColName
    Public Const PreBillFeeColName As String = Pcm.JobsFeeColName
    Public Const PreBillHoursColName As String = Pcm.TcHoursColName
    Public Const PreBillJobNameColName As String = Pcm.JobsJobNameColName
    Public Const PreBillJobNoColName As String = Pcm.JobsJobNoColName
    Public Const PreBillProjNameColName As String = Pcm.ProjNameColName

#End Region

#Region " PriorFeeBilling Column Names "

    Public Const PfBillEwJobIdColName As String = Pcm.Fee2BEwJobIdColName
    Public Const PfBillAmntSumColName As String = "SumOf" & Pcm.AicvCurAmountColName

#End Region

#Region " ProjBldgs Column Names "

    Public Const ProjBldgsBldgTypeColName As String = "BldgType"
    Public Const ProjBldgsIdColName As String = "ProjBldgsId"
    Public Const ProjBldgsProjIdColName As String = "ProjId"

#End Region

#Region " ProjCodes Column Names "

    Public Const ProjCodesCodeColName As String = "Code"
    Public Const ProjCodesIdColName As String = "ProjCodesId"
    Public Const ProjCodesProjIdColName As String = "ProjId"

#End Region

#Region " ProjDrawLog Column Names "

    Public Const ProjDLIdColName As String = "DrawLogId"
    Public Const ProjDLDateColName As String = "DrawDate"
    Public Const ProjDLPhaseColName As String = "DrawPhase"
    Public Const ProjDLProjIdColName As String = "ProjId"
    Public Const ProjDLSizeColName As String = "DrawSize"

#End Region

#Region " Projects Column Names "

    Public Const ProjAddrColName As String = "Address"
    Public Const ProjBdgtColName As String = "ProjBudget"
    Public Const ProjCityColName As String = "City"
    Public Const ProjCommColName As String = "Comments"
    Public Const ProjConsDocsColName As String = "ConsDocs"
    Public Const ProjConsPhaseColName As String = "ConsPhase"
    Public Const ProjCountryColName As String = "Country"
    Public Const ProjDsgnDevColName As String = "DesignDev"
    Public Const ProjEMailColName As String = "EMail"
    Public Const ProjEquipBdgtColName As String = "EquipBud"
    Public Const ProjEscsColName As String = "Escs"
    Public Const ProjEwcgPCaptColName As String = "EwcgProjCapt"
    Public Const ProjEwcgProjColName As String = "EwcgProject"
    Public Const ProjEwcgJobNoColName As String = "EwcgJobNo"
    Public Const ProjFilePathColName As String = "FilePath"
    Public Const ProjFirstEntColName As String = "FirstEntry"
    Public Const ProjFloorsColName As String = "Floors"
    Public Const ProjFtpColName As String = "FTPSite"
    Public Const ProjFtpPassColName As String = "FtpPassword"
    Public Const ProjFtpUserColName As String = "FtpUsername"
    Public Const ProjGsfColName As String = "GrossSqFeet"
    Public Const ProjGsmColName As String = "GrossSqMeters"
    Public Const ProjHeightFColName As String = "HeightFeet"
    Public Const ProjHeightMColName As String = "HeightMeters"
    Public Const ProjIdColName As String = "ProjId"
    Public Const ProjLastEntColName As String = "LastEntry"
    Public Const ProjLeadColName As String = "ProjectLead"
    Public Const ProjMmCommColName As String = "MmComm"
    Public Const ProjMWalksColName As String = "MovingWalks"
    Public Const ProjNameColName As String = "Name"
    Public Const ProjPdoxColName As String = "PDox35"
    Public Const ProjProjBdgtColName As String = "ProjBudget"
    Public Const ProjPropDateColName As String = "ProposalDate"
    Public Const ProjRentSqfColName As String = "RentableSqFeet"
    Public Const ProjRentSqmColName As String = "RentableSqMeters"
    Public Const ProjSchmDsgnColName As String = "SchemDesign"
    Public Const ProjStateColName As String = "StateRegion"
    Public Const ProjTotElevsColName As String = "TotalElevs"
    Public Const ProjTotEquipColName As String = "TotalEquip"
    Public Const ProjTPassElevColName As String = "TowerPassElevs"
    Public Const ProjTServElevColName As String = "TowerServiceElevs"
    Public Const ProjUofMeasColName As String = "UnitsOfMeasure"
    Public Const ProjURoomsColName As String = "UnitsRooms"
    Public Const ProjUseSqfColName As String = "UsableSqFeet"
    Public Const ProjUseSqmColName As String = "UsableSqMeters"
    Public Const ProjYrCompColName As String = "YearComplete"

#End Region

#Region " ProjElevs Column Names "

    Public Const ProjElevsBdgtColName As String = "Budget"
    Public Const ProjElevsCapColName As String = "Capacity"
    Public Const ProjElevsCntrColName As String = "Contractor"
    Public Const ProjElevsIdColName As String = "ElevId"
    Public Const ProjElevsFlrsColName As String = "Floors"
    Public Const ProjElevsNameColName As String = "Name"
    Public Const ProjElevsProjIdColName As String = "ProjId"
    Public Const ProjElevsQtyColName As String = "Quantity"
    Public Const ProjElevsSpeedColName As String = "Speed"
    Public Const ProjElevsTypeColName As String = "Type"

#End Region

#Region " ProjEscs Column Names "

    Public Const ProjEscsBdgtColName As String = "Budget"
    Public Const ProjEscsCntrColName As String = "Contractor"
    Public Const ProjEscsIdColName As String = "EscId"
    Public Const ProjEscsNameColName As String = "Name"
    Public Const ProjEscsProjIdColName As String = "ProjId"
    Public Const ProjEscsQtyColName As String = "Quantity"
    Public Const ProjEscsSpeedColName As String = "Speed"
    Public Const ProjEscsRiseColName As String = "Rise"
    Public Const ProjEscsWidthColName As String = "Width"

#End Region

#Region " ProjHistPlayers Column Names "

    Public Const ProjHPContIdColName As String = "ContId"
    Public Const ProjHPFirstNameColName As String = "FirstName"
    Public Const ProjHPIdColName As String = "ProjHPId"
    Public Const ProjHPLastNameColName As String = "LastName"
    Public Const ProjHPPrimColName As String = "PrimaryContact"
    Public Const ProjHPRoleIdColName As String = "PhRoleId"

#End Region

#Region " ProjHistRoles Column Names "

    Public Const ProjHRCityColName As String = "City"
    Public Const ProjHRCompIdColName As String = "CompId"    
    Public Const ProjHRNameColName As String = "Name"
    Public Const ProjHRProjIdColName As String = "ProjId"
    Public Const ProjHRRoleColName As String = "Role"
    Public Const ProjHRRoleIdColName As String = "PhRoleId"

#End Region

#Region " ProjMaint Column Names "

    Public Const ProjMaintCntrColName As String = "Contractor"
    Public Const ProjMaintDateColName As String = "MaintDate"
    Public Const ProjMaintIdColName As String = "MaintId"
    Public Const ProjMaintPerMonColName As String = "PerMonth"
    Public Const ProjMaintProjIdColName As String = "ProjId"
    Public Const ProjMaintTermColName As String = "Term"

#End Region

#Region " ProjMWalks Column Names "

    Public Const ProjMwBdgtColName As String = "Budget"
    Public Const ProjMwCntrColName As String = "Contractor"
    Public Const ProjMwIdColName As String = "MWalkId"
    Public Const ProjMwLengthColName As String = "Length"
    Public Const ProjMwNameColName As String = "Name"
    Public Const ProjMwProjIdColName As String = "ProjId"
    Public Const ProjMwQtyColName As String = "Quantity"
    Public Const ProjMwSpeedColName As String = "Speed"
    Public Const ProjMwWidthColName As String = "Width"

#End Region

#Region " ProjRoles Column Names "

    Public Const ProjRolesCompIdColName As String = "CompId"
    Public Const ProjRolesProjIdColName As String = "ProjId"
    Public Const ProjRolesRoleColName As String = "Role"
    Public Const ProjRolesRoleIdColName As String = "RoleId"

    Public Const ProjRolesCityUbColName As String = "CityGridColumn"

#End Region

#Region " ProjScope Column Names "

    Public Const ProjScopeBdgtColName As String = "Budget"
    Public Const ProjScopeCntrColName As String = "Contractor"
    Public Const ProjScopeDateColName As String = "ScopeDate"
    Public Const ProjScopeProjIdColName As String = "ProjId"
    Public Const ProjScopeIdColName As String = "ScopeId"
    Public Const ProjScopeWorkColName As String = "WorkDone"

#End Region

#Region " ProjSpecFeat Column Names "

    Public Const ProjSfFeatureColName As String = "Feature"
    Public Const ProjSfIdColName As String = "ProjSfId"
    Public Const ProjSfProjIdColName As String = "ProjId"

#End Region

#Region " ReimbSum Column Names "

    Public Const RSumInvNoColName As String = "InvoiceNo"
    Public Const RSumSumColName As String = "SumReimb"

#End Region

#Region " ReimHistory Column Names "

    Public Const RHistAmountColName As String = "Amount"
    Public Const RHistDelivColName As String = "Deliveries"
    Public Const RHistInvNoColName As String = "InvoiceNo"
    Public Const RHistInvVerColName As String = "InvVersion"
    Public Const RHistMarkupColName As String = "Markup"
    Public Const RHistReimbMkupColName As String = "ReimbMkup"
    Public Const RHistTravelColName As String = "Travel"

#End Region

#Region " ReimToBill Column Names "

    Public Const Reim2BAmountColName As String = "Amount"
    Public Const Reim2BDelivColName As String = "Deliveries"
    Public Const Reim2BEwJobIdColName As String = "EwJobId"
    Public Const Reim2BJobNoColName As String = "EwJobId"
    Public Const Reim2BInvNoColName As String = "InvoiceNo"
    Public Const Reim2BJobNameColName As String = "JobName"
    Public Const Reim2BMarkupColName As String = "Markup"
    Public Const Reim2BNameColName As String = "Name"
    Public Const Reim2BOkToBillColName As String = "OkToBill"
    Public Const Reim2BRMkupColName As String = "ReimbMkup"
    Public Const Reim2BTravelColName As String = "Travel"

#End Region

#Region " Roles Column Names "

    Public Const RoleIdColName As String = "RoleId"
    Public Const RoleRoleColName As String = "Role"

#End Region

#Region " RptInvoices Column Names "

    Public Const RptInvCurrentColName As String = "Current"

#End Region

#Region " RptInvFf Column Names "

    ' all other column names are same as FfHistory
    Public Const RiFfTotAmntColName As String = "TotalAmount"

#End Region

#Region " RptOutInv Column Names "

    'Public Const RoiAmountColName As String = Pcm.InvAmountColName
    Public Const RoiAmountColName As String = Pcm.InvOrigAmountColName
    Public Const RoiBalanceColName As String = Pcm.InvBalanceColName
    Public Const RoiFomBilledColName As String = Pcm.InvFomBilledColName
    Public Const RoiInvColName As String = Pcm.InvInvNoColName
    Public Const RoiJobNameColName As String = Pcm.JobsJobNameColName
    Public Const RoiJobNoColName As String = Pcm.JobsJobNoColName    

#End Region

#Region " SpecialFeatures Column Names "

    Public Const SpecFeatFeatureColName As String = "Feature"
    Public Const SpecFeatIdColName As String = "SpecFeatId"

#End Region

#Region " States Column Names "

    Public Const StatesIdColName As String = "StateId"
    Public Const StatesNameColName As String = "Name"

#End Region

#Region " Suffix Column Names "

    Public Const SuffixIdColName As String = "SuffixId"
    Public Const SuffixNameColName As String = "Name"

#End Region

#Region " Timecard Columns Names "

    Public Const TcCodeColName As String = "Code"
    Public Const TcCommentsColName As String = "Comments"
    Public Const TcDateColName As String = "TcDate"
    Public Const TcEmpNoColName As String = "EmpNo"
    Public Const TcEwJobIdColName As String = "EwJobId"
    Public Const TcEwJobId2ColName As String = "EwJobId2"
    Public Const TcHoursColName As String = "Hours"
    Public Const TcIdColName As String = "TcId"

#End Region

#Region " TcForm Column Names "

    Public Const TfCodeColName As String = Pcm.TcCodeColName
    Public Const TfCommentsColName As String = Pcm.TcCommentsColName
    Public Const TfDateColName As String = Pcm.TcDateColName
    Public Const TfEmpNoColName As String = Pcm.TcEmpNoColName
    Public Const TfEmpNo2ColName As String = Pcm.TcEmpNoColName & "2"
    Public Const TfEwJobIdColName As String = Pcm.TcEwJobIdColName
    Public Const TfEwJobId2ColName As String = Pcm.TcEwJobIdColName & "2"
    Public Const TfEwJobId3ColName As String = Pcm.TcEwJobIdColName & "3"
    Public Const TfFirstNameColName As String = Pcm.EwEmpFirstNameColName
    Public Const TfHoursColName As String = Pcm.TcHoursColName
    Public Const TfIdColName As String = Pcm.TcIdColName
    Public Const TfJobNameColName As String = Pcm.JobsJobNameColName
    Public Const TfLastNameColName As String = Pcm.EwEmpLastNameColName
    Public Const TfProjNameColName As String = Pcm.ProjNameColName

#End Region

#Region " Titles Column Names "

    Public Const TitlesIdColName As String = "TitlesId"
    Public Const TitlesNameColName As String = "Name"

#End Region

#Region " Travel Column Names "

    Public Const TravCommColName As String = "Comments"
    Public Const TravDoNoBillColName As String = "DoNoBill"
    Public Const TravEmpNoColName As String = "EmpNo"
    Public Const TravEmpNo2ColName As String = "EmpNo2"
    Public Const TravEndDateColName As String = "EndDate"
    Public Const TravEwJobIdColName As String = "EwJobId"
    Public Const TravEwJobId2ColName As String = "EwJobId2"
    Public Const TravFirstNameColName As String = Pcm.EwEmpFirstNameColName
    Public Const TravIdColName As String = "TravelId"
    Public Const TravInvNoColName As String = "InvoiceNo"
    Public Const TravItinColName As String = "Itinerary"
    Public Const TravJobNameCol As String = Pcm.JobsJobNameColName
    Public Const TravJobNoColName As String = Pcm.JobsJobNoColName
    Public Const TravLastNameColName As String = Pcm.EwEmpLastNameColName
    Public Const TravPerMileColName As String = "PerMile"
    Public Const TravProjNameColName As String = Pcm.ProjNameColName
    Public Const TravStartDateColName As String = "StartDate"
    Public Const TravTotalColName As String = "Total"

#End Region

#Region " TravelCabs Column Names "

    Public Const TCabDateColName As String = "CabDate"
    Public Const TCabFareColName As String = "Fare"
    Public Const TCabFromToColName As String = "FromTo"
    Public Const TCabIdColName As String = "CabId"
    Public Const TCabItemColName As String = "CabItem"    
    Public Const TCabTdIdColName As String = Pcm.TdIdColName

#End Region

#Region " TravelDaily Column Names "

    Public Const TdAirfareColName As String = "Airfare"
    Public Const TdAutoRentColName As String = "AutoRental"
    Public Const TdBreakfastColName As String = "Breakfast"
    Public Const TdCabsColName As String = "Cabs"
    Public Const TdDinnerColName As String = "Dinner"
    Public Const TdEntColName As String = "Ent"
    Public Const TdHotelColName As String = "Hotel"
    Public Const TdIdColName As String = "TdId"
    Public Const TdInvColName As String = Pcm.TravInvNoColName
    Public Const TdLunchColName As String = "Lunch"
    Public Const TdMileageColName As String = "Mileage"
    Public Const TdMileCostColName As String = "MileageCost"
    Public Const TdMiscColName As String = "Misc"
    Public Const TdParkingColName As String = "Parking"
    Public Const TdTollsColName As String = "Tolls"
    Public Const TdTotalColName As String = "Total"
    Public Const TdDateColName As String = "TravelDate"
    Public Const TdTravIdColName As String = Pcm.TravIdColName

#End Region

#Region " TravelDailyTotal Column Names "

    Public Const TdtAirfareColName As String = "Airfare"
    Public Const TdtAutoColName As String = "AutoRental"
    Public Const TdtBfastColName As String = "Breakfast"
    Public Const TdtCabsColName As String = "Cabs"
    Public Const TdtDinnerColName As String = "Dinner"
    Public Const TdtEntColName As String = "Ent"
    Public Const TdtHotelColName As String = "Hotel"
    Public Const TdtLunchColName As String = "Lunch"
    Public Const TdtMileageColName As String = "MileageCost"
    Public Const TdtMiscColName As String = "Misc"
    Public Const TdtParkingColName As String = "Parking"
    Public Const TdtTollsColName As String = "Tolls"
    Public Const TdtTravelIdColName As String = "TravelId"

#End Region

#Region " TravelEnt Column Names "

    Public Const TEntAmountColName As String = "Amount"
    Public Const TEntDateColName As String = "EntDate"
    Public Const TEntIdColName As String = "EntId"
    Public Const TEntItemColName As String = "EntItem"
    Public Const TEntPersonsColName As String = "Persons"
    Public Const TEntReasonColName As String = "Reason"
    Public Const TEntTitlesColName As String = "Titles"    
    Public Const TEntTdIdColName As String = Pcm.TdIdColName

#End Region

#Region " TravelPerMile Column Names "

    Public Const TpmIdColName As String = "ID"
    Public Const TpmPerMileColName As String = "PerMile"

#End Region

#Region " WdTemp Column Names "

    Public Const WdtActiveColName As String = Pcm.JobsActiveColName
    Public Const WdtAmountColName As String = "Amount"
    Public Const WdtBillRateColName As String = Pcm.JobsBillRateColName
    Public Const WdtBillTypeColName As String = Pcm.JobsBillTypeColName
    Public Const WdtEwJobIdColName As String = Pcm.JobsEwJobIdColName
    Public Const WdtHoursColName As String = "Hours"
    Public Const WdtJobNameColName As String = Pcm.JobsJobNameColName
    Public Const WdtJobNoColName As String = Pcm.JobsJobNoColName

#End Region

#Region " WorkDone Column Names "

    Public Const WdAmountColName As String = "Amount"
    Public Const WdCommentsColName As String = "Comments"
    Public Const WdEmpNoColName As String = "EmpNo"
    Public Const WdEmpNo2ColName As String = "EmpNo2"
    Public Const WdEwJobIdColName As String = "EwJobId"
    Public Const WdEwJobId2ColName As String = "EwJobId2"
    Public Const WdFomColName As String = "FirstOfMonth"
    Public Const WdHoursColName As String = "Hours"
    Public Const WdIdColName As String = "WorkDoneId"

#End Region

#Region " WorkTypes Column Names "

    Public Const WorkTypesColName As String = "Types"
    Public Const WorkTypesIdColName As String = "WorkTypesId"

#End Region

#End Region

#Region " Column Sizes "

    Public Const AffilNameSize As Integer = 30
    Public Const BldgTypeSize As Integer = 30
    Public Const CitiesNameSize As Integer = 50
    Public Const ElevCntrsNameSize As Integer = 50
    Public Const EwJobNameSize As Integer = 30
    Public Const IndvRolesRoleSize As Integer = 50
    Public Const NamePrefixNameSize As Integer = 20
    Public Const RoleRoleSize As Integer = 50
    Public Const SpecFeatFeatureSize As Integer = 50
    Public Const StatesNameSize As Integer = 30
    Public Const SuffixNameSize As Integer = 20
    Public Const TitlesNameSize As Integer = 100

#End Region

#Region " Child Grid Class "

    Public Const kftOther As Integer = -1
    Public Const kftNone As Integer = 0
    Public Const kftInteger As Integer = 1
    Public Const kftString As Integer = 2
    Public Const kftDateTime As Integer = 3

    Public Const lftOther As Integer = -1
    Public Const lftNone As Integer = 0
    Public Const lftInteger As Integer = 1
    Public Const lftString As Integer = 2
    Public Const lftDateTime As Integer = 3

    Public Const cgGridControlNotSetErr As Integer = -1
    Public Const cgGridViewNotSetErr As Integer = -2
    Public Const cgGridTableNotSetErr As Integer = -3
    Public Const cgKeyFieldNotSetErr As Integer = -4
    Public Const cgKeyColumnNotSetErr As Integer = -5
    Public Const cgKeyFieldTypeNotSetErr As Integer = -6
    Public Const cgLinkFieldNotSetErr As Integer = -7
    Public Const cgKeyFieldNotFoundErr As Integer = -8
    Public Const cgKeyFieldTypeInvalidErr As Integer = -9
    Public Const cgLinkFieldNotFoundErr As Integer = -10
    Public Const cgFieldNameNotSetErr As Integer = -11
    Public Const cgFieldNotFoundErr As Integer = -12
    Public Const cgInvalidRowHandleErr As Integer = -13
    Public Const cgGetRowFieldValueErr As Integer = -14

    Public Const cgIdFieldNotSetErr As Integer = -20
    Public Const cgIdColumnNotSetErr As Integer = -21
    Public Const cgIdFieldNotFoundErr As Integer = -22
    Public Const cgLookupFieldNotSetErr As Integer = -23
    Public Const cgLookupColumnNotSetErr As Integer = -24
    Public Const cgLookupFieldTypeNotSetErr As Integer = -25
    Public Const cgLookupFieldNotFoundErr As Integer = -26
    Public Const cgLookupFieldTypeInvalidErr As Integer = -27

    Public Const cgEndEditErr As Integer = -50
    Public Const cgOtherErr As Integer = -100

#End Region

#Region " Registry "

    Public Const rkPcdbRegKeyName As String = "Software\Ewcg\Pcdb"

    ' database/path 
    Public Const rkCopyToLocal As String = "CopyToLocal"
    Public Const rkDatabaseName As String = "DatabaseName"
    Public Const rkDataPath As String = "DataPath"
    Public Const rkLocalDataPath As String = "LocalDataPath"
    'Public Const rkMergeFromLocal As String = "MergeFromLocal"
    Public Const rkSyncFromLocal As String = "SyncFromLocal"
    Public Const rkTempDataPath As String = "TempDataPath"

    ' employee number
    Public Const rkEmpNo As String = "EmpNo"

    ' delivery default selection values
    Public Const rkDelivDateRange As String = "DeliveryDsvDateRange"
    Public Const rkDelivFiltNotBilled As String = "DeliveryDsvFilterNotBilled"
    Public Const rkDelivFiltCanBill As String = "DeliveryDsvFilterCanBill"
    Public Const rkDelivUseDefs As String = "DeliveryDsvUseDefaults"

    ' invoice default selection values
    Public Const rkInvDateRange As String = "InvoiceDsvDateRange"
    Public Const rkInvType As String = "InvoiceDsvType"
    Public Const rkInvFilterBalanced As String = "InvoiceDsvBalanced"
    Public Const rkInvUseDefs As String = "InvoiceDsvUseDefaults"

    ' payment default selection values
    Public Const rkPayDateRange As String = "PaymentDsvDateRange"    
    Public Const rkPayUseDefs As String = "PaymentDsvUseDefaults"

    ' timecard default selection values
    Public Const rkTimecardEmpId As String = "TimecardDsvEmpId"
    Public Const rkTimecardDateRange As String = "TimecardDsvDateRange"
    Public Const rkTimecardUseDefs As String = "TimecardDsvUseDefs"

    ' travel default selection values
    Public Const rkTravelDateRange As String = "TravelDsvDateRange"
    Public Const rkTravelFiltCanBill As String = "TravelDsvFilterCanBill"
    Public Const rkTravelUseDefs As String = "TravelDsvUseDefs"

    ' work done default selection values
    Public Const rkWorkDoneEmpId As String = "WorkDoneDsvEmpId"
    Public Const rkWorkDoneDateRange As String = "WorkDoneDsvDateRange"
    Public Const rkWorkDoneMonthPeriod As String = "WorkDoneDsvMonthPeriod"
    Public Const rkWorkDoneUseDefs As String = "WorkDoneDsvUseDefs"

#End Region

#Region " SQL "

    Public Const sqlCtInteger As String = "Integer"
    Public Const sqlCtText As String = "Text"
    Public Const sqlCtCurrency As String = "Currency"
    Public Const sqlCtDate As String = "Date"
    Public Const sqlCtDouble As String = "Double"
    Public Const sqlCtBoolean As String = "Logical"
    Public Const sqlCtMemo As String = "Memo"

    Public Const sqlWhereErrorText As String = "SQL WHERE ERROR"

    Public Const FfHistSelectSql As String = "SELECT Amount, AmountThisBill, AmountToDate, FfId, FfName, FfOrder, InvoiceNo, PercentThisBill, PercentToDate "
    Public Const FfHistFromSql As String = "FROM FfHistory "
    Public Const FfHistWhereEqualInvSql As String = "WHERE (InvoiceNo = ?)"
    Public Const FfHistWhereGreaterEqualInvSql As String = "WHERE (InvoiceNo >= ?)"

#End Region

#Region " Print "

    Public Const plError As Integer = -3
    Public Const plUserCancel As Integer = -2
    Public Const plNotSet As Integer = -1
    Public Const plPrinter As Integer = 0
    Public Const plDialog As Integer = 1
    Public Const plViewer As Integer = 2
    Public Const plControl As Integer = 3

    Public Const pUseLandscape As Boolean = True
    Public Const pUseProtrait As Boolean = False

#End Region

#Region " Travel Expense and Dispersal "

    Public Const InitPerMile As Double = 0.01

    Public Const tedMileage As Integer = 1
    Public Const tedTolls As Integer = 2
    Public Const tedParking As Integer = 3
    Public Const tedAuto As Integer = 4
    Public Const tedCabs As Integer = 5
    Public Const tedAir As Integer = 6
    Public Const tedBFast As Integer = 7
    Public Const tedLunch As Integer = 8
    Public Const tedDinner As Integer = 9
    Public Const tedHotel As Integer = 10
    Public Const tedMisc As Integer = 11
    Public Const tedEnt As Integer = 12
    Public Const tedCount As Integer = 12

#End Region

#Region " Payments "

    Public Const pdNewPayments As Boolean = True
    Public Const pdEditPayments As Boolean = False

    Public Const pdSplitCompanyName As String = "Split between Companies"
    Public Const pdCloseToZero As Double = 0.005            ' (1/2 cent)

    Public Const crCreditNoStart As String = "CR"
    Public Const crCreditNoFormat As String = "0000000"

#End Region

#Region " File/Ext/Folder/ Names "

    Public Const sdMyDocs As String = "My Documents"
    Public Const sdDeskTop As String = "Desktop"
    Public Const DataFileDefaultPath As String = Pcm.sdMyDocs & "\"
    Public Const LocalFolder As String = "PcdbLocal\"
    Public Const PcdbFolder As String = "Pcdb\"
    Public Const TempFolder As String = "Temp\"

    Public Const fxDataFile As String = "mdb"
    Public Const fxInvCons As String = "icf"
    Public Const fxInvDetl As String = "idf"
    Public Const fxInvFilter As String = "igf"
    Public Const fxJobFilter As String = "jgf"
    Public Const fxLayout As String = "lay"
    Public Const fxXml As String = "xml"

    Public Const FilterSep As String = "|"      ' filter seperator
    Public Const DataFileName As String = "Ewcg2000." & fxDataFile
    Public Const DataFileFilter As String = "EWCG Access Database file (" & DataFileName & ")" & FilterSep & DataFileName & FilterSep & "All Access Database files (*." & fxDataFile & ")" & FilterSep & "*." & fxDataFile
    Public Const DataFileTitle As String = "Open EWCG Access Database"

    Public Const ffInvFilter As String = "Invoice Grid Filters (*." & fxInvFilter & ")" & FilterSep & "*." & fxInvFilter
    Public Const ffJobFilter As String = "Job Grid Filters (*." & fxJobFilter & ")" & FilterSep & "*." & fxJobFilter
    Public Const ffLayoutFilter As String = "Pcdb Grid Layout Files (*." & fxLayout & ")" & FilterSep & "*." & fxLayout & FilterSep & "XML Files (*." & fxXml & ")" & FilterSep & "*." & fxXml & FilterSep & "All Files (*.*)" & FilterSep & "*.*"

    Public Const cfPcdb As String = "Pcdb"
    Public Const cfDefaultName As String = cfPcdb & "Config" & "." & fxXml


#End Region

#Region " Cloud Computing "

    Public Const ccNetwork As Integer = 0
    Public Const ccLocal As Integer = 1
    Public Const ccOutlook As Integer = 2

    Public Const cetNotSet As Integer = -1
    Public Const cetDelete As Integer = 0
    Public Const cetNew As Integer = 1
    Public Const cetEdit As Integer = 2

    Public Const cesNotSet As String = "Not Set"
    Public Const cesDelete As String = "Delete"
    Public Const cesNew As String = "Insert"
    Public Const cesEdit As String = "Update"

    Public Const cekTableNameIndex As Integer = 0
    Public Const cekKeyValueIndex As Integer = 1
    Public Const cekKeyCount As Integer = 2

    Public Const cstDeleted As Integer = 0
    Public Const cstInserted As Integer = 1
    Public Const cstUpdated As Integer = 2
    Public Const cstOutdated As Integer = 3

    Public Const ccBaseId As Integer = 1000000000    ' 1 billion 

    Public Const ccDataLocColName As String = "DataLocInt"
    Public Const ccUpdatedWhenColName As String = "UpdatedWhen"
    Public Const ccUpdatedByColName As String = "UpdatedByEmpNo"

    Public Const ccSaveParentToo As Boolean = True

#End Region

#Region " Outlook "

    Public Const olsCompBaseTableName As String = "CompOls"
    Public Const olsContBaseTableName As String = "ContOls"

    Public Const olsTableFormat As String = "00000"
    Public Const olsTableBaseExt As String = "00000"

    Public Const CompOlsBaseTableName As String = olsCompBaseTableName & olsTableBaseExt
    Public Const ContOlsBaseTableName As String = olsContBaseTableName & olsTableBaseExt

    Public Const olsSyncColName As String = "Sync"
    Public Const olsLinkedColName As String = "Linked"
    Public Const olsEmailNoColName As String = "EmailNo"
    Public Const olsSyncedWhen As String = "SyncedWhen"

#End Region

#Region " Pre Invoice "

    Public Const piUsePreInvoice As Integer = 0
    Public Const piUseBilledToDate As Integer = 1

#End Region

#Region " Merge Types "

    Public Const mtNone As Integer = -1
    Public Const mtProject As Integer = Pcm.biProjects  ' 0
    Public Const mtCompany As Integer = Pcm.biCompanies ' 1
    Public Const mtContact As Integer = Pcm.biContacts  ' 2

    Public Const mtParadoxProject As Integer = 1000
    Public Const mtPcdbProject As Integer = 1001

    Public Const mtPcdbSimilarCompanies As Integer = 1100
    Public Const mtParadoxCompany As Integer = 1101
    Public Const mtConvertParadoxCompany As Integer = 1102

    Public Const mtParadoxContact As Integer = 1200
    Public Const mtPcdbContact As Integer = 1201

#End Region

#Region " DoConfirm/NoConfirm "

    Public Const DoConfirm As Boolean = True
    Public Const NoConfirm As Boolean = False

#End Region

#Region " Configuration "

    ' Section Names
    Public Const snConfig As String = "Configuration"
    Public Const snDelivery As String = "Delivery"
    Public Const snInvoice As String = "Invoice"
    Public Const snPayment As String = "Payment"
    Public Const snSettings As String = "Settings"
    Public Const snTimecard As String = "Timecard"
    Public Const snTravel As String = "Travel"
    Public Const snWorkDone As String = "WorkDone"

    ' settings
    'Public Const vkCopyToLocal As String = "CopyToLocal"
    'Public Const vkDatabaseName As String = "DatabaseName"
    'Public Const vkDataPath As String = "DataPath"
    'Public Const vkEmpNo As String = "EmpNo"
    'Public Const vkLastFolder As String = "LastFolder"
    'Public Const vkLocalDataPath As String = "LocalDataPath"
    'Public Const vkMergeFromLocal As String = "MergeFromLocal"
    'Public Const vkSyncFromLocal As String = "SyncFromLocal"

    ' delivery 
    Public Const vkDelivDateRange As String = "DeliveryDateRange"
    Public Const vkDelivFiltNonBilled As String = "DeliveryFilterNonBilled"
    Public Const vkDelivFiltCanBill As String = "DeliveryFilterCanBill"
    Public Const vkDelivUseDefs As String = "DeliveryUseDefaults"

    ' invoice
    Public Const vkInvDateRange As String = "InvoiceDateRange"
    Public Const vkInvType As String = "InvoiceType"
    Public Const vkInvFilterBalanced As String = "InvoiceBalanced"
    Public Const vkInvUseDefs As String = "InvoiceUseDefaults"

    ' payment 
    Public Const vkPayDateRange As String = "PaymentDateRange"
    Public Const vkPayUseDefs As String = "PaymentUseDefaults"

    ' timecard 
    Public Const vkTimecardEmpId As String = "TimecardEmpId"
    Public Const vkTimecardDateRange As String = "TimecardDateRange"
    Public Const vkTimecardUseDefs As String = "TimecardUseDefs"

    ' travel 
    Public Const vkTravelDateRange As String = "TravelDateRange"
    Public Const vkTravelFiltCanBill As String = "TravelFilterCanBill"
    Public Const vkTravelUseDefs As String = "TravelUseDefs"

    ' work done default selection values
    Public Const vkWorkDoneEmpId As String = "WorkDoneEmpId"
    Public Const vkWorkDoneDateRange As String = "WorkDoneDateRange"
    Public Const vkWorkDoneMonthPeriod As String = "WorkDoneMonthPeriod"
    Public Const vkWorkDoneUseDefs As String = "WorkDoneUseDefs"

#End Region

#Region " Form Size/Loaction "

    Public Const MaxHeightAdj As Integer = 25
    Public Const DwonFormTop As Integer = 12

#End Region

End Module
