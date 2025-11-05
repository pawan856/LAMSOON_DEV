Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports NPOI.XSSF.UserModel
Imports NPOI.HPSF
Imports NPOI.POIFS.FileSystem
Imports NPOI.SS.Util
Imports NPOI.XSSF.Util


Partial Class RPT_SLCABLE_RCD_main
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private nf As New NPOIFuncX

    Protected Const IMG_PATH As String = "../../images/"
    Protected Const ROOT_PATH As String = "../../"
    Protected Const FUN_CODE As String = "RPT_STK_MOV"

    Protected isGenDownload As Boolean = False
    Private ExcelRow_num As Long = 0

    Dim customDateTimeFormat As DateTimeFormatInfo = New DateTimeFormatInfo()


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ar = New AccessRightUtils(FUN_CODE, Session("usr_id"), Me)

        If Session("usr_id") Is Nothing Or Session("usr_id") = "" Then
            Session.Remove("PAGE_SESSION_MENU_CODE")
            ar.Force_PageEndCtrlClear(Me, False)
            Me.Visible = False
            Exit Sub
        End If

        'Response.Write(Session("GENERIC_SESSION_COMPLETE_SQL"))
        'Response.Write("<br><br>")
        'Response.Write(Session("GENERIC_SESSION_SRCH_SQL"))
        'Response.Write("<br><br>")
        'Response.Write(Session("GENERIC_SESSION_SRCH_WHERE_SQL"))
        'Response.Write("<br><br>")
        'Response.Write(Session("GENERIC_SESSION_SRCH_GROUPBY_SQL"))
        'Response.Write("<br><br>")
        'Response.Write(Session("GENERIC_SESSION_SRCH_ORDERBY_SQL"))
        'Response.Write("<br><br>")

        'Session("GENERIC_SESSION_SRCH_WHERE_SQL")


        If Not IsPostBack Then
            'Response.Write(Session("SEARCH_SESSION_PAGE_FR_DATE_RANGE"))
            'Response.Write(Session("SEARCH_SESSION_PAGE_TO_DATE_RANGE"))
            'Response.Write(Session("SEARCH_SESSION_PAGE_DAY_MONTH"))
            'Response.Write(Session("SEARCH_SESSION_PAGE_TEAM_CODE"))

            isGenDownload = True
            tr_download.Style.Add("display", "none")
            ibtn_download.Visible = True


        End If

        ar.hideForm(Me)
    End Sub

    Protected Sub genDownloadFile()

        Dim result_dt As New DataTable
        Dim total_dt As New DataTable

        Dim ExcelWBObj As XSSFWorkbook
        Dim TemplateSheet, ReportSheet, NewSheet As XSSFSheet


        Dim pa1 As GlobalDBFunc.DBCmdPara


        Dim sql_string As String = ""
        Dim addSQL As String = ""


        Dim tmpltPath, tmpltFileName As String
        Dim tmpFileStream As FileStream
        Dim ReadFileStream As FileStream

        customDateTimeFormat.DateSeparator = "/"
        customDateTimeFormat.TimeSeparator = ":"
        customDateTimeFormat.ShortDatePattern = "dd/MM/yyyy"
        customDateTimeFormat.LongDatePattern = "dd/MM/yyyy"
        customDateTimeFormat.ShortTimePattern = "HH:mm"
        customDateTimeFormat.LongTimePattern = "HH:mm"
        customDateTimeFormat.FullDateTimePattern = "dd/MM/yyyy HH:mm"

        'tmpltPath = System.Configuration.ConfigurationManager.AppSettings.Item("EXCEL_TMPLT_PATH").ToString

        tmpltFileName = "STOCK_MOVE_RPT_TMPL.xlsx"

        tmpltPath = Server.MapPath(tmpltFileName)
        ReadFileStream = New FileStream(tmpltPath, FileMode.Open, FileAccess.Read)

        Dim sysFileName, sysFilePath, fileFolder As String

        fileFolder = HttpContext.Current.Cache("SYSP_TEMP_DIR")

        If Not System.IO.Directory.Exists(fileFolder) Then
            System.IO.Directory.CreateDirectory(fileFolder)
        End If


        Try
            Dim whereSQL As String = ""
            Dim appSQL As String = ""
            Dim IO_TYPE As String = ""

            pa1 = New GlobalDBFunc.DBCmdPara

            whereSQL = Session("GENERIC_SESSION_SRCH_WHERE_SQL")

            If Not String.IsNullOrWhiteSpace(whereSQL) Then
                whereSQL = " AND " & whereSQL
            End If

            IO_TYPE = Session("SEARCH_SESSION_PAGE_IO_TYPE")

            If Not String.IsNullOrWhiteSpace(IO_TYPE) Then
                Select Case IO_TYPE
                    Case "IST"
                        appSQL &= " AND IO_DOC='STF' AND SUBSTRING(IO_DOC_ID,1,2)='IT' "
                    Case "SRL"
                        appSQL &= " AND IO_DOC='STF' AND SUBSTRING(IO_DOC_ID,1,2)='RL' "
                    Case Else
                        appSQL &= " AND IO_DOC='" & gU.dbEncode(IO_TYPE) & "' "
                End Select
            End If


            sql_string = " SELECT WMS_IN_TX.IO_DATETIME, WMS_IN_TX.IO_TYPE,Convert(varchar,WMS_IN_TX.IO_DATETIME,103) as IO_DATETIME_D, WMS_IN_TX.IO_DOC, WMS_IN_TX.IO_DOC_ID, WMS_IN_TX.IO_BATCH_NO, WMS_ITEM.ITM_SKU_NO, " & _
                         " WMS_IN_TX.IO_QTY, Case when wms_item.itm_type='CABLE' then WMS_IN_S_TX.IOSX_UOM2 else WMS_ITEM.ITM_UOM end AS ITM_UOM, WMS_ITEM.ITM_PCS_PER_UOM, Case when wms_item.itm_type='CABLE' then WMS_IN_S_TX.IOSX_QTY2 else WMS_IN_TX.IO_QTY end as IN_QTY, null as OUT_QTY, Case when wms_item.itm_type='CABLE' then WMS_IN_S_TX.IOSX_QTY2 else WMS_IN_TX.IO_QTY end *isnull(WMS_ITEM.ITM_PCS_PER_UOM,1) as TOTAL_NUMBER,'' + GR_EDI_PO_NO as ORDER_NO, 'PO' as TRANS_TYPE, " & _
                         " WMS_IN_TX.IO_LOC, WMS_WAREHOUSE.WH_MAIN_WH,  WMS_IN_S_TX.IOSX_SERIAL_NO, WMS_IN_TX.ITM_CODE, WMS_IN_TX.PACK_KEY, " & _
                         " WMS_IN_TX.IO_WH, WMS_WH_BIN.BN_CSMS_CODE, WMS_IN_S_TX.IOSX_DRUM_ID, NULL as RT_TYPE, WMS_ITEM.ITM_TYPE " & _
                         " FROM WMS_IN_TX LEFT OUTER JOIN " & _
                         " WMS_ITEM ON WMS_IN_TX.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_IN_TX.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                         " WMS_IN_TX.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_IN_TX.PACK_KEY = WMS_ITEM.PACK_KEY INNER JOIN " & _
                         " WMS_GOODSRCV ON WMS_IN_TX.IMP_CODE = WMS_GOODSRCV.IMP_CODE AND WMS_IN_TX.STORER_CODE = WMS_GOODSRCV.STORER_CODE AND  " & _
                         " WMS_IN_TX.IO_DOC_ID = WMS_GOODSRCV.GR_CODE " & _
                         " LEFT OUTER JOIN WMS_IN_S_TX ON WMS_IN_TX.IMP_CODE = WMS_IN_S_TX.IMP_CODE AND WMS_IN_TX.STORER_CODE = WMS_IN_S_TX.STORER_CODE AND " & _
                         " WMS_IN_TX.IO_SYS_SEQ = WMS_IN_S_TX.IO_SYS_SEQ AND WMS_IN_S_TX.ITM_CODE = WMS_IN_TX.ITM_CODE AND WMS_IN_S_TX.PACK_KEY = WMS_IN_TX.PACK_KEY " & _
                         " LEFT OUTER JOIN WMS_WAREHOUSE ON WMS_IN_TX.IO_WH = WMS_WAREHOUSE.WH_CODE AND WMS_IN_TX.IMP_CODE = WMS_WAREHOUSE.IMP_CODE " & _
                         " LEFT OUTER JOIN WMS_WH_BIN ON WMS_IN_TX.IMP_CODE = WMS_WH_BIN.IMP_CODE AND WMS_IN_TX.IO_LOC = WMS_WH_BIN.LOC_KEY " & _
                         " WHERE IO_DOC='GR' " & whereSQL & appSQL & _
                         " UNION " & _
                         " SELECT WMS_OUT_TX.IO_DATETIME, WMS_OUT_TX.IO_TYPE, Convert(varchar,WMS_OUT_TX.IO_DATETIME,103) as IO_DATETIME_D, WMS_OUT_TX.IO_DOC + ' UNPOST', WMS_OUT_TX.IO_DOC_ID, WMS_OUT_TX.IO_BATCH_NO, WMS_ITEM.ITM_SKU_NO, " & _
                         " WMS_OUT_TX.IO_QTY, CASE WHEN wms_item.itm_type = 'CABLE' THEN WMS_OUT_S_TX.IOSX_UOM2 ELSE WMS_ITEM.ITM_UOM END, WMS_ITEM.ITM_PCS_PER_UOM, NULL as IN_QTY,CASE WHEN wms_item.itm_type = 'CABLE' THEN WMS_OUT_S_TX.IOSX_QTY2 ELSE WMS_OUT_TX.IO_QTY END as OUT_QTY, CASE WHEN wms_item.itm_type = 'CABLE' THEN WMS_OUT_S_TX.IOSX_QTY2 ELSE WMS_OUT_TX.IO_QTY END * isnull(WMS_ITEM.ITM_PCS_PER_UOM,1) * -1 as TOTAL_NUMBER,'' + GR_EDI_PO_NO as ORDER_NO, 'PO' as TRANS_TYPE, " & _
                         " WMS_OUT_TX.IO_LOC, WMS_WAREHOUSE.WH_MAIN_WH, WMS_OUT_S_TX.IOSX_SERIAL_NO, WMS_OUT_TX.ITM_CODE, WMS_OUT_TX.PACK_KEY, " & _
                         " WMS_OUT_TX.IO_WH, WMS_WH_BIN.BN_CSMS_CODE, WMS_OUT_S_TX.IOSX_DRUM_ID, NULL as RT_TYPE, WMS_ITEM.ITM_TYPE " & _
                         " FROM WMS_OUT_TX INNER JOIN " & _
                         " WMS_GOODSRCV ON WMS_OUT_TX.IMP_CODE = WMS_GOODSRCV.IMP_CODE AND WMS_OUT_TX.STORER_CODE = WMS_GOODSRCV.STORER_CODE AND  " & _
                         " WMS_OUT_TX.IO_DOC_ID = WMS_GOODSRCV.GR_CODE LEFT OUTER JOIN " & _
                         " WMS_ITEM ON WMS_OUT_TX.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_OUT_TX.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                         " WMS_OUT_TX.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_OUT_TX.PACK_KEY = WMS_ITEM.PACK_KEY " & _
                         " LEFT OUTER JOIN WMS_OUT_S_TX ON WMS_OUT_TX.IMP_CODE = WMS_OUT_S_TX.IMP_CODE AND WMS_OUT_TX.STORER_CODE = WMS_OUT_S_TX.STORER_CODE AND " & _
                         " WMS_OUT_TX.IO_SYS_SEQ = WMS_OUT_S_TX.IO_SYS_SEQ AND WMS_OUT_TX.ITM_CODE = WMS_OUT_S_TX.ITM_CODE AND WMS_OUT_TX.PACK_KEY = WMS_OUT_S_TX.PACK_KEY " & _
                         " LEFT OUTER JOIN WMS_WAREHOUSE ON WMS_OUT_TX.IO_WH = WMS_WAREHOUSE.WH_CODE AND WMS_OUT_TX.IMP_CODE = WMS_WAREHOUSE.IMP_CODE " & _
                         " LEFT OUTER JOIN WMS_WH_BIN ON WMS_OUT_TX.IMP_CODE = WMS_WH_BIN.IMP_CODE AND WMS_OUT_TX.IO_LOC = WMS_WH_BIN.LOC_KEY " & _
                         " where IO_DOC='GR' and IO_TYPE='UNPOST' " & whereSQL & appSQL & _
                         " UNION " & _
                         " SELECT IO_DATETIME, WMS_OUT_TX.IO_TYPE, Convert(varchar,WMS_OUT_TX.IO_DATETIME,103) as IO_DATETIME_D, WMS_OUT_TX.IO_DOC, WMS_OUT_TX.IO_DOC_ID, WMS_OUT_TX.IO_BATCH_NO, WMS_ITEM.ITM_SKU_NO, " & _
                         " WMS_OUT_TX.IO_QTY, CASE WHEN wms_item.itm_type = 'CABLE' THEN WMS_OUT_S_TX.IOSX_UOM2 ELSE WMS_ITEM.ITM_UOM END, WMS_ITEM.ITM_PCS_PER_UOM, NULL as IN_QTY,CASE WHEN wms_item.itm_type = 'CABLE' THEN WMS_OUT_S_TX.IOSX_QTY2 ELSE WMS_OUT_TX.IO_QTY END as OUT_QTY, CASE WHEN wms_item.itm_type = 'CABLE' THEN WMS_OUT_S_TX.IOSX_QTY2 ELSE WMS_OUT_TX.IO_QTY END * isnull(WMS_ITEM.ITM_PCS_PER_UOM,1) * -1 as TOTAL_NUMBER,'' + DO_EDI_SIR_NO as ORDER_NO, 'SIR' as TRANS_TYPE, " & _
                         " WMS_OUT_TX.IO_LOC, WMS_WAREHOUSE.WH_MAIN_WH, WMS_OUT_S_TX.IOSX_SERIAL_NO, WMS_OUT_TX.ITM_CODE, WMS_OUT_TX.PACK_KEY, " & _
                         " WMS_OUT_TX.IO_WH, WMS_WH_BIN.BN_CSMS_CODE, WMS_OUT_S_TX.IOSX_DRUM_ID, NULL as RT_TYPE, WMS_ITEM.ITM_TYPE " & _
                         " FROM WMS_OUT_TX INNER JOIN " & _
                         " WMS_DELV_ORDER ON WMS_OUT_TX.IMP_CODE = WMS_DELV_ORDER.IMP_CODE AND WMS_OUT_TX.STORER_CODE = WMS_DELV_ORDER.STORER_CODE AND  " & _
                         " WMS_OUT_TX.IO_DOC_ID = WMS_DELV_ORDER.DO_CODE LEFT OUTER JOIN " & _
                         " WMS_ITEM ON WMS_OUT_TX.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_OUT_TX.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                         " WMS_OUT_TX.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_OUT_TX.PACK_KEY = WMS_ITEM.PACK_KEY " & _
                         " LEFT OUTER JOIN WMS_OUT_S_TX ON WMS_OUT_TX.IMP_CODE = WMS_OUT_S_TX.IMP_CODE AND WMS_OUT_TX.STORER_CODE = WMS_OUT_S_TX.STORER_CODE AND " & _
                         " WMS_OUT_TX.IO_SYS_SEQ = WMS_OUT_S_TX.IO_SYS_SEQ AND WMS_OUT_TX.ITM_CODE = WMS_OUT_S_TX.ITM_CODE AND WMS_OUT_TX.PACK_KEY = WMS_OUT_S_TX.PACK_KEY " & _
                         " LEFT OUTER JOIN WMS_WAREHOUSE ON WMS_OUT_TX.IO_WH = WMS_WAREHOUSE.WH_CODE AND WMS_OUT_TX.IMP_CODE = WMS_WAREHOUSE.IMP_CODE " & _
                         " LEFT OUTER JOIN WMS_WH_BIN ON WMS_OUT_TX.IMP_CODE = WMS_WH_BIN.IMP_CODE AND WMS_OUT_TX.IO_LOC = WMS_WH_BIN.LOC_KEY " & _
                         " where IO_DOC='DO' " & whereSQL & appSQL & _
                         " UNION " & _
                         " SELECT IO_DATETIME, WMS_IN_TX.IO_TYPE,Convert(varchar,WMS_IN_TX.IO_DATETIME,103) as IO_DATETIME_D, WMS_IN_TX.IO_DOC + ' UNPOST', WMS_IN_TX.IO_DOC_ID, WMS_IN_TX.IO_BATCH_NO, WMS_ITEM.ITM_SKU_NO, " & _
                         " WMS_IN_TX.IO_QTY, Case when wms_item.itm_type='CABLE' then WMS_IN_S_TX.IOSX_UOM2 else WMS_ITEM.ITM_UOM end AS ITM_UOM, WMS_ITEM.ITM_PCS_PER_UOM, Case when wms_item.itm_type='CABLE' then WMS_IN_S_TX.IOSX_QTY2 else WMS_IN_TX.IO_QTY end as IN_QTY, null as OUT_QTY, Case when wms_item.itm_type='CABLE' then WMS_IN_S_TX.IOSX_QTY2 else WMS_IN_TX.IO_QTY end *isnull(WMS_ITEM.ITM_PCS_PER_UOM,1) as TOTAL_NUMBER,'' + DO_EDI_SIR_NO as ORDER_NO, 'SIR' as TRANS_TYPE, " & _
                         " WMS_IN_TX.IO_LOC, WMS_WAREHOUSE.WH_MAIN_WH,  WMS_IN_S_TX.IOSX_SERIAL_NO, WMS_IN_TX.ITM_CODE, WMS_IN_TX.PACK_KEY, " & _
                         " WMS_IN_TX.IO_WH, WMS_WH_BIN.BN_CSMS_CODE, WMS_IN_S_TX.IOSX_DRUM_ID, NULL as RT_TYPE, WMS_ITEM.ITM_TYPE " & _
                         " FROM WMS_IN_TX LEFT OUTER JOIN " & _
                         " WMS_ITEM ON WMS_IN_TX.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_IN_TX.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                         " WMS_IN_TX.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_IN_TX.PACK_KEY = WMS_ITEM.PACK_KEY INNER JOIN " & _
                         " WMS_DELV_ORDER ON WMS_IN_TX.IMP_CODE = WMS_DELV_ORDER.IMP_CODE AND WMS_IN_TX.STORER_CODE = WMS_DELV_ORDER.STORER_CODE AND  " & _
                         " WMS_IN_TX.IO_DOC_ID = WMS_DELV_ORDER.DO_CODE " & _
                         " LEFT OUTER JOIN WMS_IN_S_TX ON WMS_IN_TX.IMP_CODE = WMS_IN_S_TX.IMP_CODE AND WMS_IN_TX.STORER_CODE = WMS_IN_S_TX.STORER_CODE AND " & _
                         " WMS_IN_TX.IO_SYS_SEQ = WMS_IN_S_TX.IO_SYS_SEQ AND WMS_IN_S_TX.ITM_CODE = WMS_IN_TX.ITM_CODE AND WMS_IN_S_TX.PACK_KEY = WMS_IN_TX.PACK_KEY " & _
                         " LEFT OUTER JOIN WMS_WAREHOUSE ON WMS_IN_TX.IO_WH = WMS_WAREHOUSE.WH_CODE AND WMS_IN_TX.IMP_CODE = WMS_WAREHOUSE.IMP_CODE " & _
                         " LEFT OUTER JOIN WMS_WH_BIN ON WMS_IN_TX.IMP_CODE = WMS_WH_BIN.IMP_CODE AND WMS_IN_TX.IO_LOC = WMS_WH_BIN.LOC_KEY " & _
                         " WHERE IO_DOC='DO' and IO_TYPE='UNPOST' " & whereSQL & appSQL & _
                         " UNION " & _
                         " SELECT IO_DATETIME, WMS_IN_TX.IO_TYPE,Convert(varchar,WMS_IN_TX.IO_DATETIME,103) as IO_DATETIME_D, WMS_IN_TX.IO_DOC, WMS_IN_TX.IO_DOC_ID, WMS_IN_TX.IO_BATCH_NO, WMS_ITEM.ITM_SKU_NO, " & _
                         " WMS_IN_TX.IO_QTY, Case when wms_item.itm_type='CABLE' then WMS_IN_S_TX.IOSX_UOM2 else WMS_ITEM.ITM_UOM end AS ITM_UOM, WMS_ITEM.ITM_PCS_PER_UOM, Case when wms_item.itm_type='CABLE' then WMS_IN_S_TX.IOSX_QTY2 else WMS_IN_TX.IO_QTY end as IN_QTY, null as OUT_QTY, Case when wms_item.itm_type='CABLE' then WMS_IN_S_TX.IOSX_QTY2 else WMS_IN_TX.IO_QTY end *isnull(WMS_ITEM.ITM_PCS_PER_UOM,1) as TOTAL_NUMBER,'' + RT_REF_NO2 as ORDER_NO, 'Credit Form' as TRANS_TYPE, " & _
                         " WMS_IN_TX.IO_LOC, WMS_WAREHOUSE.WH_MAIN_WH,  WMS_IN_S_TX.IOSX_SERIAL_NO, WMS_IN_TX.ITM_CODE, WMS_IN_TX.PACK_KEY, " & _
                         " WMS_IN_TX.IO_WH, WMS_WH_BIN.BN_CSMS_CODE, WMS_IN_S_TX.IOSX_DRUM_ID, WMS_COL_CODE.COLC_ENG_VALUE as RT_TYPE, WMS_ITEM.ITM_TYPE " & _
                         " FROM WMS_IN_TX LEFT OUTER JOIN " & _
                         " WMS_ITEM ON WMS_IN_TX.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_IN_TX.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                         " WMS_IN_TX.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_IN_TX.PACK_KEY = WMS_ITEM.PACK_KEY INNER JOIN " & _
                         " WMS_STOCK_RETURN ON WMS_IN_TX.IMP_CODE = WMS_STOCK_RETURN.IMP_CODE AND WMS_IN_TX.STORER_CODE = WMS_STOCK_RETURN.STORER_CODE AND  " & _
                         " WMS_IN_TX.IO_DOC_ID = WMS_STOCK_RETURN.RT_CODE " & _
                         " LEFT OUTER JOIN WMS_COL_CODE ON WMS_STOCK_RETURN.RT_TYPE = WMS_COL_CODE.COLC_CODE AND WMS_COL_CODE.COLC_TABCOL='WMS_STOCK_RETURN.RT_TYPE' " & _
                         " LEFT OUTER JOIN WMS_IN_S_TX ON WMS_IN_TX.IMP_CODE = WMS_IN_S_TX.IMP_CODE AND WMS_IN_TX.STORER_CODE = WMS_IN_S_TX.STORER_CODE AND " & _
                         " WMS_IN_TX.IO_SYS_SEQ = WMS_IN_S_TX.IO_SYS_SEQ AND WMS_IN_S_TX.ITM_CODE = WMS_IN_TX.ITM_CODE AND WMS_IN_S_TX.PACK_KEY = WMS_IN_TX.PACK_KEY " & _
                         " LEFT OUTER JOIN WMS_WAREHOUSE ON WMS_IN_TX.IO_WH = WMS_WAREHOUSE.WH_CODE AND WMS_IN_TX.IMP_CODE = WMS_WAREHOUSE.IMP_CODE " & _
                         " LEFT OUTER JOIN WMS_WH_BIN ON WMS_IN_TX.IMP_CODE = WMS_WH_BIN.IMP_CODE AND WMS_IN_TX.IO_LOC = WMS_WH_BIN.LOC_KEY " & _
                         " WHERE IO_DOC='SR' " & whereSQL & appSQL & _
                         " UNION " & _
                         " SELECT WMS_OUT_TX.IO_DATETIME, WMS_OUT_TX.IO_TYPE, Convert(varchar,WMS_OUT_TX.IO_DATETIME,103) as IO_DATETIME_D, 'Stock Return UNPOST', WMS_OUT_TX.IO_DOC_ID, WMS_OUT_TX.IO_BATCH_NO, WMS_ITEM.ITM_SKU_NO, " & _
                         " WMS_OUT_TX.IO_QTY, Case when wms_item.itm_type = 'CABLE' then WMS_OUT_S_TX.IOSX_UOM2 else WMS_ITEM.ITM_UOM end AS ITM_UOM, WMS_ITEM.ITM_PCS_PER_UOM, NULL as IN_QTY, CASE WHEN wms_item.itm_type = 'CABLE' THEN WMS_OUT_S_TX.IOSX_QTY2 ELSE WMS_OUT_TX.IO_QTY END as OUT_QTY, Case when wms_item.itm_type = 'CABLE' then WMS_OUT_S_TX.IOSX_QTY2 else WMS_OUT_TX.IO_QTY end * isnull(WMS_ITEM.ITM_PCS_PER_UOM,1) * -1 as TOTAL_NUMBER,'' + RT_REF_NO2 as ORDER_NO, 'Credit Form' as TRANS_TYPE, " & _
                         " WMS_OUT_TX.IO_LOC, WMS_WAREHOUSE.WH_MAIN_WH,  WMS_OUT_S_TX.IOSX_SERIAL_NO, WMS_OUT_TX.ITM_CODE, WMS_OUT_TX.PACK_KEY, " & _
                         " WMS_OUT_TX.IO_WH, WMS_WH_BIN.BN_CSMS_CODE, WMS_OUT_S_TX.IOSX_DRUM_ID, WMS_COL_CODE.COLC_ENG_VALUE as RT_TYPE, WMS_ITEM.ITM_TYPE " & _
                         " FROM WMS_OUT_TX LEFT OUTER JOIN " & _
                         " WMS_ITEM ON WMS_OUT_TX.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_OUT_TX.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                         " WMS_OUT_TX.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_OUT_TX.PACK_KEY = WMS_ITEM.PACK_KEY INNER JOIN " & _
                         " WMS_STOCK_RETURN ON WMS_OUT_TX.IMP_CODE = WMS_STOCK_RETURN.IMP_CODE AND WMS_OUT_TX.STORER_CODE = WMS_STOCK_RETURN.STORER_CODE AND  " & _
                         " WMS_OUT_TX.IO_DOC_ID = WMS_STOCK_RETURN.RT_CODE " & _
                         " LEFT OUTER JOIN WMS_COL_CODE ON WMS_STOCK_RETURN.RT_TYPE = WMS_COL_CODE.COLC_CODE AND WMS_COL_CODE.COLC_TABCOL='WMS_STOCK_RETURN.RT_TYPE' " & _
                         " LEFT OUTER JOIN WMS_OUT_S_TX ON WMS_OUT_TX.IMP_CODE = WMS_OUT_S_TX.IMP_CODE AND WMS_OUT_TX.STORER_CODE = WMS_OUT_S_TX.STORER_CODE AND " & _
                         " WMS_OUT_TX.IO_SYS_SEQ = WMS_OUT_S_TX.IO_SYS_SEQ AND WMS_OUT_S_TX.ITM_CODE = WMS_OUT_TX.ITM_CODE AND WMS_OUT_S_TX.PACK_KEY = WMS_OUT_TX.PACK_KEY " & _
                         " LEFT OUTER JOIN WMS_WAREHOUSE ON WMS_OUT_TX.IO_WH = WMS_WAREHOUSE.WH_CODE AND WMS_OUT_TX.IMP_CODE = WMS_WAREHOUSE.IMP_CODE " & _
                         " LEFT OUTER JOIN WMS_WH_BIN ON WMS_OUT_TX.IMP_CODE = WMS_WH_BIN.IMP_CODE AND WMS_OUT_TX.IO_LOC = WMS_WH_BIN.LOC_KEY " & _
                         " WHERE IO_DOC='SR' and IO_TYPE='UNPOST' " & whereSQL & appSQL & _
                         " UNION " & _
                         " SELECT IO_DATETIME, WMS_OUT_TX.IO_TYPE, Convert(varchar,WMS_OUT_TX.IO_DATETIME,103) as IO_DATETIME_D, WMS_OUT_TX.IO_DOC, WMS_OUT_TX.IO_DOC_ID, WMS_OUT_TX.IO_BATCH_NO, WMS_ITEM.ITM_SKU_NO, " & _
                         " WMS_OUT_TX.IO_QTY, CASE WHEN wms_item.itm_type = 'CABLE' THEN WMS_OUT_S_TX.IOSX_UOM2 ELSE WMS_ITEM.ITM_UOM END, WMS_ITEM.ITM_PCS_PER_UOM, NULL as IN_QTY,CASE WHEN wms_item.itm_type = 'CABLE' THEN WMS_OUT_S_TX.IOSX_QTY2 ELSE WMS_OUT_TX.IO_QTY END as OUT_QTY, CASE WHEN wms_item.itm_type = 'CABLE' THEN WMS_OUT_S_TX.IOSX_QTY2 ELSE WMS_OUT_TX.IO_QTY END * isnull(WMS_ITEM.ITM_PCS_PER_UOM,1) * -1 as TOTAL_NUMBER,'' + WMS_OUT_TX.IO_DOC_ID as ORDER_NO, WMS_COL_CODE.COLC_ENG_VALUE as TRANS_TYPE, " & _
                         " WMS_OUT_TX.IO_LOC, WMS_WAREHOUSE.WH_MAIN_WH, WMS_OUT_S_TX.IOSX_SERIAL_NO, WMS_OUT_TX.ITM_CODE, WMS_OUT_TX.PACK_KEY, " & _
                         " WMS_OUT_TX.IO_WH, WMS_WH_BIN.BN_CSMS_CODE, WMS_OUT_S_TX.IOSX_DRUM_ID, NULL as RT_TYPE, WMS_ITEM.ITM_TYPE " & _
                         " FROM WMS_OUT_TX INNER JOIN " & _
                         " WMS_STOCK_ISSUE ON WMS_OUT_TX.IMP_CODE = WMS_STOCK_ISSUE.IMP_CODE AND WMS_OUT_TX.STORER_CODE = WMS_STOCK_ISSUE.STORER_CODE AND  " & _
                         " WMS_OUT_TX.IO_DOC_ID = WMS_STOCK_ISSUE.IS_CODE LEFT OUTER JOIN " & _
                         " WMS_ITEM ON WMS_OUT_TX.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_OUT_TX.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                         " WMS_OUT_TX.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_OUT_TX.PACK_KEY = WMS_ITEM.PACK_KEY " & _
                         " LEFT OUTER JOIN WMS_COL_CODE ON WMS_STOCK_ISSUE.IS_TYPE = WMS_COL_CODE.COLC_CODE AND WMS_COL_CODE.COLC_TABCOL='WMS_STOCK_ISSUE.IS_TYPE' " & _
                         " LEFT OUTER JOIN WMS_OUT_S_TX ON WMS_OUT_TX.IMP_CODE = WMS_OUT_S_TX.IMP_CODE AND WMS_OUT_TX.STORER_CODE = WMS_OUT_S_TX.STORER_CODE AND " & _
                         " WMS_OUT_TX.IO_SYS_SEQ = WMS_OUT_S_TX.IO_SYS_SEQ AND WMS_OUT_TX.ITM_CODE = WMS_OUT_S_TX.ITM_CODE AND WMS_OUT_TX.PACK_KEY = WMS_OUT_S_TX.PACK_KEY " & _
                         " LEFT OUTER JOIN WMS_WAREHOUSE ON WMS_OUT_TX.IO_WH = WMS_WAREHOUSE.WH_CODE AND WMS_OUT_TX.IMP_CODE = WMS_WAREHOUSE.IMP_CODE " & _
                         " LEFT OUTER JOIN WMS_WH_BIN ON WMS_OUT_TX.IMP_CODE = WMS_WH_BIN.IMP_CODE AND WMS_OUT_TX.IO_LOC = WMS_WH_BIN.LOC_KEY " & _
                         " where IO_DOC='SI' " & whereSQL & appSQL & _
                         " UNION " & _
                         " SELECT IO_DATETIME, WMS_IN_TX.IO_TYPE,Convert(varchar,WMS_IN_TX.IO_DATETIME,103) as IO_DATETIME_D, 'IST' as IO_DOC, WMS_IN_TX.IO_DOC_ID, WMS_IN_TX.IO_BATCH_NO, WMS_ITEM.ITM_SKU_NO, " & _
                         " WMS_IN_TX.IO_QTY, Case when wms_item.itm_type='CABLE' then WMS_IN_S_TX.IOSX_UOM2 else WMS_ITEM.ITM_UOM end AS ITM_UOM, WMS_ITEM.ITM_PCS_PER_UOM, Case when wms_item.itm_type='CABLE' then WMS_IN_S_TX.IOSX_QTY2 else WMS_IN_TX.IO_QTY end as IN_QTY, null as OUT_QTY, Case when wms_item.itm_type='CABLE' then WMS_IN_S_TX.IOSX_QTY2 else WMS_IN_TX.IO_QTY end *isnull(WMS_ITEM.ITM_PCS_PER_UOM,1) as TOTAL_NUMBER,'' + WMS_STOCK_TRANSFER.TR_TQ_NO as ORDER_NO, 'RFT' as TRANS_TYPE, " & _
                         " WMS_IN_TX.IO_LOC, WMS_WAREHOUSE.WH_MAIN_WH,  WMS_IN_S_TX.IOSX_SERIAL_NO, WMS_IN_TX.ITM_CODE, WMS_IN_TX.PACK_KEY, " & _
                         " WMS_IN_TX.IO_WH, WMS_WH_BIN.BN_CSMS_CODE, WMS_IN_S_TX.IOSX_DRUM_ID, NULL as RT_TYPE, WMS_ITEM.ITM_TYPE " & _
                         " FROM WMS_IN_TX LEFT OUTER JOIN " & _
                         " WMS_ITEM ON WMS_IN_TX.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_IN_TX.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                         " WMS_IN_TX.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_IN_TX.PACK_KEY = WMS_ITEM.PACK_KEY INNER JOIN " & _
                         " WMS_STOCK_TRANSFER ON WMS_IN_TX.IMP_CODE = WMS_STOCK_TRANSFER.IMP_CODE AND WMS_IN_TX.STORER_CODE = WMS_STOCK_TRANSFER.STORER_CODE AND  " & _
                         " WMS_IN_TX.IO_DOC_ID = WMS_STOCK_TRANSFER.TR_CODE " & _
                         " LEFT OUTER JOIN WMS_IN_S_TX ON WMS_IN_TX.IMP_CODE = WMS_IN_S_TX.IMP_CODE AND WMS_IN_TX.STORER_CODE = WMS_IN_S_TX.STORER_CODE AND " & _
                         " WMS_IN_TX.IO_SYS_SEQ = WMS_IN_S_TX.IO_SYS_SEQ AND WMS_IN_S_TX.ITM_CODE = WMS_IN_TX.ITM_CODE AND WMS_IN_S_TX.PACK_KEY = WMS_IN_TX.PACK_KEY " & _
                         " LEFT OUTER JOIN WMS_WAREHOUSE ON WMS_IN_TX.IO_WH = WMS_WAREHOUSE.WH_CODE AND WMS_IN_TX.IMP_CODE = WMS_WAREHOUSE.IMP_CODE " & _
                         " LEFT OUTER JOIN WMS_WH_BIN ON WMS_IN_TX.IMP_CODE = WMS_WH_BIN.IMP_CODE AND WMS_IN_TX.IO_LOC = WMS_WH_BIN.LOC_KEY " & _
                         " WHERE IO_DOC='STF' AND SUBSTRING(IO_DOC_ID,1,2)='IT' " & whereSQL & appSQL & _
                         " UNION " & _
                         " SELECT IO_DATETIME, WMS_OUT_TX.IO_TYPE, Convert(varchar,WMS_OUT_TX.IO_DATETIME,103) as IO_DATETIME_D, 'IST' as IO_DOC, WMS_OUT_TX.IO_DOC_ID, WMS_OUT_TX.IO_BATCH_NO, WMS_ITEM.ITM_SKU_NO, " & _
                         " WMS_OUT_TX.IO_QTY, CASE WHEN wms_item.itm_type = 'CABLE' THEN WMS_OUT_S_TX.IOSX_UOM2 ELSE WMS_ITEM.ITM_UOM END, WMS_ITEM.ITM_PCS_PER_UOM, NULL as IN_QTY,CASE WHEN wms_item.itm_type = 'CABLE' THEN WMS_OUT_S_TX.IOSX_QTY2 ELSE WMS_OUT_TX.IO_QTY END as OUT_QTY, CASE WHEN wms_item.itm_type = 'CABLE' THEN WMS_OUT_S_TX.IOSX_QTY2 ELSE WMS_OUT_TX.IO_QTY END * isnull(WMS_ITEM.ITM_PCS_PER_UOM,1) * -1 as TOTAL_NUMBER,'' + WMS_STOCK_TRANSFER.TR_TQ_NO as ORDER_NO, 'RFT' as TRANS_TYPE, " & _
                         " WMS_OUT_TX.IO_LOC, WMS_WAREHOUSE.WH_MAIN_WH, WMS_OUT_S_TX.IOSX_SERIAL_NO, WMS_OUT_TX.ITM_CODE, WMS_OUT_TX.PACK_KEY, " & _
                         " WMS_OUT_TX.IO_WH, WMS_WH_BIN.BN_CSMS_CODE, WMS_OUT_S_TX.IOSX_DRUM_ID, NULL as RT_TYPE, WMS_ITEM.ITM_TYPE " & _
                         " FROM WMS_OUT_TX INNER JOIN " & _
                         " WMS_STOCK_TRANSFER ON WMS_OUT_TX.IMP_CODE = WMS_STOCK_TRANSFER.IMP_CODE AND WMS_OUT_TX.STORER_CODE = WMS_STOCK_TRANSFER.STORER_CODE AND  " & _
                         " WMS_OUT_TX.IO_DOC_ID = WMS_STOCK_TRANSFER.TR_CODE LEFT OUTER JOIN " & _
                         " WMS_ITEM ON WMS_OUT_TX.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_OUT_TX.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                         " WMS_OUT_TX.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_OUT_TX.PACK_KEY = WMS_ITEM.PACK_KEY " & _
                         " LEFT OUTER JOIN WMS_OUT_S_TX ON WMS_OUT_TX.IMP_CODE = WMS_OUT_S_TX.IMP_CODE AND WMS_OUT_TX.STORER_CODE = WMS_OUT_S_TX.STORER_CODE AND " & _
                         " WMS_OUT_TX.IO_SYS_SEQ = WMS_OUT_S_TX.IO_SYS_SEQ AND WMS_OUT_TX.ITM_CODE = WMS_OUT_S_TX.ITM_CODE AND WMS_OUT_TX.PACK_KEY = WMS_OUT_S_TX.PACK_KEY " & _
                         " LEFT OUTER JOIN WMS_WAREHOUSE ON WMS_OUT_TX.IO_WH = WMS_WAREHOUSE.WH_CODE AND WMS_OUT_TX.IMP_CODE = WMS_WAREHOUSE.IMP_CODE " & _
                         " LEFT OUTER JOIN WMS_WH_BIN ON WMS_OUT_TX.IMP_CODE = WMS_WH_BIN.IMP_CODE AND WMS_OUT_TX.IO_LOC = WMS_WH_BIN.LOC_KEY " & _
                         " where IO_DOC='STF' AND SUBSTRING(IO_DOC_ID,1,2)='IT' " & whereSQL & appSQL & _
                         " UNION " & _
                         " SELECT IO_DATETIME, WMS_IN_TX.IO_TYPE,Convert(varchar,WMS_IN_TX.IO_DATETIME,103) as IO_DATETIME_D, 'SRL' as IO_DOC, WMS_IN_TX.IO_DOC_ID, WMS_IN_TX.IO_BATCH_NO, WMS_ITEM.ITM_SKU_NO, " & _
                         " WMS_IN_TX.IO_QTY, Case when wms_item.itm_type='CABLE' then WMS_IN_S_TX.IOSX_UOM2 else WMS_ITEM.ITM_UOM end AS ITM_UOM, WMS_ITEM.ITM_PCS_PER_UOM, Case when wms_item.itm_type='CABLE' then WMS_IN_S_TX.IOSX_QTY2 else WMS_IN_TX.IO_QTY end as IN_QTY, null as OUT_QTY, Case when wms_item.itm_type='CABLE' then WMS_IN_S_TX.IOSX_QTY2 else WMS_IN_TX.IO_QTY end *isnull(WMS_ITEM.ITM_PCS_PER_UOM,1) as TOTAL_NUMBER,'' + WMS_STOCK_TRANSFER.TR_TQ_NO as ORDER_NO, '' as TRANS_TYPE, " & _
                         " WMS_IN_TX.IO_LOC, WMS_WAREHOUSE.WH_MAIN_WH,  WMS_IN_S_TX.IOSX_SERIAL_NO, WMS_IN_TX.ITM_CODE, WMS_IN_TX.PACK_KEY, " & _
                         " WMS_IN_TX.IO_WH, WMS_WH_BIN.BN_CSMS_CODE, WMS_IN_S_TX.IOSX_DRUM_ID, NULL as RT_TYPE, WMS_ITEM.ITM_TYPE " & _
                         " FROM WMS_IN_TX LEFT OUTER JOIN " & _
                         " WMS_ITEM ON WMS_IN_TX.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_IN_TX.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                         " WMS_IN_TX.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_IN_TX.PACK_KEY = WMS_ITEM.PACK_KEY INNER JOIN " & _
                         " WMS_STOCK_TRANSFER ON WMS_IN_TX.IMP_CODE = WMS_STOCK_TRANSFER.IMP_CODE AND WMS_IN_TX.STORER_CODE = WMS_STOCK_TRANSFER.STORER_CODE AND  " & _
                         " WMS_IN_TX.IO_DOC_ID = WMS_STOCK_TRANSFER.TR_CODE " & _
                         " LEFT OUTER JOIN WMS_IN_S_TX ON WMS_IN_TX.IMP_CODE = WMS_IN_S_TX.IMP_CODE AND WMS_IN_TX.STORER_CODE = WMS_IN_S_TX.STORER_CODE AND " & _
                         " WMS_IN_TX.IO_SYS_SEQ = WMS_IN_S_TX.IO_SYS_SEQ AND WMS_IN_S_TX.ITM_CODE = WMS_IN_TX.ITM_CODE AND WMS_IN_S_TX.PACK_KEY = WMS_IN_TX.PACK_KEY " & _
                         " LEFT OUTER JOIN WMS_WAREHOUSE ON WMS_IN_TX.IO_WH = WMS_WAREHOUSE.WH_CODE AND WMS_IN_TX.IMP_CODE = WMS_WAREHOUSE.IMP_CODE " & _
                         " LEFT OUTER JOIN WMS_WH_BIN ON WMS_IN_TX.IMP_CODE = WMS_WH_BIN.IMP_CODE AND WMS_IN_TX.IO_LOC = WMS_WH_BIN.LOC_KEY " & _
                         " WHERE IO_DOC='STF' AND SUBSTRING(IO_DOC_ID,1,2)='RL' " & whereSQL & appSQL & _
                         " UNION " & _
                         " SELECT IO_DATETIME, WMS_OUT_TX.IO_TYPE, Convert(varchar,WMS_OUT_TX.IO_DATETIME,103) as IO_DATETIME_D, 'SRL' as IO_DOC, WMS_OUT_TX.IO_DOC_ID, WMS_OUT_TX.IO_BATCH_NO, WMS_ITEM.ITM_SKU_NO, " & _
                         " WMS_OUT_TX.IO_QTY, CASE WHEN wms_item.itm_type = 'CABLE' THEN WMS_OUT_S_TX.IOSX_UOM2 ELSE WMS_ITEM.ITM_UOM END, WMS_ITEM.ITM_PCS_PER_UOM, NULL as IN_QTY,CASE WHEN wms_item.itm_type = 'CABLE' THEN WMS_OUT_S_TX.IOSX_QTY2 ELSE WMS_OUT_TX.IO_QTY END as OUT_QTY, CASE WHEN wms_item.itm_type = 'CABLE' THEN WMS_OUT_S_TX.IOSX_QTY2 ELSE WMS_OUT_TX.IO_QTY END * isnull(WMS_ITEM.ITM_PCS_PER_UOM,1) * -1 as TOTAL_NUMBER,'' + WMS_STOCK_TRANSFER.TR_TQ_NO as ORDER_NO, '' as TRANS_TYPE, " & _
                         " WMS_OUT_TX.IO_LOC, WMS_WAREHOUSE.WH_MAIN_WH, WMS_OUT_S_TX.IOSX_SERIAL_NO, WMS_OUT_TX.ITM_CODE, WMS_OUT_TX.PACK_KEY, " & _
                         " WMS_OUT_TX.IO_WH, WMS_WH_BIN.BN_CSMS_CODE, WMS_OUT_S_TX.IOSX_DRUM_ID, NULL as RT_TYPE, WMS_ITEM.ITM_TYPE " & _
                         " FROM WMS_OUT_TX INNER JOIN " & _
                         " WMS_STOCK_TRANSFER ON WMS_OUT_TX.IMP_CODE = WMS_STOCK_TRANSFER.IMP_CODE AND WMS_OUT_TX.STORER_CODE = WMS_STOCK_TRANSFER.STORER_CODE AND  " & _
                         " WMS_OUT_TX.IO_DOC_ID = WMS_STOCK_TRANSFER.TR_CODE LEFT OUTER JOIN " & _
                         " WMS_ITEM ON WMS_OUT_TX.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_OUT_TX.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                         " WMS_OUT_TX.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_OUT_TX.PACK_KEY = WMS_ITEM.PACK_KEY " & _
                         " LEFT OUTER JOIN WMS_OUT_S_TX ON WMS_OUT_TX.IMP_CODE = WMS_OUT_S_TX.IMP_CODE AND WMS_OUT_TX.STORER_CODE = WMS_OUT_S_TX.STORER_CODE AND " & _
                         " WMS_OUT_TX.IO_SYS_SEQ = WMS_OUT_S_TX.IO_SYS_SEQ AND WMS_OUT_TX.ITM_CODE = WMS_OUT_S_TX.ITM_CODE AND WMS_OUT_TX.PACK_KEY = WMS_OUT_S_TX.PACK_KEY " & _
                         " LEFT OUTER JOIN WMS_WAREHOUSE ON WMS_OUT_TX.IO_WH = WMS_WAREHOUSE.WH_CODE AND WMS_OUT_TX.IMP_CODE = WMS_WAREHOUSE.IMP_CODE " & _
                         " LEFT OUTER JOIN WMS_WH_BIN ON WMS_OUT_TX.IMP_CODE = WMS_WH_BIN.IMP_CODE AND WMS_OUT_TX.IO_LOC = WMS_WH_BIN.LOC_KEY " & _
                         " where IO_DOC='STF' AND SUBSTRING(IO_DOC_ID,1,2)='RL' " & whereSQL & appSQL & _
                         " UNION " & _
                         " SELECT IO_DATETIME, WMS_IN_TX.IO_TYPE,Convert(varchar,WMS_IN_TX.IO_DATETIME,103) as IO_DATETIME_D, IO_DOC, WMS_IN_TX.IO_DOC_ID, WMS_IN_TX.IO_BATCH_NO, WMS_ITEM.ITM_SKU_NO, " & _
                         " WMS_IN_TX.IO_QTY, Case when wms_item.itm_type='CABLE' then WMS_IN_S_TX.IOSX_UOM2 else WMS_ITEM.ITM_UOM end AS ITM_UOM, WMS_ITEM.ITM_PCS_PER_UOM, Case when wms_item.itm_type='CABLE' then WMS_IN_S_TX.IOSX_QTY2 else WMS_IN_TX.IO_QTY end as IN_QTY, null as OUT_QTY, Case when wms_item.itm_type='CABLE' then WMS_IN_S_TX.IOSX_QTY2 else WMS_IN_TX.IO_QTY end *isnull(WMS_ITEM.ITM_PCS_PER_UOM,1) as TOTAL_NUMBER,'' + WMS_STOCK_ADJUST.AD_CODE as ORDER_NO, 'Adjustment' as TRANS_TYPE, " & _
                         " WMS_IN_TX.IO_LOC, WMS_WAREHOUSE.WH_MAIN_WH,  WMS_IN_S_TX.IOSX_SERIAL_NO, WMS_IN_TX.ITM_CODE, WMS_IN_TX.PACK_KEY, " & _
                         " WMS_IN_TX.IO_WH, WMS_WH_BIN.BN_CSMS_CODE, WMS_IN_S_TX.IOSX_DRUM_ID, NULL as RT_TYPE, WMS_ITEM.ITM_TYPE " & _
                         " FROM WMS_IN_TX LEFT OUTER JOIN " & _
                         " WMS_ITEM ON WMS_IN_TX.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_IN_TX.STORER_CODE = WMS_ITEM.STORER_CODE AND   " & _
                         " WMS_IN_TX.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_IN_TX.PACK_KEY = WMS_ITEM.PACK_KEY INNER JOIN " & _
                         " WMS_STOCK_ADJUST ON WMS_IN_TX.IMP_CODE = WMS_STOCK_ADJUST.IMP_CODE AND WMS_IN_TX.STORER_CODE = WMS_STOCK_ADJUST.STORER_CODE AND   " & _
                         " WMS_IN_TX.IO_DOC_ID = WMS_STOCK_ADJUST.AD_CODE  " & _
                         " LEFT OUTER JOIN WMS_IN_S_TX ON WMS_IN_TX.IMP_CODE = WMS_IN_S_TX.IMP_CODE AND WMS_IN_TX.STORER_CODE = WMS_IN_S_TX.STORER_CODE AND  " & _
                         " WMS_IN_TX.IO_SYS_SEQ = WMS_IN_S_TX.IO_SYS_SEQ  " & _
                         " LEFT OUTER JOIN WMS_WAREHOUSE ON WMS_IN_TX.IO_WH = WMS_WAREHOUSE.WH_CODE AND WMS_IN_TX.IMP_CODE = WMS_WAREHOUSE.IMP_CODE " & _
                         " LEFT OUTER JOIN WMS_WH_BIN ON WMS_IN_TX.IMP_CODE = WMS_WH_BIN.IMP_CODE AND WMS_IN_TX.IO_LOC = WMS_WH_BIN.LOC_KEY " & _
                         " WHERE IO_DOC='SADJ' " & whereSQL & appSQL & _
                         " UNION  " & _
                         " SELECT IO_DATETIME, WMS_OUT_TX.IO_TYPE, Convert(varchar,WMS_OUT_TX.IO_DATETIME,103) as IO_DATETIME_D, IO_DOC, WMS_OUT_TX.IO_DOC_ID, WMS_OUT_TX.IO_BATCH_NO, WMS_ITEM.ITM_SKU_NO,  " & _
                         " WMS_OUT_TX.IO_QTY, CASE WHEN wms_item.itm_type = 'CABLE' THEN WMS_OUT_S_TX.IOSX_UOM2 ELSE WMS_ITEM.ITM_UOM END, WMS_ITEM.ITM_PCS_PER_UOM, NULL as IN_QTY,CASE WHEN wms_item.itm_type = 'CABLE' THEN WMS_OUT_S_TX.IOSX_QTY2 ELSE WMS_OUT_TX.IO_QTY END as OUT_QTY, CASE WHEN wms_item.itm_type = 'CABLE' THEN WMS_OUT_S_TX.IOSX_QTY2 ELSE WMS_OUT_TX.IO_QTY END * isnull(WMS_ITEM.ITM_PCS_PER_UOM,1) * -1 as TOTAL_NUMBER,'' + WMS_STOCK_ADJUST.AD_CODE as ORDER_NO, 'Adjustment' as TRANS_TYPE,  " & _
                         " WMS_OUT_TX.IO_LOC, WMS_WAREHOUSE.WH_MAIN_WH, WMS_OUT_S_TX.IOSX_SERIAL_NO, WMS_OUT_TX.ITM_CODE, WMS_OUT_TX.PACK_KEY, " & _
                         " WMS_OUT_TX.IO_WH, WMS_WH_BIN.BN_CSMS_CODE, WMS_OUT_S_TX.IOSX_DRUM_ID, NULL as RT_TYPE, WMS_ITEM.ITM_TYPE " & _
                         " FROM WMS_OUT_TX INNER JOIN  " & _
                         " WMS_STOCK_ADJUST ON WMS_OUT_TX.IMP_CODE = WMS_STOCK_ADJUST.IMP_CODE AND WMS_OUT_TX.STORER_CODE = WMS_STOCK_ADJUST.STORER_CODE AND   " & _
                         " WMS_OUT_TX.IO_DOC_ID = WMS_STOCK_ADJUST.AD_CODE LEFT OUTER JOIN  " & _
                         " WMS_ITEM ON WMS_OUT_TX.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_OUT_TX.STORER_CODE = WMS_ITEM.STORER_CODE AND   " & _
                         " WMS_OUT_TX.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_OUT_TX.PACK_KEY = WMS_ITEM.PACK_KEY  " & _
                         " LEFT OUTER JOIN WMS_OUT_S_TX ON WMS_OUT_TX.IMP_CODE = WMS_OUT_S_TX.IMP_CODE AND WMS_OUT_TX.STORER_CODE = WMS_OUT_S_TX.STORER_CODE AND  " & _
                         " WMS_OUT_TX.IO_SYS_SEQ = WMS_OUT_S_TX.IO_SYS_SEQ  " & _
                         " LEFT OUTER JOIN WMS_WAREHOUSE ON WMS_OUT_TX.IO_WH = WMS_WAREHOUSE.WH_CODE AND WMS_OUT_TX.IMP_CODE = WMS_WAREHOUSE.IMP_CODE " & _
                         " LEFT OUTER JOIN WMS_WH_BIN ON WMS_OUT_TX.IMP_CODE = WMS_WH_BIN.IMP_CODE AND WMS_OUT_TX.IO_LOC = WMS_WH_BIN.LOC_KEY " & _
                         " WHERE IO_DOC='SADJ' " & whereSQL & appSQL & _
                         " ORDER BY IO_DATETIME, ITM_SKU_NO, IO_BATCH_NO"


            'Dim tempSQL As String = gDB.getCmdSql(sql_string, pa1)
            result_dt = gDB.getDataTable(sql_string)

            If result_dt.Rows.Count > 0 Then

                ExcelWBObj = New XSSFWorkbook(ReadFileStream)
                'TemplateSheet = ExcelWBObj.GetSheet("TEMPLATE")
                ReportSheet = ExcelWBObj.GetSheetAt(0)

                'XXXXXXXXXXXXXXXXXX Excel Cell StyleXXXXXXXXXXXXXXXXXXXXX
                'Dim warpStyle As HSSFCellStyle
                'warpStyle = ExcelWBObj.CreateCellStyle()
                'warpStyle.WrapText = True

                'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


                'XXXXXXXXXXX Set Param Sheet XXXXXXXXXXXXXXXXXXXXXX


                'nf.setCellValue(ParamSheet, 13, 3, gU.decodeNullOrEmpty(FR_DO_SCH_DATE, Now.Date.ToString("dd/MM/yyyy")))
                'nf.setCellValue(ParamSheet, 13, 8, gU.decodeNullOrEmpty(TO_DO_SCH_DATE, Now.Date.ToString("dd/MM/yyyy")))

                'XXXXXXXXXXXXX TITLE XXXXXXXXXXXXXXXXXXXXXXX

                'nf.CopyRange(TemplateSheet, ReportSheet, New CellRangeAddress(0, 0, 0, 17), 0, 0)
                'For i = 0 To 17
                '    ReportSheet.SetColumnWidth(i, TemplateSheet.GetColumnWidth(i))
                'Next


                ''XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

                'XXXXXXXXXXXXXXXXXX Column Header XXXXXXXXXXXXXXXXXXXXXXXXXXXXX

                'REM XXXXXXXXXXXXXXXXXXX CONTENT BLOCK XXXXXXXXXXXXXXXXXXX


                ExcelRow_num = 1

                For i = 0 To result_dt.Rows.Count - 1

                    'If PrevSKU <> result_dt.Rows(i).Item("itm_sku_no").ToString.Trim Then

                    'NewSheet = ExcelWBObj.CloneSheet(0)
                    'sheetIDX = ExcelWBObj.GetSheetIndex(NewSheet)
                    '    ExcelWBObj.SetSheetName(sheetIDX, "Sort " & result_dt.Rows(i).Item("itm_sku_no").ToString.Trim)
                    'ReportSheet = NewSheet

                    '    PrevSKU = result_dt.Rows(i).Item("itm_sku_no").ToString.Trim
                    'End If

                    'nf.CopyRange(TemplateSheet, ReportSheet, New CellRangeAddress(4, 4, 0, 1), ExcelRow_num, 0)
                    nf.setCellValue(ReportSheet, ExcelRow_num, 0, result_dt.Rows(i).Item("IO_DATETIME").ToString.Trim)
                    nf.setCellValue(ReportSheet, ExcelRow_num, 1, result_dt.Rows(i).Item("WH_MAIN_WH").ToString.Trim)
                    nf.setCellValue(ReportSheet, ExcelRow_num, 2, result_dt.Rows(i).Item("IO_WH").ToString.Trim)
                    nf.setCellValue(ReportSheet, ExcelRow_num, 3, result_dt.Rows(i).Item("BN_CSMS_CODE").ToString.Trim)
                    nf.setCellValue(ReportSheet, ExcelRow_num, 4, result_dt.Rows(i).Item("ITM_SKU_NO").ToString.Trim)
                    nf.setCellValue(ReportSheet, ExcelRow_num, 5, result_dt.Rows(i).Item("IOSX_DRUM_ID").ToString.Trim)
                    nf.setCellValue(ReportSheet, ExcelRow_num, 6, result_dt.Rows(i).Item("IOSX_SERIAL_NO").ToString.Trim)
                    nf.setCellValue(ReportSheet, ExcelRow_num, 7, result_dt.Rows(i).Item("IO_BATCH_NO").ToString.Trim)
                    nf.setCellValue(ReportSheet, ExcelRow_num, 8, returnDOCType(result_dt.Rows(i).Item("IO_DOC").ToString.Trim))
                    nf.setCellValue(ReportSheet, ExcelRow_num, 9, result_dt.Rows(i).Item("TRANS_TYPE").ToString.Trim)
                    nf.setCellValue(ReportSheet, ExcelRow_num, 10, result_dt.Rows(i).Item("ORDER_NO").ToString.Trim)
                    nf.setCellValue(ReportSheet, ExcelRow_num, 11, result_dt.Rows(i).Item("RT_TYPE").ToString.Trim)
                    nf.setCellValue(ReportSheet, ExcelRow_num, 12, result_dt.Rows(i).Item("IN_QTY").ToString.Trim, "DEC")
                    nf.setCellValue(ReportSheet, ExcelRow_num, 13, result_dt.Rows(i).Item("OUT_QTY").ToString.Trim, "DEC")
                    nf.setCellValue(ReportSheet, ExcelRow_num, 14, result_dt.Rows(i).Item("ITM_UOM").ToString.Trim)
                    nf.setCellValue(ReportSheet, ExcelRow_num, 15, result_dt.Rows(i).Item("TOTAL_NUMBER").ToString.Trim, "DEC")
                    nf.setCellValue(ReportSheet, ExcelRow_num, 16, result_dt.Rows(i).Item("ITM_TYPE").ToString.Trim)

                    ExcelRow_num += 1

                Next


                'ExcelWBObj.RemoveSheetAt(ExcelWBObj.GetSheetIndex("TEMPLATE"))

                sysFileName = "RPT_STK_MOVE" & Now.ToString("yyyyMMddHHmmssfff")
                sysFileName = sysFileName & "." & LCase("XLSX")
                sysFilePath = fileFolder & "\" & sysFileName


                tmpFileStream = New FileStream(sysFilePath, FileMode.Create)
                ExcelWBObj.Write(tmpFileStream)

                tmpFileStream.Close()
                tmpFileStream.Dispose()

                Response.Write("<script language=""JavaScript"">dsp_status.innerHTML = ""DONE"";</script>")
                Response.Write("<script language=""JavaScript"">document.forms[0].hdf_file_path.value = """ & gU.jsString(sysFilePath) & """;</script>")
                Response.Write("<script language=""JavaScript"">document.getElementById(""tr_download"").style.display = """";</script>")
                Response.Flush()

            Else

                Response.Write("<script language=""JavaScript"">dsp_status.innerHTML = ""No Record Found."";</script>")

            End If

            REM XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

        Catch ex As Exception
            Throw ex
        Finally
            'tmpStreamWriter.Close()
            'tmpStreamWriter.Dispose()
            'tmpStreamWriter = Nothing

            ReadFileStream.Close()
            ReadFileStream.Dispose()
            tmpFileStream = Nothing
            ReadFileStream = Nothing
        End Try



    End Sub


    Protected Sub ibtn_download_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ibtn_download.Click
        Dim fullFilePath As String = hdf_file_path.Value
        Dim userFileName As String = "STOCK_MOVEMENT_RPT" & Now.ToString("yyyyMMdd")

        If fullFilePath <> "" Then
            Dim nFile As System.IO.FileInfo = New System.IO.FileInfo(fullFilePath)

            If nFile.Exists Then
                Response.Clear()
                Response.AddHeader("Content-Disposition", "attachment; filename=" & userFileName & nFile.Extension)
                Response.AddHeader("Content-Length", nFile.Length.ToString())
                Response.ContentType = "application/octet-stream"
                Response.WriteFile(nFile.FullName)
                Response.End()
            Else
                Response.Write("This file does not exist.")
            End If
        End If
    End Sub

    Private Function returnDOCType(ByVal io_doc As String) As String
        Dim returnType As String

        Select Case io_doc
            Case "GR"
                returnType = "Goods Receiving"
            Case "DO"
                returnType = "Goods Issuing"
            Case "SI"
                returnType = "Stock Issue"
            Case "SR"
                returnType = "Stock Return"
            Case "SRL"
                returnType = "Stock Relocation"
            Case "IST"
                returnType = "Inter-Store Transfer"
            Case "SADJ"
                returnType = "Stock Adjust"
            Case Else
                returnType = io_doc
        End Select

        Return returnType
    End Function
End Class
