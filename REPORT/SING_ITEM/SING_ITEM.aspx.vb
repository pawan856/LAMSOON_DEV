Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports Microsoft.Reporting.WebForms

Partial Class RPT_SING_ITEM
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private wmsFun As New WMSFunc

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'ar = New AccessRightUtils("OB_DO", Session("usr_id"), Me)

        'ar.hideForm(Me)

        Dim sqlString As String
        Dim nDataSource As New DataTable
        Dim itm_code As String = ""
        Dim imp_code As String = ""
        Dim temp_where1 As String = ""
        Dim temp_where2 As String = ""
        Dim temp_total_where As String = ""

        Dim Total_crt As Integer = 0

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")

        Dim sku_no As String = ""
        Dim pack_key As String = ""
        Dim STORER_CODE As String = ""

        sku_no = Session("SEARCH_SESSION_PAGE_ITM_SKU_NO")
        itm_code = Session("SEARCH_SESSION_PAGE_ITM_CODE")
        pack_key = Session("SEARCH_SESSION_PAGE_PACK_KEY")
        STORER_CODE = Session("SEARCH_SESSION_PAGE_STORER_CODE")

        If Session("SEARCH_SESSION_PAGE_ITM_SKU_NO") = "" Then
            If Session("SEARCH_SESSION_PAGE_ITM_CODE") = "" OrElse Session("SEARCH_SESSION_PAGE_PACK_KEY") = "" Then
                Response.Write("Please Enter Stock No or Item Code and Pack Key for generating Report!")
                Response.End()
            End If
        End If


        Dim tempStr As String = ""

        If sku_no <> "" Then tempStr &= " AND upper(ITM_SKU_NO)=upper('" & gU.dbEncode(sku_no) & "') "
        If STORER_CODE <> "" Then tempStr &= " AND STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' "
        If itm_code <> "" Then tempStr &= " AND itm_code='" & gU.dbEncode(itm_code) & "' "
        If pack_key <> "" Then tempStr &= " AND PACK_KEY='" & gU.dbEncode(pack_key) & "' "

        sqlString = "Select count(*) from wms_item Where imp_code='" & gU.dbEncode(imp_code) & "'" & tempStr

        Dim itmCount As Integer = 0
        itmCount = gU.decodeEmptyCInt(DB.getValueFromSQL(sqlString), 0)

        If itmCount = 0 Then
            Response.Write("No item has been found!")
            Response.End()
        ElseIf itmCount > 1 Then
            Response.Write("More than one item has been found! Please Use Item Code and pack key for exact item match.")
            Response.End()
        End If


        If Session("SEARCH_SESSION_PAGE_ITM_SKU_NO") <> "" Then
            temp_where1 &= " AND Upper(WMS_item.itm_sku_no)=upper('" & gU.dbEncode(Session("SEARCH_SESSION_PAGE_ITM_SKU_NO")) & "') "
            temp_where2 &= " AND upper(WMS_item.itm_sku_no)=upper('" & gU.dbEncode(Session("SEARCH_SESSION_PAGE_ITM_SKU_NO")) & "') "
        End If

        If itm_code <> "" Then
            temp_where1 &= " AND WMS_IN_TX.itm_code='" & gU.dbEncode(itm_code) & "' "
            temp_where2 &= " AND WMS_OUT_TX.itm_code='" & gU.dbEncode(itm_code) & "' "
        End If

        If pack_key <> "" Then
            temp_where1 &= " AND WMS_IN_TX.PACK_KEY='" & gU.dbEncode(pack_key) & "' "
            temp_where2 &= " AND WMS_OUT_TX.PACK_KEY='" & gU.dbEncode(pack_key) & "' "
        End If


        If Session("SEARCH_SESSION_PAGE_STORER_CODE") <> "" Then
            temp_where1 &= " AND WMS_IN_TX.STORER_CODE='" & gU.dbEncode(Session("SEARCH_SESSION_PAGE_STORER_CODE")) & "' "
            temp_where2 &= " AND WMS_OUT_TX.STORER_CODE='" & gU.dbEncode(Session("SEARCH_SESSION_PAGE_STORER_CODE")) & "' "
            temp_total_where &= " AND WMS_ITEM.STORER_CODE='" & gU.dbEncode(Session("SEARCH_SESSION_PAGE_STORER_CODE")) & "' "
        End If

        If Session("SEARCH_SESSION_PAGE_FR_IO_DATETIME") <> "" Then

            temp_where1 &= " AND WMS_IN_TX.IO_DATETIME>= CONVERT(DATE,'" & Session("SEARCH_SESSION_PAGE_FR_IO_DATETIME") & "'," & gU.getConfig("DDFORMATNO") & ")"
            temp_where2 &= " AND WMS_OUT_TX.IO_DATETIME>= CONVERT(DATE,'" & Session("SEARCH_SESSION_PAGE_FR_IO_DATETIME") & "'," & gU.getConfig("DDFORMATNO") & ")"

            If Session("SEARCH_SESSION_PAGE_TO_IO_DATETIME") <> "" Then

                temp_where1 &= " AND WMS_IN_TX.IO_DATETIME < dateadd(d,1,CONVERT(DATE,'" & Session("SEARCH_SESSION_PAGE_TO_IO_DATETIME") & "'," & gU.getConfig("DDFORMATNO") & ")) "
                temp_where2 &= " AND WMS_OUT_TX.IO_DATETIME < dateadd(d,1,CONVERT(DATE,'" & Session("SEARCH_SESSION_PAGE_TO_IO_DATETIME") & "'," & gU.getConfig("DDFORMATNO") & ")) "

            End If
        End If


        'imp_code = "WMS"
        'do_code = "000001"
        'storer_code = "001"

        'sqlString = "(SELECT     WMS_STORER.STO_SHORTNAME, WMS_GOODSRCV.GR_REF_NO AS ERP_NO, WMS_IN_TX.ITM_CODE, '' as INV_NO, " & _
        '            "            Convert(varchar, WMS_IN_TX.IO_DATETIME,103) as IO_DATETIME, SUM(WMS_IN_TX.IO_QTY * WMS_IN_TX.PACK_KEY) AS IN_IO_QTY,null as OUT_IO_QTY, WMS_IN_TX.IO_BATCH_NO, SUM(WMS_IN_TX.IO_QTY * WMS_IN_TX.PACK_KEY) AS CURR_QTY " & _
        '            " FROM       WMS_IN_TX Left outer Join " & _
        '            "            WMS_GOODSRCV ON WMS_IN_TX.IMP_CODE = WMS_GOODSRCV.IMP_CODE AND " & _
        '            "            WMS_IN_TX.STORER_CODE = WMS_GOODSRCV.STORER_CODE AND WMS_IN_TX.IO_DOC_ID = WMS_GOODSRCV.GR_CODE INNER JOIN " & _
        '            "            WMS_STORER ON WMS_IN_TX.IMP_CODE = WMS_STORER.IMP_CODE AND WMS_IN_TX.STORER_CODE = WMS_STORER.STORER_CODE " & _
        '            " where WMS_IN_TX.ITM_CODE='" & gU.dbEncode(itm_code) & "' AND WMS_IN_TX.IMP_CODE='" & gU.dbEncode(imp_code) & "'" & temp_where1 & _
        '            " GROUP BY WMS_STORER.STO_SHORTNAME, WMS_GOODSRCV.GR_REF_NO,  Convert(varchar, WMS_IN_TX.IO_DATETIME,103), " & _
        '            "            WMS_IN_TX.ITM_CODE, WMS_IN_TX.IO_BATCH_NO) " & _
        '            "            union " & _
        '            " (SELECT    WMS_STORER.STO_SHORTNAME, '' AS ERP_NO, WMS_OUT_TX.ITM_CODE,  WMS_GOODSRCV.GR_INV_NO AS INV_NO, " & _
        '            "            Convert(varchar, WMS_OUT_TX.IO_DATETIME,103) as IO_DATETIME,null as IN_IO_QTY, SUM(WMS_OUT_TX.IO_QTY * WMS_OUT_TX.PACK_KEY) AS OUT_IO_QTY, WMS_OUT_TX.IO_BATCH_NO, 0 - SUM(WMS_OUT_TX.IO_QTY * WMS_OUT_TX.PACK_KEY) AS CURR_QTY " & _
        '            " FROM       WMS_GOODSRCV INNER JOIN " & _
        '            "            WMS_OUT_TX ON WMS_GOODSRCV.IMP_CODE = WMS_OUT_TX.IMP_CODE AND " & _
        '            "            WMS_GOODSRCV.STORER_CODE = WMS_OUT_TX.STORER_CODE AND WMS_GOODSRCV.GR_CODE = WMS_OUT_TX.IO_DOC_ID INNER JOIN " & _
        '            "            WMS_STORER ON WMS_OUT_TX.IMP_CODE = WMS_STORER.IMP_CODE AND " & _
        '            "            WMS_OUT_TX.STORER_CODE = WMS_STORER.STORER_CODE " & _
        '            " where WMS_OUT_TX.IO_DOC='GR' AND WMS_OUT_TX.ITM_CODE='" & gU.dbEncode(itm_code) & "' AND WMS_OUT_TX.IMP_CODE='" & gU.dbEncode(imp_code) & "'" & temp_where2 & _
        '            " GROUP BY WMS_STORER.STO_SHORTNAME, WMS_GOODSRCV.GR_INV_NO, Convert(varchar, WMS_OUT_TX.IO_DATETIME,103), " & _
        '            "            WMS_OUT_TX.ITM_CODE, WMS_OUT_TX.IO_BATCH_NO)union " & _
        '            " (SELECT    WMS_STORER.STO_SHORTNAME, '' as ERP_NO,WMS_OUT_TX.ITM_CODE,WMS_DELV_ORDER.DO_INV_NO as INV_NO, Convert(varchar, WMS_OUT_TX.IO_DATETIME,103) AS IO_DATETIME, NULL " & _
        '            "            AS IN_IO_QTY, SUM(WMS_OUT_TX.IO_QTY * WMS_OUT_TX.PACK_KEY) AS OUT_IO_QTY, WMS_OUT_TX.IO_BATCH_NO,0 - SUM(WMS_OUT_TX.IO_QTY * WMS_OUT_TX.PACK_KEY) AS CURR_QTY " & _
        '            " FROM       WMS_OUT_TX INNER JOIN " & _
        '            "            WMS_STORER ON WMS_OUT_TX.IMP_CODE = WMS_STORER.IMP_CODE AND " & _
        '            "            WMS_OUT_TX.STORER_CODE = WMS_STORER.STORER_CODE LEFT outer join " & _
        '            "            WMS_DELV_ORDER ON WMS_OUT_TX.IMP_CODE = WMS_DELV_ORDER.IMP_CODE AND " & _
        '            "            WMS_OUT_TX.STORER_CODE = WMS_DELV_ORDER.STORER_CODE AND WMS_OUT_TX.IO_DOC_ID = WMS_DELV_ORDER.DO_CODE " & _
        '            " where WMS_OUT_TX.IO_DOC='DO' AND WMS_OUT_TX.ITM_CODE='" & gU.dbEncode(itm_code) & "' AND WMS_OUT_TX.IMP_CODE='" & gU.dbEncode(imp_code) & "'" & temp_where2 & _
        '            " GROUP BY WMS_STORER.STO_SHORTNAME, Convert(varchar, WMS_OUT_TX.IO_DATETIME,103), WMS_OUT_TX.ITM_CODE, WMS_OUT_TX.IO_BATCH_NO, " & _
        '            "                       WMS_DELV_ORDER.DO_INV_NO) " & _
        '            " order by IO_DATETIME"
        If Not IsPostBack Then


            'sqlString = "(SELECT WMS_STORER.STO_SHORTNAME,  WMS_GOODSRCV.GR_REF_NO AS ERP_NO,  WMS_IN_TX.ITM_CODE, WMS_IN_TX.IO_DOC, WMS_IN_TX.IO_DOC_ID, " & _
            '            "  ''                                         AS INV_NO, Convert(varchar, WMS_IN_TX.IO_DATETIME,103)    AS IO_DATETIME, " & _
            '            "  SUM(WMS_IN_TX.IO_QTY) AS IN_IO_QTY,   NULL AS OUT_IO_QTY, '' as UNPOST_YN, " & _
            '            "  WMS_IN_TX.IO_BATCH_NO,   SUM(WMS_IN_TX.IO_QTY * ISNULL(case when aitm_pcs_per_pack=0 THEN TO_NUMBER(WMS_IN_TX.PACK_KEY) else aitm_pcs_per_pack end,WMS_IN_TX.PACK_KEY)) AS CURR_QTY,'2' as SORT_COL " & _
            '            " FROM WMS_IN_TX " & _
            '            " LEFT OUTER JOIN WMS_GOODSRCV " & _
            '            " ON WMS_IN_TX.IMP_CODE     = WMS_GOODSRCV.IMP_CODE  AND WMS_IN_TX.STORER_CODE = WMS_GOODSRCV.STORER_CODE  AND WMS_IN_TX.IO_DOC_ID   = WMS_GOODSRCV.GR_CODE " & _
            '            " INNER JOIN WMS_STORER " & _
            '            " ON WMS_IN_TX.IMP_CODE     = WMS_STORER.IMP_CODE  AND WMS_IN_TX.STORER_CODE = WMS_STORER.STORER_CODE " & _
            '            " LEFT OUTER JOIN " & _
            '            " (SELECT     MAX(AITM_PCS_PER_PACK) AS aitm_pcs_per_pack, IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY " & _
            '            " FROM          WMS_ALT_VEND_ITEM d GROUP BY IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY) t ON WMS_IN_TX.IMP_CODE = t.IMP_CODE AND " & _
            '            " WMS_IN_TX.STORER_CODE = t.STORER_CODE And WMS_IN_TX.ITM_CODE = t.ITM_CODE And WMS_IN_TX.PACK_KEY = t.PACK_KEY " & _
            '            " WHERE   WMS_IN_TX.IO_TYPE='IN' AND WMS_IN_TX.ITM_CODE='" & gU.dbEncode(itm_code) & "' AND WMS_IN_TX.IMP_CODE='" & gU.dbEncode(imp_code) & "'" & temp_where1 & _
            '            " AND WMS_IN_TX.PACK_KEY='" & gU.dbEncode(pack_key) & "' " & _
            '            " GROUP BY WMS_STORER.STO_SHORTNAME,  WMS_GOODSRCV.GR_REF_NO,  Convert(varchar, WMS_IN_TX.IO_DATETIME,103),  WMS_IN_TX.ITM_CODE,  WMS_IN_TX.IO_DOC, WMS_IN_TX.IO_DOC_ID, WMS_IN_TX.IO_BATCH_NO " & _
            '            " ) " & _
            '            " UNION " & _
            '            " (SELECT WMS_STORER.STO_SHORTNAME,  '' AS ERP_NO,  WMS_OUT_TX.ITM_CODE, WMS_OUT_TX.IO_DOC, WMS_OUT_TX.IO_DOC_ID, WMS_DELV_ORDER.DO_INV_NO AS INV_NO,  " & _
            '            "    Convert(varchar, WMS_OUT_TX.IO_DATETIME,103)   AS IO_DATETIME,   NULL AS IN_IO_QTY, SUM(WMS_OUT_TX.IO_QTY) AS OUT_IO_QTY, '' as UNPOST_YN, " & _
            '            "    WMS_OUT_TX.IO_BATCH_NO,  0 - SUM(WMS_OUT_TX.IO_QTY * ISNULL(case when aitm_pcs_per_pack=0 THEN TO_NUMBER(WMS_OUT_TX.PACK_KEY) else aitm_pcs_per_pack end,WMS_OUT_TX.PACK_KEY)) AS CURR_QTY,'2' as SORT_COL " & _
            '            "  FROM WMS_OUT_TX " & _
            '            "  INNER JOIN WMS_STORER " & _
            '            "  ON WMS_OUT_TX.IMP_CODE     = WMS_STORER.IMP_CODE  AND WMS_OUT_TX.STORER_CODE = WMS_STORER.STORER_CODE " & _
            '            "  LEFT OUTER JOIN WMS_DELV_ORDER " & _
            '            "  ON WMS_OUT_TX.IMP_CODE     = WMS_DELV_ORDER.IMP_CODE  AND WMS_OUT_TX.STORER_CODE = WMS_DELV_ORDER.STORER_CODE AND WMS_OUT_TX.IO_DOC_ID   = WMS_DELV_ORDER.DO_CODE " & _
            '            " LEFT OUTER JOIN (SELECT     MAX(AITM_PCS_PER_PACK) AS aitm_pcs_per_pack, IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY " & _
            '            " FROM          WMS_ALT_VEND_ITEM d GROUP BY IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY) t ON WMS_OUT_TX.IMP_CODE = t.IMP_CODE AND " & _
            '            " WMS_OUT_TX.STORER_CODE = t.STORER_CODE And WMS_OUT_TX.ITM_CODE = t.ITM_CODE And WMS_OUT_TX.PACK_KEY = t.PACK_KEY " & _
            '            "  WHERE WMS_OUT_TX.IO_TYPE    ='OUT' AND WMS_OUT_TX.ITM_CODE='" & gU.dbEncode(itm_code) & "' AND WMS_OUT_TX.IMP_CODE='" & gU.dbEncode(imp_code) & "'" & temp_where2 & _
            '            " AND WMS_OUT_TX.PACK_KEY='" & gU.dbEncode(pack_key) & "' " & _
            '            "  GROUP BY WMS_STORER.STO_SHORTNAME,  Convert(varchar, WMS_OUT_TX.IO_DATETIME,103),  WMS_OUT_TX.ITM_CODE,  WMS_OUT_TX.IO_BATCH_NO,  WMS_OUT_TX.IO_DOC, WMS_OUT_TX.IO_DOC_ID, WMS_DELV_ORDER.DO_INV_NO " & _
            '            "  ) " & _
            '            "  union " & _
            '            "  (SELECT WMS_STORER.STO_SHORTNAME,  '' AS ERP_NO,  WMS_IN_TX.ITM_CODE, WMS_IN_TX.IO_DOC, WMS_IN_TX.IO_DOC_ID, WMS_DELV_ORDER.DO_INV_NO AS INV_NO, Convert(varchar, WMS_IN_TX.IO_DATETIME,103)    AS IO_DATETIME, " & _
            '            "  SUM(WMS_IN_TX.IO_QTY) AS IN_IO_QTY,  NULL AS OUT_IO_QTY, 'UNPOST' as UNPOST_YN, WMS_IN_TX.IO_BATCH_NO,  SUM(WMS_IN_TX.IO_QTY * ISNULL(case when aitm_pcs_per_pack=0 THEN TO_NUMBER(WMS_IN_TX.PACK_KEY) else aitm_pcs_per_pack end,WMS_IN_TX.PACK_KEY)) AS CURR_QTY,'1' as SORT_COL " & _
            '            " FROM WMS_IN_TX " & _
            '            " LEFT OUTER JOIN WMS_DELV_ORDER " & _
            '            " ON WMS_IN_TX.IMP_CODE     = WMS_DELV_ORDER.IMP_CODE  AND WMS_IN_TX.STORER_CODE = WMS_DELV_ORDER.STORER_CODE  AND WMS_IN_TX.IO_DOC_ID   = WMS_DELV_ORDER.DO_CODE " & _
            '            " INNER JOIN WMS_STORER " & _
            '            " ON WMS_IN_TX.IMP_CODE     = WMS_STORER.IMP_CODE  AND WMS_IN_TX.STORER_CODE = WMS_STORER.STORER_CODE " & _
            '            " LEFT OUTER JOIN " & _
            '            " (SELECT     MAX(AITM_PCS_PER_PACK) AS aitm_pcs_per_pack, IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY " & _
            '            " FROM          WMS_ALT_VEND_ITEM d GROUP BY IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY) t ON WMS_IN_TX.IMP_CODE = t.IMP_CODE AND " & _
            '            " WMS_IN_TX.STORER_CODE = t.STORER_CODE And WMS_IN_TX.ITM_CODE = t.ITM_CODE And WMS_IN_TX.PACK_KEY = t.PACK_KEY " & _
            '            " WHERE WMS_IN_TX.IO_TYPE='UNPOST' AND WMS_IN_TX.ITM_CODE='" & gU.dbEncode(itm_code) & "' AND WMS_IN_TX.IMP_CODE='" & gU.dbEncode(imp_code) & "'" & temp_where1 & _
            '            " AND WMS_IN_TX.PACK_KEY='" & gU.dbEncode(pack_key) & "' " & _
            '            " GROUP BY WMS_STORER.STO_SHORTNAME,  WMS_DELV_ORDER.DO_INV_NO,  Convert(varchar, WMS_IN_TX.IO_DATETIME,103),  WMS_IN_TX.ITM_CODE, WMS_IN_TX.IO_DOC, WMS_IN_TX.IO_DOC_ID, WMS_IN_TX.IO_BATCH_NO " & _
            '            " ) " & _
            '            " UNION " & _
            '            " (SELECT WMS_STORER.STO_SHORTNAME,  WMS_GOODSRCV.GR_REF_NO AS ERP_NO,  WMS_OUT_TX.ITM_CODE, WMS_OUT_TX.IO_DOC, WMS_OUT_TX.IO_DOC_ID, '' AS INV_NO,   Convert(varchar, WMS_OUT_TX.IO_DATETIME,103)   AS IO_DATETIME, " & _
            '            "    NULL AS IN_IO_QTY,  SUM(WMS_OUT_TX.IO_QTY) AS OUT_IO_QTY, 'UNPOST' as UNPOST_YN, WMS_OUT_TX.IO_BATCH_NO, " & _
            '            "    0 - SUM(WMS_OUT_TX.IO_QTY * ISNULL(case when aitm_pcs_per_pack=0 THEN TO_NUMBER(WMS_OUT_TX.PACK_KEY) else aitm_pcs_per_pack end,WMS_OUT_TX.PACK_KEY)) AS CURR_QTY,'1' as SORT_COL " & _
            '            "  FROM WMS_OUT_TX " & _
            '            "  INNER JOIN WMS_STORER " & _
            '            "  ON WMS_OUT_TX.IMP_CODE     = WMS_STORER.IMP_CODE  AND WMS_OUT_TX.STORER_CODE = WMS_STORER.STORER_CODE " & _
            '            "  LEFT OUTER JOIN WMS_GOODSRCV " & _
            '            "  ON WMS_OUT_TX.IMP_CODE     = WMS_GOODSRCV.IMP_CODE  AND WMS_OUT_TX.STORER_CODE = WMS_GOODSRCV.STORER_CODE AND WMS_OUT_TX.IO_DOC_ID   = WMS_GOODSRCV.GR_CODE " & _
            '            " LEFT OUTER JOIN (SELECT     MAX(AITM_PCS_PER_PACK) AS aitm_pcs_per_pack, IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY " & _
            '            " FROM          WMS_ALT_VEND_ITEM d GROUP BY IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY) t ON WMS_OUT_TX.IMP_CODE = t.IMP_CODE AND " & _
            '            " WMS_OUT_TX.STORER_CODE = t.STORER_CODE And WMS_OUT_TX.ITM_CODE = t.ITM_CODE And WMS_OUT_TX.PACK_KEY = t.PACK_KEY " & _
            '            " WHERE WMS_OUT_TX.IO_TYPE    ='UNPOST' AND WMS_OUT_TX.ITM_CODE='" & gU.dbEncode(itm_code) & "' AND WMS_OUT_TX.IMP_CODE='" & gU.dbEncode(imp_code) & "'" & temp_where2 & _
            '            " AND WMS_OUT_TX.PACK_KEY='" & gU.dbEncode(pack_key) & "' " & _
            '            " GROUP BY WMS_STORER.STO_SHORTNAME, Convert(varchar, WMS_OUT_TX.IO_DATETIME,103),  WMS_OUT_TX.ITM_CODE,  WMS_OUT_TX.IO_BATCH_NO,  WMS_OUT_TX.IO_DOC, WMS_OUT_TX.IO_DOC_ID, WMS_GOODSRCV.GR_REF_NO " & _
            '            " ) " & _
            '            " ORDER BY IO_DATETIME,SORT_COL "

            sqlString = " (SELECT WMS_STORER.STO_SHORTNAME, " & _
                        "   WMS_GOODSRCV.GR_REF_NO AS ERP_NO, " & _
                        "   WMS_IN_TX.ITM_CODE, " & _
                        "   WMS_IN_TX.PACK_KEY, " & _
                        "   WMS_ITEM.ITM_SKU_NO, " & _
                        "   WMS_IN_TX.IO_DOC, " & _
                        "   WMS_IN_TX.IO_DOC_ID, " & _
                        "   ''                                          AS INV_NO, " & _
                        "   Convert(varchar, WMS_IN_TX.IO_DATETIME,103) AS IO_DATETIME, " & _
                        "   WMS_IN_TX.IO_DATETIME AS IO_DATETIME_ORD, " & _
                        "   SUM(WMS_IN_TX.IO_QTY)                       AS IN_IO_QTY, " & _
                        "   NULL                                        AS OUT_IO_QTY, " & _
                        "   ''                                          AS UNPOST_YN, " & _
                        "   WMS_IN_TX.IO_BATCH_NO, " & _
                        "   SUM(WMS_IN_TX.IO_QTY * ISNULL(wms_item.ITM_PCS_PER_UOM,1)) AS CURR_QTY, " & _
                        "   wms_item.itm_uom, ISNULL(wms_item.ITM_PCS_PER_UOM,1) as ITM_PCS_PER_UOM , '' as cus_name, " & _
                        "   '2'                      AS SORT_COL " & _
                        " FROM WMS_IN_TX " & _
                        " LEFT OUTER JOIN WMS_GOODSRCV " & _
                        " ON WMS_IN_TX.IMP_CODE     = WMS_GOODSRCV.IMP_CODE " & _
                        " AND WMS_IN_TX.STORER_CODE = WMS_GOODSRCV.STORER_CODE " & _
                        " AND WMS_IN_TX.IO_DOC_ID   = WMS_GOODSRCV.GR_CODE " & _
                        " INNER JOIN WMS_STORER " & _
                        " ON WMS_IN_TX.IMP_CODE     = WMS_STORER.IMP_CODE " & _
                        " AND WMS_IN_TX.STORER_CODE = WMS_STORER.STORER_CODE " & _
                        " Left Outer Join wms_item  " & _
                        " on wms_in_tx.imp_code = wms_item.imp_code " & _
                        " and wms_in_tx.storer_code = wms_item.storer_code " & _
                        " and wms_in_tx.itm_code = wms_item.itm_code " & _
                        " and wms_in_tx.pack_key = wms_item.pack_key " & _
                        " WHERE   WMS_IN_TX.IO_TYPE='IN' AND WMS_IN_TX.IMP_CODE='" & gU.dbEncode(imp_code) & "'" & _
                        temp_where1 & _
                        " GROUP BY WMS_IN_TX.ITM_CODE, WMS_IN_TX.IO_DOC, WMS_GOODSRCV.GR_REF_NO, WMS_STORER.STO_SHORTNAME, WMS_IN_TX.IO_BATCH_NO,  " & _
                        " WMS_IN_TX.IO_DATETIME, WMS_IN_TX.IO_DOC_ID,wms_item.itm_uom, wms_item.ITM_PCS_PER_UOM, WMS_IN_TX.PACK_KEY, WMS_ITEM.ITM_SKU_NO ) " & _
                        " UNION " & _
                        "   (SELECT WMS_STORER.STO_SHORTNAME, " & _
                        "     '' AS ERP_NO, " & _
                        "     WMS_OUT_TX.ITM_CODE, " & _
                        "   WMS_OUT_TX.PACK_KEY, " & _
                        "   WMS_ITEM.ITM_SKU_NO, " & _
                        "     WMS_OUT_TX.IO_DOC, " & _
                        "     WMS_OUT_TX.IO_DOC_ID, " & _
                        "     WMS_DELV_ORDER.DO_INV_NO                     AS INV_NO, " & _
                        "     Convert(varchar, WMS_OUT_TX.IO_DATETIME,103) AS IO_DATETIME, " & _
                        "     WMS_OUT_TX.IO_DATETIME AS IO_DATETIME_ORD, " & _
                        "     NULL                                         AS IN_IO_QTY, " & _
                        "     SUM(WMS_OUT_TX.IO_QTY)                       AS OUT_IO_QTY, " & _
                        "     ''                                           AS UNPOST_YN, " & _
                        "     WMS_OUT_TX.IO_BATCH_NO, " & _
                        "     0 - SUM(WMS_OUT_TX.IO_QTY * ISNULL(wms_item.ITM_PCS_PER_UOM,1)) AS CURR_QTY, " & _
                        "     wms_item.itm_uom, ISNULL(wms_item.ITM_PCS_PER_UOM,1) as ITM_PCS_PER_UOM, ISNULL(WMS_DELV_ORDER.CUS_NAME, WMS_DELV_ORDER.CUS_CODE) as cus_name, " & _
                        "     '2'                       AS SORT_COL " & _
                        " FROM WMS_OUT_TX " & _
                        "   INNER JOIN WMS_STORER " & _
                        "   ON WMS_OUT_TX.IMP_CODE     = WMS_STORER.IMP_CODE " & _
                        "   AND WMS_OUT_TX.STORER_CODE = WMS_STORER.STORER_CODE " & _
                        "   LEFT OUTER JOIN WMS_DELV_ORDER " & _
                        "   ON WMS_OUT_TX.IMP_CODE     = WMS_DELV_ORDER.IMP_CODE " & _
                        "   AND WMS_OUT_TX.STORER_CODE = WMS_DELV_ORDER.STORER_CODE " & _
                        "   AND WMS_OUT_TX.IO_DOC_ID   = WMS_DELV_ORDER.DO_CODE " & _
                        "   Left Outer Join wms_item  " & _
                        "   on WMS_OUT_TX.imp_code = wms_item.imp_code " & _
                        "   and WMS_OUT_TX.storer_code = wms_item.storer_code " & _
                        "   and WMS_OUT_TX.itm_code = wms_item.itm_code " & _
                        "   and WMS_OUT_TX.pack_key = wms_item.pack_key " & _
                        " WHERE WMS_OUT_TX.IO_TYPE    ='OUT' AND WMS_OUT_TX.IMP_CODE='" & gU.dbEncode(imp_code) & "'" & _
                        temp_where2 & _
                        "   GROUP BY WMS_OUT_TX.IO_DATETIME, WMS_OUT_TX.IO_DOC, WMS_OUT_TX.IO_DOC_ID, WMS_DELV_ORDER.DO_INV_NO, WMS_STORER.STO_SHORTNAME,  " & _
                        " WMS_OUT_TX.IO_BATCH_NO, WMS_OUT_TX.ITM_CODE, wms_item.itm_uom, wms_item.ITM_PCS_PER_UOM,WMS_OUT_TX.PACK_KEY, WMS_ITEM.ITM_SKU_NO,ISNULL(WMS_DELV_ORDER.CUS_NAME, WMS_DELV_ORDER.CUS_CODE)) " & _
                        " UNION " & _
                        "   (SELECT WMS_STORER.STO_SHORTNAME, " & _
                        "     '' AS ERP_NO, " & _
                        " WMS_IN_TX.ITM_CODE, " & _
                        "   WMS_IN_TX.PACK_KEY, " & _
                        "   WMS_ITEM.ITM_SKU_NO, " & _
                        " WMS_IN_TX.IO_DOC, " & _
                        " WMS_IN_TX.IO_DOC_ID, " & _
                        " WMS_DELV_ORDER.DO_INV_NO                    AS INV_NO, " & _
                        " Convert(varchar, WMS_IN_TX.IO_DATETIME,103) AS IO_DATETIME, " & _
                        " WMS_IN_TX.IO_DATETIME AS IO_DATETIME_ORD, " & _
                        " SUM(WMS_IN_TX.IO_QTY)                       AS IN_IO_QTY, " & _
                        " NULL                                        AS OUT_IO_QTY, " & _
                        " 'UNPOST'                                    AS UNPOST_YN, " & _
                        "     WMS_IN_TX.IO_BATCH_NO, " & _
                        "     SUM(WMS_IN_TX.IO_QTY *ISNULL(wms_item.ITM_PCS_PER_UOM,1)) AS CURR_QTY, " & _
                        "     wms_item.itm_uom, ISNULL(wms_item.ITM_PCS_PER_UOM,1) as ITM_PCS_PER_UOM, ISNULL(WMS_DELV_ORDER.CUS_NAME, WMS_DELV_ORDER.CUS_CODE) as cus_name, " & _
                        "     '1'                      AS SORT_COL " & _
                        "   FROM WMS_IN_TX " & _
                        "   LEFT OUTER JOIN WMS_DELV_ORDER " & _
                        "   ON WMS_IN_TX.IMP_CODE     = WMS_DELV_ORDER.IMP_CODE " & _
                        "   AND WMS_IN_TX.STORER_CODE = WMS_DELV_ORDER.STORER_CODE " & _
                        "   AND WMS_IN_TX.IO_DOC_ID   = WMS_DELV_ORDER.DO_CODE " & _
                        "   INNER JOIN WMS_STORER " & _
                        "   ON WMS_IN_TX.IMP_CODE     = WMS_STORER.IMP_CODE " & _
                        "   AND WMS_IN_TX.STORER_CODE = WMS_STORER.STORER_CODE " & _
                        "   Left Outer Join wms_item  " & _
                        "   on wms_in_tx.imp_code = wms_item.imp_code " & _
                        "   and wms_in_tx.storer_code = wms_item.storer_code " & _
                        "   and wms_in_tx.itm_code = wms_item.itm_code " & _
                        "   and wms_in_tx.pack_key = wms_item.pack_key " & _
                        " WHERE WMS_IN_TX.IO_TYPE='UNPOST' AND WMS_IN_TX.IMP_CODE='" & gU.dbEncode(imp_code) & "'" & _
                        temp_where1 & _
                        "   GROUP BY WMS_IN_TX.ITM_CODE, WMS_IN_TX.IO_DOC, WMS_DELV_ORDER.DO_INV_NO, WMS_STORER.STO_SHORTNAME, WMS_IN_TX.IO_BATCH_NO,  " & _
                        " WMS_IN_TX.IO_DATETIME, WMS_IN_TX.IO_DOC_ID,wms_item.itm_uom, wms_item.ITM_PCS_PER_UOM,WMS_IN_TX.PACK_KEY, WMS_ITEM.ITM_SKU_NO, ISNULL(WMS_DELV_ORDER.CUS_NAME, WMS_DELV_ORDER.CUS_CODE))  " & _
                        " UNION  " & _
                        "   (SELECT WMS_STORER.STO_SHORTNAME, " & _
                        "     WMS_GOODSRCV.GR_REF_NO AS ERP_NO, " & _
                        "     WMS_OUT_TX.ITM_CODE, " & _
                        "   WMS_OUT_TX.PACK_KEY, " & _
                        "   WMS_ITEM.ITM_SKU_NO, " & _
                        "     WMS_OUT_TX.IO_DOC, " & _
                        "     WMS_OUT_TX.IO_DOC_ID, " & _
                        "     ''                                           AS INV_NO, " & _
                        "   Convert(varchar, WMS_OUT_TX.IO_DATETIME,103) AS IO_DATETIME, " & _
                        "   WMS_OUT_TX.IO_DATETIME AS IO_DATETIME_ORD, " & _
                        "     NULL                                         AS IN_IO_QTY, " & _
                        "     SUM(WMS_OUT_TX.IO_QTY)                       AS OUT_IO_QTY, " & _
                        "     'UNPOST'                                     AS UNPOST_YN, " & _
                        "     WMS_OUT_TX.IO_BATCH_NO, " & _
                        "     0 - SUM(WMS_OUT_TX.IO_QTY * ISNULL(wms_item.ITM_PCS_PER_UOM,1)) AS CURR_QTY, " & _
                        "     wms_item.itm_uom, ISNULL(wms_item.ITM_PCS_PER_UOM,1) as ITM_PCS_PER_UOM, '' as cus_name, " & _
                        "     '1'                       AS SORT_COL " & _
                        "   FROM WMS_OUT_TX " & _
                        "   INNER JOIN WMS_STORER " & _
                        "   ON WMS_OUT_TX.IMP_CODE     = WMS_STORER.IMP_CODE " & _
                        "   AND WMS_OUT_TX.STORER_CODE = WMS_STORER.STORER_CODE " & _
                        "   LEFT OUTER JOIN WMS_GOODSRCV " & _
                        "   ON WMS_OUT_TX.IMP_CODE     = WMS_GOODSRCV.IMP_CODE " & _
                        "   AND WMS_OUT_TX.STORER_CODE = WMS_GOODSRCV.STORER_CODE " & _
                        "   AND WMS_OUT_TX.IO_DOC_ID   = WMS_GOODSRCV.GR_CODE  " & _
                        "   Left Outer Join wms_item  " & _
                        "   on WMS_OUT_TX.imp_code = wms_item.imp_code " & _
                        "   and WMS_OUT_TX.storer_code = wms_item.storer_code " & _
                        "   and WMS_OUT_TX.itm_code = wms_item.itm_code " & _
                        "   and WMS_OUT_TX.pack_key = wms_item.pack_key " & _
                        " WHERE WMS_OUT_TX.IO_TYPE    ='UNPOST' AND WMS_OUT_TX.IMP_CODE='" & gU.dbEncode(imp_code) & "'" & _
                        temp_where2 & _
                        "   GROUP BY WMS_OUT_TX.IO_DATETIME, WMS_OUT_TX.IO_DOC, WMS_GOODSRCV.GR_REF_NO, WMS_OUT_TX.IO_DOC_ID, WMS_STORER.STO_SHORTNAME,  " & _
                        " WMS_OUT_TX.IO_BATCH_NO, WMS_OUT_TX.ITM_CODE, wms_item.itm_uom, wms_item.ITM_PCS_PER_UOM,WMS_OUT_TX.PACK_KEY, WMS_ITEM.ITM_SKU_NO ) " & _
                        " ORDER BY IO_DATETIME_ORD,  SORT_COL  "


            nDataSource = gDB.getDataTable(sqlString)


            Dim total_sql As String = ""
            Dim tobal_dt As New DataTable
            Dim total_bal As Decimal = 0



            'total_sql = "SELECT WMS_ITEM.ITM_CODE, SUM(ISNULL(case when aitm_pcs_per_pack=0 then to_number(WMS_ITEM.PACK_KEY) else aitm_pcs_per_pack end,to_number(WMS_ITEM.PACK_KEY)) * ITM_BALANCE) AS Total_bal FROM WMS_ITEM " & _
            '            " LEFT OUTER JOIN (SELECT     MAX(AITM_PCS_PER_PACK) AS aitm_pcs_per_pack, IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY " & _
            '            " FROM WMS_ALT_VEND_ITEM d GROUP BY IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY) t ON WMS_ITEM.IMP_CODE = t.IMP_CODE AND " & _
            '            " WMS_ITEM.STORER_CODE = t.STORER_CODE And WMS_ITEM.ITM_CODE = t.ITM_CODE And WMS_ITEM.PACK_KEY = t.PACK_KEY " & _
            '            " WHERE WMS_ITEM.ITM_CODE = '" & gU.dbEncode(itm_code) & "'" & temp_total_where & " GROUP BY WMS_ITEM.ITM_CODE "

            'tobal_dt = gDB.getDataTable(total_sql)

            'If tobal_dt.Rows.Count > 0 Then
            '    total_bal = gU.decodeNullOrEmpty(tobal_dt.Rows(0).Item("Total_bal").ToString, 0)
            'End If

            'Dim keyList As New ArrayList
            'wmsFun.getCartonQty(Total_crt, "PAD_CARTON_NO", keyList, True, , nDataSource)
            Dim ReportPara(2) As ReportParameter

            ReportPara(0) = New ReportParameter("total_bal", 0)


            If nDataSource.Rows.Count > 0 Then
                ReportPara(1) = New ReportParameter("pack_key", nDataSource.Rows(0).Item("PACK_KEY").ToString.Trim)
                ReportPara(2) = New ReportParameter("sku_no", nDataSource.Rows(0).Item("ITM_SKU_NO").ToString.Trim)

                reportSource(nDataSource, ReportPara)
            Else
                Response.Write("No Item Record Found.")
            End If
        End If
    End Sub

    Private Sub reportSource(ByVal sourcetbl As DataTable, Optional ByVal paraarray() As ReportParameter = Nothing)
        Try
            ReportViewer1.LocalReport.ReportPath = "REPORT\SING_ITEM\singitem.rdlc"

            ReportViewer1.LocalReport.DataSources.Clear()

            ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("DataSet1", sourcetbl))

            If Not IsNothing(paraarray) Then
                ReportViewer1.LocalReport.SetParameters(paraarray)
            End If

            Try
                ReportViewer1.LocalReport.Refresh()

            Catch ex As OutOfMemoryException
                GC.Collect()
                ReportViewer1.LocalReport.Refresh()
            End Try

        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    Public Function isValidDate(ByVal aDate As String, Optional ByVal dateFormat As String = "") As Boolean
        Dim myculture As New System.Globalization.CultureInfo("en-US")
        Dim datePattern As String
        Dim inputDate As String = aDate

        If dateFormat = "" Then
            datePattern = "dd/MM/yyyy"
        ElseIf InStr(dateFormat, "/") <= 0 Then
            inputDate = chgDateFormat(inputDate, dateFormat, "dd/MM/yyyy")
            datePattern = "dd/MM/yyyy"
        Else
            datePattern = dateFormat
        End If

        datePattern = Replace(Replace(UCase(datePattern), "DD", "dd"), "YY", "yy")

        myculture.DateTimeFormat.ShortDatePattern = datePattern
        myculture.DateTimeFormat.LongDatePattern = datePattern

        If DateTime.TryParse(inputDate, myculture, System.Globalization.DateTimeStyles.None, New DateTime) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function chgDateFormat(ByVal aDate As String, ByVal orgFormat As String, ByVal tarFormat As String) As String
        Dim yearStr, monthStr, dayStr As String
        Dim orgDateFormat As String = UCase(orgFormat.Trim)
        Dim tarDateFormat As String = UCase(tarFormat.Trim)
        Dim inputDate As String = UCase(aDate.Trim)
        Dim resultDate As String

        If inputDate <> "" Then
            resultDate = tarDateFormat

            If InStr(orgDateFormat, "YYYY") Then
                yearStr = Mid(inputDate, InStr(orgDateFormat, "YYYY"), 4)
                resultDate = Replace(resultDate, "YYYY", yearStr)
            Else
                yearStr = Mid(inputDate, InStr(orgDateFormat, "YY"), 2)
                resultDate = Replace(resultDate, "YY", yearStr)
            End If

            monthStr = Mid(inputDate, InStr(orgDateFormat, "MM"), 2)
            resultDate = Replace(resultDate, "MM", monthStr)

            dayStr = Mid(inputDate, InStr(orgDateFormat, "DD"), 2)
            resultDate = Replace(resultDate, "DD", dayStr)
        Else
            resultDate = ""
        End If

        Return resultDate
    End Function
End Class
