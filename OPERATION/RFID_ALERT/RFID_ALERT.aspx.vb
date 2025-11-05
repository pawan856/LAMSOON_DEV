Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Imports System.Reflection
Imports System.Collections.Generic
Imports Microsoft.Reporting.WebForms

Partial Class OPERATION_RFID_ALERT_RFID_ALERT
    Inherits System.Web.UI.Page
    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private cm As CommonMenu
    Private st As New StockTrans

    Dim checkBoxValue() As String = {"Y", ""}

    Private moduleAction As String = ""
    Private dt As New DataTable
    Private r_dt As DataTable
    Private DDFORMAT As String = "DD/MM/YYYY"


    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        DDFORMAT = gU.getConfig("DDFORMATNO")

        REM ****************************
        REM Modify Access Right Here
        ar = New AccessRightUtils(Session("PAGE_SESSION_MENU_CODE"), Session("usr_id"), Me)

        moduleAction = Request("moduleAction")

        If ar.sessionExpired = "Y" Then
            ar.Force_PageEndCtrlClear(Me, False)
            Me.Visible = False
            Exit Sub
        End If

        If Session("PAGE_SESSION_MENU_CODE") Is Nothing Then
            Exit Sub
        End If

        If Not IsPostBack Then
            ViewState("dt") = Nothing
            'BindGV()
        End If

    End Sub

    Private Sub BindGV()
        Dim selectSQL As String = " SELECT '' as chkNew, WRST_CODE, WRAL_KEY, WRMV_BATCH_ID, WRMV_TAG_ID,convert(varchar,cast(WRMV_BATCH_TIME as date),103) as WRMV_BATCH_DATE, " & _
                                  " cast(WRMV_BATCH_TIME as time(0)) as WRMV_BATCH_TIME, WRMV_BATCH_TIME  as WRMV_BATCH_TIME_FULL, WRAL_IMP_CODE, WRAL_STORER_CODE, WRAL_ITM_CODE, WRAL_PACK_KEY, WRAL_ITM_SKU_NO, WRAL_ITM_NAME, " & _
                                  " WRAL_DSP_ALERT_TYPE,CASE WHEN WRAL_DSP_ALERT_TYPE = 'STOCK_MISS' then 'Possible Stock Out' else Null END as ALERT_DESC, WRMV_STOCK_OUT_TIME, WRAL_STATUS, SYS_LUB, SYS_LUD, SYS_CD, SYS_CB " & _
                                  " FROM WMS_RFID_ALERT_WMS " & _
                                  " WHERE 1=1 and WRAL_DSP_ALERT_TYPE NOT IN ('SURPLUS_OUT') "

        Dim tempStr As String = ""

        If Not String.IsNullOrWhiteSpace(gU.getChkBoxListValue(src_Alerts)) Then
            Dim valueString As String = "'" & gU.getChkBoxListValue(src_Alerts).Replace(", ", "', '") & "'"
            tempStr &= " AND WRAL_STATUS IN (" & valueString & ")"
        End If


        If Not String.IsNullOrWhiteSpace(gU.getChkBoxListValue(src_Status)) Then
            Dim valueString As String = "'" & gU.getChkBoxListValue(src_Status).Replace(", ", "', '") & "'"
            tempStr &= " AND WRAL_DSP_ALERT_TYPE IN (" & valueString & ")"
        End If

        If Not String.IsNullOrWhiteSpace(FR_DATE.Text.Trim) Then
            tempStr &= " AND WRMV_BATCH_TIME >= convert(datetime,'" & FR_DATE.Text.Trim & "',103)"
        End If

        If Not String.IsNullOrWhiteSpace(TO_DATE.Text.Trim) Then
            tempStr &= " AND WRMV_BATCH_TIME < convert(datetime,'" & TO_DATE.Text.Trim & "',103) + 1"
        End If

        selectSQL &= tempStr & " and LEN(WRMV_TAG_ID) <= 16 ORDER BY WRMV_BATCH_TIME_FULL DESC"

        dt = gDB.getDataTable(selectSQL)

        If dt.Rows.Count > 0 Then
            Gridview1.DataSource = dt
            total_count.Text = dt.Rows.Count

        Else
            Gridview1.DataSource = Nothing
        End If

        ViewState("dt") = dt
        Gridview1.DataBind()


    End Sub

    Protected Sub Gridview1_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles Gridview1.RowCommand
        Select Case e.CommandName
            Case "VIDEO"
                Dim gvRow As GridViewRow = CType(CType(e.CommandSource, Control).NamingContainer, GridViewRow)
                Dim WRST_CODE As String = ""
                Dim WRAL_KEY As String = ""

                If Not IsNothing(gvRow) Then
                    WRST_CODE = DirectCast(gvRow.FindControl("WRST_CODE"), HiddenField).Value
                    WRAL_KEY = DirectCast(gvRow.FindControl("WRAL_KEY"), HiddenField).Value
                End If
                showVDOPanel(WRST_CODE, WRAL_KEY)
        End Select
    End Sub

    Protected Sub Gridview1_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles Gridview1.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow

                CType(e.Row.FindControl("WRAL_KEY"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "WRAL_KEY").ToString.Trim
                CType(e.Row.FindControl("WRST_CODE"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "WRST_CODE").ToString.Trim
                CType(e.Row.FindControl("WRAL_DSP_ALERT_TYPE"), Label).Text = DataBinder.Eval(e.Row.DataItem, "WRAL_DSP_ALERT_TYPE").ToString.Trim
                CType(e.Row.FindControl("ALERT_DESC"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ALERT_DESC").ToString.Trim

                If DataBinder.Eval(e.Row.DataItem, "chkNew").ToString.Trim = "Y" Then
                    CType(e.Row.FindControl("chkNew"), CheckBox).Checked = True
                End If

        End Select
    End Sub

    Protected Sub Gridview1_PageIndexChanged(sender As Object, e As System.EventArgs) Handles Gridview1.PageIndexChanged
        Dim ddlPager As DropDownList = TryCast(Gridview1.BottomPagerRow.FindControl("pager_select"), DropDownList)
        Dim tempDT As DataTable = ViewState("dt")

        If cU.gfBuildDataTableforGridView(tempDT, Gridview1, True, checkBoxValue) Then
            If ddlPager.SelectedValue <> Gridview1.PageIndex + 1 Then
                Gridview1.PageIndex = ddlPager.SelectedValue - 1
                Gridview1.DataSource = tempDT
                Gridview1.DataBind()
            ElseIf ddlPager.SelectedValue = "1" Then
                Gridview1.PageIndex = 0
                Gridview1.DataSource = tempDT
                Gridview1.DataBind()
            End If
        End If
    End Sub

    Protected Sub Gridview1_PageIndexChanging(sender As Object, e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles Gridview1.PageIndexChanging
        Dim tempDT As DataTable = ViewState("dt")

        If cU.gfBuildDataTableforGridView(tempDT, Gridview1, True, checkBoxValue) Then
            Gridview1.PageIndex = e.NewPageIndex
            Gridview1.DataSource = tempDT
            Gridview1.DataBind()
        End If
    End Sub

    Protected Sub Gridview1_DataBound(sender As Object, e As System.EventArgs) Handles Gridview1.DataBound
        If Gridview1.DataSource IsNot Nothing Then
            Dim ddlPager As DropDownList = TryCast(Gridview1.BottomPagerRow.FindControl("pager_select"), DropDownList)
            Dim tmpPagerBtn As ImageButton

            If ddlPager.Items.Count = 0 Then
                If ddlPager IsNot Nothing Then
                    For i As Integer = 1 To Gridview1.PageCount
                        ddlPager.Items.Add(New ListItem(i.ToString(), i.ToString()))
                    Next
                End If
            End If

            ddlPager.SelectedValue = Gridview1.PageIndex + 1

            If Gridview1.PageIndex = 0 Then
                tmpPagerBtn = TryCast(Gridview1.BottomPagerRow.FindControl("pager_first"), ImageButton)
                tmpPagerBtn.Enabled = False

                tmpPagerBtn = TryCast(Gridview1.BottomPagerRow.FindControl("pager_previous"), ImageButton)
                tmpPagerBtn.Enabled = False
            End If

            If Gridview1.PageIndex = Gridview1.PageCount - 1 Then
                tmpPagerBtn = TryCast(Gridview1.BottomPagerRow.FindControl("pager_next"), ImageButton)
                tmpPagerBtn.Enabled = False

                tmpPagerBtn = TryCast(Gridview1.BottomPagerRow.FindControl("pager_last"), ImageButton)
                tmpPagerBtn.Enabled = False
            End If

        End If
    End Sub

    Protected Sub btnSearch_Click(sender As Object, e As System.EventArgs) Handles btnSearch.Click
        BindGV()
    End Sub

    Protected Sub btnReset_Click(sender As Object, e As System.EventArgs) Handles btnReset.Click
        Dim rmtPost As New RemotePost
        rmtPost.Url = "RFID_ALERT.aspx"
        rmtPost.Post()
    End Sub

    Protected Sub btnRead_Click(sender As Object, e As System.EventArgs) Handles btnRead.Click
        Dim tempDT As DataTable
        Dim gConn As SqlConnection
        Dim updateSQL As String = ""
        Dim successFlag As Boolean = False
        tempDT = ViewState("dt")

        If tempDT IsNot Nothing AndAlso tempDT.Rows.Count > 0 Then
            gConn = gDB.getConnection()
            Dim transaction As SqlTransaction

            transaction = gConn.BeginTransaction()

            Try
                If cU.gfBuildDataTableforGridView(tempDT, Gridview1, True, checkBoxValue) Then
                    For i = 0 To tempDT.Rows.Count - 1
                        If tempDT.Rows(i).Item("chkNew").ToString.Trim = "Y" Then

                            updateSQL = "UPDATE WMS_RFID_ALERT_WMS SET " & _
                                        "WRAL_STATUS='CHECKED', SYS_LUD=getdate(), SYS_LUB='" & gU.dbEncode(Session("usr_id")) & "' " & _
                                        "WHERE WRST_CODE='" & gU.dbEncode(tempDT.Rows(i).Item("WRST_CODE").ToString.Trim) & "' " & _
                                        " AND WRAL_KEY='" & gU.dbEncode(tempDT.Rows(i).Item("WRAL_KEY").ToString.Trim) & "' "

                            gDB.amendData(updateSQL, gConn, transaction)

                        End If
                    Next

                    transaction.Commit()
                    successFlag = True
                End If
            Catch ex As Exception
                transaction.Rollback()
                Response.Write(ex.Message)
                uiFun.displayMsgNew(gvUDP, "1008", "", Session("gLang"))
            Finally
                If gConn IsNot Nothing Then
                    If gConn.State = ConnectionState.Open Then
                        gConn.Close()
                        gConn.Dispose()
                    End If
                End If
            End Try
        End If


        If successFlag Then
            Call BindGV()
            uiFun.displayMsgNew(gvUDP, "", "Status Successfully Changed.", Session("gLang"))
        End If
    End Sub

    Protected Sub showVDOPanel(ByVal WRST_CODE As String, ByVal WRAL_KEY As String)

        Dim sqlString As String = "Select WRMV_BATCH_TIME from WMS_RFID_ALERT_WMS WHERE WRST_CODE='" & WRST_CODE & "' and WRAL_KEY='" & WRAL_KEY & "'"

        Dim WRMV_BATCH_TIME As DateTime = Nothing

        Dim tempdt As DataTable = gDB.getDataTable(sqlString)

        If tempdt.Rows.Count > 0 Then
            If tempdt.Rows(0).Item("WRMV_BATCH_TIME").ToString.Trim <> "" Then
                WRMV_BATCH_TIME = tempdt.Rows(0).Item("WRMV_BATCH_TIME")
                GetFileList(WRMV_BATCH_TIME)
            End If

        Else
            GVVideo.DataSource = Nothing
            GVVideo.DataBind()
        End If

        pnlVDO_ModalPopupExtender.Show()
        
    End Sub

    Protected Function GetFileListF(ByVal iDateTime As DateTime) As Boolean
        Dim ftp As String = gU.getConfig("FTP_SERVER")
        'FTP Folder name. Leave blank if you want to list files from root folder.
        Dim ftpFolder As String = gU.getConfig("FTP_VIDEO_FOLDER")

        Dim ftpID As String = gU.getConfig("FTP_ID")
        Dim ftpPWD As String = gU.getConfig("FTP_PASSWORD")
        Dim returnFlag As Boolean = True
        'Dim dateFolder As String()

        Dim videoFileFolder As String = "record\"




        Dim dateFolder As String = "20161223\"

        ftpFolder = ftpFolder & dateFolder & videoFileFolder

        Try
            Dim dtFiles As New DataTable()
            dtFiles.Columns.AddRange(New DataColumn(2) { _
                                     New DataColumn("FileName", GetType(String)), _
                                     New DataColumn("Size", GetType(Decimal)), _
                                     New DataColumn("videoDate", GetType(String))})
            'Create FTP Request.

            Dim request As FtpWebRequest = DirectCast(WebRequest.Create(ftp & ftpFolder), FtpWebRequest)

            request.Method = WebRequestMethods.Ftp.ListDirectoryDetails
            'Enter FTP Server credentials.
            request.Credentials = New NetworkCredential(ftpID, ftpPWD)
            request.UsePassive = True
            request.UseBinary = True
            request.EnableSsl = False

            'Fetch the Response and read it using StreamReader.
            Dim response As FtpWebResponse = DirectCast(request.GetResponse(), FtpWebResponse)
            Dim entries As New List(Of String)()

            Using reader As New StreamReader(response.GetResponseStream())

                'Read the Response as String and split using New Line character.

                entries = reader.ReadToEnd().Split(New String() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries).ToList()

            End Using

            response.Close()



            'Create a DataTable.


            'Loop and add details of each File to the DataTable.

            For Each entry As String In entries

                Dim splits As String() = entry.Split(New String() {" "}, StringSplitOptions.RemoveEmptyEntries)

                'Determine whether entry is for File or Directory.

                'Dim isFile As Boolean = splits(0).Substring(0, 1) <> "d"
                Dim isFile As Boolean = True
                'Dim isDirectory As Boolean = splits(0).Substring(0, 1) = "d"

                'If entry is for File, add details to DataTable.

                If isFile Then

                    dtFiles.Rows.Add()
                    dtFiles.Rows(dtFiles.Rows.Count - 1)("Size") = Decimal.Parse(splits(2)) / 1024
                    dtFiles.Rows(dtFiles.Rows.Count - 1)("videoDate") = String.Join(" ", splits(0), splits(1))
                    dtFiles.Rows(dtFiles.Rows.Count - 1)("FileName") = String.Join(" ", splits(3))
                End If

            Next

            'Bind the GridView.
            GVVideo.DataSource = dtFiles
            GVVideo.DataBind()

        Catch ex As WebException
            uiFun.displayMsgNew(gvUDP, "", ex.Message, Session("gLang"))
            'Throw New Exception(TryCast(ex.Response, FtpWebResponse).StatusDescription)
            returnFlag = False
        End Try
        Return returnFlag
    End Function

    Protected Sub GVVideo_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GVVideo.RowCommand
         Select e.CommandName
            Case "DOWNLOAD"
                Dim gvRow As GridViewRow = CType(CType(e.CommandSource, Control).NamingContainer, GridViewRow)
                Dim camFolder As String = DirectCast(gvRow.FindControl("CameraFolder"), HiddenField).Value

                DownloadFile(e.CommandArgument, camFolder)
                pnlVDO_ModalPopupExtender.Show()
        End Select
    End Sub

    Protected Sub GVVideo_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GVVideo.RowDataBound
          Select e.Row.RowType
            Case DataControlRowType.DataRow

                CType(e.Row.FindControl("FileName"), LinkButton).Text = DataBinder.Eval(e.Row.DataItem, "fileName").ToString.Trim
                CType(e.Row.FindControl("FileName"), LinkButton).CommandArgument = DataBinder.Eval(e.Row.DataItem, "fileName").ToString.Trim

                CType(e.Row.FindControl("CameraFolder"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "CameraFolder").ToString.Trim
        End Select

    End Sub


    Protected Function DownloadFileF(ByVal filename As String) As Boolean

        'FTP Server URL.

        Dim ftp As String = gU.getConfig("FTP_SERVER")
        'FTP Folder name. Leave blank if you want to list files from root folder.
        Dim ftpFolder As String = gU.getConfig("FTP_VIDEO_FOLDER")

        Dim ftpID As String = gU.getConfig("FTP_ID")
        Dim ftpPWD As String = gU.getConfig("FTP_PASSWORD")

        'Dim dateFolder As String()

        Dim videoFileFolder As String = "record\"
        Dim dateFolder As String = ""

        Dim nameArr As String() = Split(filename, "_")




        ftpFolder = ftpFolder & dateFolder & videoFileFolder
        'FTP Folder name. Leave blank if you want to Download file from root folder.

        Try

            'Create FTP Request.

            Dim request As FtpWebRequest = DirectCast(WebRequest.Create(Convert.ToString(ftp & ftpFolder) & filename), FtpWebRequest)
            request.Method = WebRequestMethods.Ftp.DownloadFile

            'Enter FTP Server credentials.
            request.Credentials = New NetworkCredential(ftpID, ftpPWD)
            request.UsePassive = True
            request.UseBinary = True
            request.EnableSsl = False

            'Fetch the Response and read it into a MemoryStream object.

            Dim resp As FtpWebResponse = DirectCast(request.GetResponse(), FtpWebResponse)

            Using stream As New MemoryStream()

                'Download the File.
                resp.GetResponseStream().CopyTo(stream)
                Response.AddHeader("content-disposition", "attachment;filename=" & filename)
                Response.Cache.SetCacheability(HttpCacheability.NoCache)
                Response.BinaryWrite(stream.ToArray())
                Response.End()

            End Using

        Catch ex As WebException

            'Throw New Exception(TryCast(ex.Response, FtpWebResponse).StatusDescription)
            uiFun.displayMsgNew(gvUDP, "", ex.Message, Session("gLang"))
            Return False
        End Try

        Return True

    End Function

    Protected Function GetFileList(ByVal WRMV_BATCH_TIME As DateTime) As Boolean
        Dim fileFolder As String = gU.getConfig("SHARE_VIDEO_FOLDER")
        Dim returnFlag As Boolean = False
        Dim videoFileFolder As String = "record\"
        Dim dateFolder As String = ""

        Dim dateArr(0) As String

        If IsDate(WRMV_BATCH_TIME) Then

            If WRMV_BATCH_TIME.AddMinutes(10).Date <> WRMV_BATCH_TIME.Date OrElse WRMV_BATCH_TIME.AddMinutes(-10).Date <> WRMV_BATCH_TIME.Date Then

                Array.Resize(dateArr, dateArr.Length + 1)
                dateArr(0) = WRMV_BATCH_TIME.AddMinutes(-10).ToString("yyyyMMdd")
                dateArr(1) = WRMV_BATCH_TIME.AddMinutes(10).ToString("yyyyMMdd")
            Else
                dateArr(0) = WRMV_BATCH_TIME.ToString("yyyyMMdd")
            End If


                'dateFolder = WRMV_BATCH_TIME.ToString("yyyyMMdd")

            Dim dtFiles As New DataTable()
            dtFiles.Columns.AddRange(New DataColumn(4) { _
                                     New DataColumn("CameraName", GetType(String)), _
                                     New DataColumn("CameraFolder", GetType(String)), _
                                     New DataColumn("FileName", GetType(String)), _
                                     New DataColumn("Size", GetType(Decimal)), _
                                     New DataColumn("videoDate", GetType(DateTime))})
            Dim newRow As DataRow

            Dim folderDT As DataTable

            folderDT = gDB.getDataTable("Select colc_code, COLC_ENG_VALUE from wms_col_code where COLC_TABCOL='BUFF_CAMERA'")

            Dim cameraROOT As String = ""
            Try
                For z = 0 To folderDT.Rows.Count - 1
                    cameraROOT = folderDT(z).Item("colc_code").ToString.Trim

                    For i = 0 To dateArr.Length - 1
                        dateFolder = dateArr(i)

                        If Directory.Exists(fileFolder & cameraROOT & "\" & dateFolder & "\" & videoFileFolder) Then
                            Dim orderedFiles = New System.IO.DirectoryInfo(fileFolder & dateFolder & "\" & videoFileFolder).GetFiles().Where(Function(x) x.CreationTime >= WRMV_BATCH_TIME.AddMinutes(-10) And x.CreationTime <= WRMV_BATCH_TIME.AddMinutes(10)).OrderByDescending(Function(x) x.CreationTime)

                            Dim fri As FileInfo

                            For Each fri In orderedFiles
                                newRow = dtFiles.NewRow

                                newRow.Item("CameraName") = folderDT(z).Item("COLC_ENG_VALUE").ToString.Trim
                                newRow.Item("CameraFolder") = cameraROOT
                                newRow.Item("FileName") = fri.Name
                                newRow.Item("videoDate") = fri.CreationTime
                                newRow.Item("Size") = fri.Length / 1024

                                dtFiles.Rows.Add(newRow)
                            Next fri
                        End If
                    Next
                Next
                dtFiles.AcceptChanges()

                GVVideo.DataSource = dtFiles
                GVVideo.DataBind()

                Return True

            Catch ex As Exception



                Return False
            End Try

        End If

        If Not returnFlag Then
            GVVideo.DataSource = Nothing
            GVVideo.DataBind()
        End If


        Return returnFlag
    End Function

    Protected Function DownloadFile(ByVal fileName As String, ByVal cameraFolder As String) As Boolean
        Dim fileFolder As String = gU.getConfig("SHARE_VIDEO_FOLDER")
        Dim returnFlag As Boolean = False
        Dim videoFileFolder As String = "record\"
        Dim dateFolder As String = ""

        Dim fileArr As String() = Split(fileName, "_")

        If fileArr.Length <= 1 Then
            uiFun.displayMsgNew(gvUDP, "", "File Name Error. Please check Video folder files.", Session("gLang"))
            Return False
        End If

        Dim fileDate As String = fileArr(0)

        dateFolder = "20" & Right(fileDate, 6)

        Dim nFile As System.IO.FileInfo = New System.IO.FileInfo(fileFolder & cameraFolder & "\" & dateFolder & "\" & videoFileFolder & fileName)
        Try

            If nFile.Exists Then
                Response.Clear()
                Response.AddHeader("Content-Disposition", "attachment; filename=" & nFile.Name)
                Response.AddHeader("Content-Length", nFile.Length.ToString())
                Response.ContentType = "application/octet-stream"
                Response.WriteFile(nFile.FullName)
                Response.End()
            Else
                Response.Write("This file does not exist.")
            End If

        Catch ex As WebException

            'Throw New Exception(TryCast(ex.Response, FtpWebResponse).StatusDescription)
            uiFun.displayMsgNew(gvUDP, "", ex.Message, Session("gLang"))
            Return False
        End Try

        Return True
    End Function

    Protected Sub btndownRPT_Click(sender As Object, e As System.EventArgs) Handles btndownRPT.Click
        Dim extension As String = String.Empty
        Dim mimeType As String = String.Empty

        Dim nDataSource As DataTable = ViewState("dt")


        If nDataSource Is Nothing Then
            uiFun.displayMsgNew(gvUDP, "", "No Record Found", Session("gLang"))
            Exit Sub
        End If


        Select Case reportFormat.SelectedValue
            Case "PDF"
                extension = "pdf"
                mimeType = "application/pdf"
                Exit Select
                'Case "EXCEL"
                '    extension = "xls"
                '    mimeType = "application/vnd.excel"
                '    Exit Select
            Case "EXCELOPENXML"
                extension = "xlsx"
                mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                Exit Select
            Case "WORD"
                extension = "doc"
                mimeType = "application/vnd.ms-word"
                Exit Select
            Case "IMAGE"
                extension = "emf"
                mimeType = "application/image"
                Exit Select
            Case Else
                Throw New Exception("Unrecognized type: " & reportFormat.SelectedValue & ". Type must be PDF, Excel, Word or Image.")
        End Select

        Dim deviceInfo As String = "<DeviceInfo>" & _
                            " <OutputFormat>" & reportFormat.SelectedValue & "</OutputFormat>" & _
                            "</DeviceInfo>"
        Dim encoding As String = ""
        Dim warnings As Warning() = Nothing
        Dim streams As String() = Nothing
        Dim byteviewer As Byte()

        Dim rptViewer As New ReportViewer

        rptViewer.LocalReport.ReportPath = "OPERATION\RFID_ALERT\ALERT_TMP.RDLC"
        rptViewer.LocalReport.DataSources.Clear()
        rptViewer.LocalReport.DataSources.Add(New ReportDataSource("DataSet1", nDataSource))

        rptViewer.LocalReport.EnableExternalImages = True

        'If Not IsNothing(paraarray) Then
        '    rptViewer.LocalReport.SetParameters(paraarray)
        'End If

        Try

            byteviewer = rptViewer.LocalReport.Render(reportFormat.SelectedValue, deviceInfo, mimeType, encoding, extension, streams, warnings)

            Response.Buffer = True
            Response.Clear()
            Response.ContentType = mimeType
            Response.AddHeader("content-disposition", "attachment; filename=RFID_ALERT" & Now.Date.ToString("ddMMyyyy") & "." & extension)
            Response.BinaryWrite(byteviewer)
            Response.End()

        Catch ex As Exception
            Dim str As String = ex.Message
        End Try

    End Sub
End Class
