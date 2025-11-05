Imports Microsoft.VisualBasic
Imports System.Net
Imports System.IO
Imports System.Data
Imports Microsoft.Reporting.WebForms
Imports System.Web.Security
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports System.Web.UI.HtmlControls
Imports System.Xml
Imports System.Xml.Xsl
Imports System.Configuration
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Drawing.Printing
Imports System.Collections.Generic

Public Class ReportUtils

    Private ngDB As New GlobalDBFunc
    Private gU As New GeneralUtils

    Private storingPath As String = ConfigurationManager.AppSettings.Item("TMP_FILE_PATH")
    Private nowTime As DateTime = Now

    Private physicalXslt As String = ""
    Private physicalRdlc As String = ""
    Private physicalSchema As String = ""

    Private ServerIP As String = System.Configuration.ConfigurationManager.AppSettings.Item("WEB_SERVER_IP")

    Public Enum PageSize
        A4 = 0
        A5 = 1
        A3 = 2
        B4 = 3
    End Enum

    Public Enum PageOrentation
        Portrait = 0
        Landscape = 1
    End Enum

    Public Function RenderReport(ByVal PrgmCode As String, ByVal data As DataTable, ByVal rpt As DataTable, _
                                 ByVal name As String, ByVal type As String, _
                                 ByVal table_name As String, ByVal gLanguage As String, _
                                 Optional ByVal report_title As String = "", _
                                 Optional ByVal clientScript As ClientScriptManager = Nothing, _
                                 Optional ByVal RenFromSearch As Boolean = False, _
                                 Optional ByVal mePage As Page = Nothing, _
                                 Optional ByVal Font_Size As String = "8pt", _
                                 Optional ByVal Page_Size As PageSize = PageSize.A4, _
                                 Optional ByVal Page_Orent As PageOrentation = PageOrentation.Portrait,
                                 Optional ByVal Border_Style As String = "", _
                                 Optional ByVal ShowHeader As String = "Y", _
                                 Optional ByVal ShowFooter As String = "Y") As Byte()

        Try
            Dim rds As New ReportDataSource()

            rds.Name = "DT" & name
            rds.Value = data
            Dim virtualRdlc As String = ""

            Dim PageWidth As String = "226cm"
            Dim PageHeight As String = "35.56cm"

            Select Case Page_Size
                Case PageSize.A3
                    If Page_Orent = PageOrentation.Portrait Then
                        PageWidth = "297mm"
                        PageHeight = "420mm"
                    ElseIf Page_Orent = PageOrentation.Landscape Then
                        PageWidth = "420mm"
                        PageHeight = "297mm"
                    End If
                Case PageSize.A4
                    If Page_Orent = PageOrentation.Portrait Then
                        PageWidth = "210mm"
                        PageHeight = "297mm"
                    ElseIf Page_Orent = PageOrentation.Landscape Then
                        PageWidth = "297mm"
                        PageHeight = "210mm"
                    End If
                Case PageSize.A5
                    If Page_Orent = PageOrentation.Portrait Then
                        PageWidth = "148mm"
                        PageHeight = "210mm"
                    ElseIf Page_Orent = PageOrentation.Landscape Then
                        PageWidth = "210mm"
                        PageHeight = "148mm"
                    End If
                Case PageSize.B4
                    If Page_Orent = PageOrentation.Portrait Then
                        PageWidth = "250mm"
                        PageHeight = "353mm"
                    ElseIf Page_Orent = PageOrentation.Landscape Then
                        PageWidth = "353mm"
                        PageHeight = "250mm"
                    End If
            End Select

            If RenFromSearch Then
                virtualRdlc = BuildRDLC_SearchCol(data, rpt, name, table_name.Replace(",", "_"), gLanguage, report_title, Font_Size, PageWidth, PageHeight, Border_Style, ShowHeader, ShowFooter)
            Else
                virtualRdlc = BuildRDLC(data, rpt, name, table_name.Replace(",", "_"), gLanguage, report_title, Font_Size, PageWidth, PageHeight, Border_Style, ShowHeader, ShowFooter)
            End If
            Dim lr As New LocalReport()

            lr.ReportPath = physicalRdlc
            lr.DataSources.Add(rds)

            If type = "PREVIEW" Then
                HttpContext.Current.Session("rdlc_DataSource") = rds
                HttpContext.Current.Session("rdlc_Path") = physicalRdlc
                HttpContext.Current.Session("schema_Path") = physicalSchema

                If report_title = "" Then
                    HttpContext.Current.Session("rdlc_report_title") = name & " Report"
                Else
                    HttpContext.Current.Session("rdlc_report_title") = report_title
                End If

                Dim strScript As String = "window.open('./cms_preview.aspx?menu_code=" & PrgmCode & "','rpt_preview_" & name & _
                                            "','menubar=no,scrollbars=yes,resizable=yes,width=1000,height=760,left=5,top=10');"

                If Not clientScript Is Nothing Then
                    If mePage IsNot Nothing Then
                        If (Not clientScript.IsStartupScriptRegistered(mePage.GetType(), "")) Then
                            clientScript.RegisterStartupScript(mePage.GetType(), "", strScript, True)
                        End If
                    Else
                        If (Not clientScript.IsStartupScriptRegistered(Me.GetType(), "")) Then
                            clientScript.RegisterStartupScript(Me.GetType(), "", strScript, True)
                        End If
                    End If
                End If

                Return Nothing
            Else
                Return RenderReport(lr, name, type, PageWidth, PageHeight)
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function RenderReport(ByVal lr As Microsoft.Reporting.WebForms.LocalReport, ByVal name As String, ByVal type As String, _
                                 ByVal Page_Width As String, ByVal Page_Height As String) As Byte()
        Dim extension As String = String.Empty
        Dim mimeType As String = String.Empty

        Select Case type.ToUpper
            Case "PDF"
                extension = "pdf"
                mimeType = "application/pdf"
                Exit Select
            Case "EXCEL"
                extension = "xls"
                mimeType = "application/vnd.excel"
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
                Throw New Exception("Unrecognized type: " & type & ". Type must be PDF, Excel, Word or Image.")
        End Select

        Dim deviceInfo As String = "<DeviceInfo>" & _
                            " <OutputFormat>" & type & "</OutputFormat>" & _
                            "</DeviceInfo>"
        Dim encoding As String = ""
        Dim warnings As Warning() = Nothing
        Dim streams As String() = Nothing
        Dim result As Byte()

        Try
            result = lr.Render(type, deviceInfo, mimeType, encoding, extension, streams, warnings)
            Call deleteTempFiles(physicalRdlc, physicalSchema)
            HttpContext.Current.Response.Clear()
            HttpContext.Current.Response.ContentType = mimeType
            HttpContext.Current.Response.AddHeader("content-disposition", ("attachment; filename=" & name & Format(nowTime, "hhmmssfff") & ".") & extension)
            HttpContext.Current.Response.BinaryWrite(result)
            'HttpContext.Current.Response.End()

            Return result

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function TransformXml(ByVal xml As String, ByVal xslFile As String) As String

        Dim result As String = String.Empty

        Using memory As New System.IO.MemoryStream(System.Text.Encoding.Unicode.GetBytes(xml))
            Dim transform As New System.Xml.Xsl.XslCompiledTransform()
            transform.Load(xslFile)

            Dim xpathDoc As New System.Xml.XPath.XPathDocument(memory)
            Dim sb As New System.Text.StringBuilder()
            Dim sw As New System.IO.StringWriter(sb)

            transform.Transform(xpathDoc, Nothing, sw)
            result = sb.ToString()
        End Using

        Return result

    End Function

    Public Sub BindControl(ByVal rv As ReportViewer, ByVal data As DataTable, ByVal name As String, ByVal virtualRldc As String)
        Dim rds As New ReportDataSource()

        rv.ProcessingMode = ProcessingMode.Local
        rds.Name = "DT" & name
        rds.Value = data

        Dim r As LocalReport = rv.LocalReport
        r.ReportPath = virtualRldc
        r.DataSources.Add(rds)
    End Sub

    Public Function BuildRDLC(ByVal data As DataTable, ByVal rpt As DataTable, ByVal name As String, ByVal table_name As String, _
                              ByVal gLanguage As String, ByVal title As String, _
                              ByVal Font_size As String, _
                              ByVal Page_Width As String, _
                              ByVal Page_Height As String, _
                              ByVal borderStyle As String, _
                              ByVal showHeader As String, _
                              ByVal showFooter As String) As String

        Dim virtualFolder As String = storingPath & "rdlc"
        Dim virtualXslt As String = virtualFolder & "/xslt"
        Dim virtualTemp As String = virtualFolder & "/temp"

        If Not IO.Directory.Exists(HttpContext.Current.Server.MapPath(virtualFolder)) Then IO.Directory.CreateDirectory(HttpContext.Current.Server.MapPath(virtualFolder))
        If Not IO.Directory.Exists(HttpContext.Current.Server.MapPath(virtualXslt)) Then IO.Directory.CreateDirectory(HttpContext.Current.Server.MapPath(virtualXslt))
        If Not IO.Directory.Exists(HttpContext.Current.Server.MapPath(virtualTemp)) Then IO.Directory.CreateDirectory(HttpContext.Current.Server.MapPath(virtualTemp))

        Dim virtualRdlc As String = virtualTemp & "/" & table_name & Format(nowTime, "hhmmssfff") & ".rdlc"
        Dim virtualSchema As String = virtualTemp & "/" & table_name & Format(nowTime, "hhmmssfff") & ".schema"

        virtualXslt = virtualXslt & "/rdlc.xsl"

        data.TableName = table_name

        physicalSchema = HttpContext.Current.Server.MapPath(virtualSchema)

        data.WriteXmlSchema(physicalSchema)

        Dim xmlDomSchema As New XmlDocument()
        xmlDomSchema.Load(physicalSchema)
        xmlDomSchema.DocumentElement.SetAttribute("Name", "DT" & name)

        Dim fieldname As String = "cold_label"

        Dim typeNode As XmlElement = xmlDomSchema.SelectSingleNode("*/*/*/*/*/*/*")
        Dim att_title As XmlAttribute
        If title = "" Then title = name & " Report"
        att_title = typeNode.OwnerDocument.CreateAttribute("title")
        att_title.Value = title

        typeNode.Attributes.Append(att_title)

        Dim tmpValue As String = ""

        Dim usedWidth As Double = 0
        Dim _pagewidth As Double = CDbl(Page_Width.Substring(0, Page_Width.Length - 2)) / 10

        _pagewidth = _pagewidth - 2

        Dim colGotWidthCnt As Integer = 0

        For i As Integer = 0 To rpt.Rows.Count - 1
            usedWidth += gU.decodeEmptyCdbl(rpt.Rows(i).Item("srcl_rpt_width").ToString.Trim, 0)
            If gU.decodeEmptyCdbl(rpt.Rows(i).Item("srcl_rpt_width").ToString.Trim, 0) > 0 Then
                colGotWidthCnt += 1
            End If
        Next

        Dim unuse_width As Double = _pagewidth - usedWidth
        Dim defColWidth As Double = unuse_width / (rpt.Rows.Count - colGotWidthCnt)

        Dim att_height As XmlAttribute = typeNode.OwnerDocument.CreateAttribute("pageheight")
        Dim att_pagewidth As XmlAttribute = typeNode.OwnerDocument.CreateAttribute("pagewidth")
        Dim att_fontsize As XmlAttribute = typeNode.OwnerDocument.CreateAttribute("fontsize")
        Dim att_linewidth As XmlAttribute = typeNode.OwnerDocument.CreateAttribute("linewidth")
        Dim att_userid As XmlAttribute = typeNode.OwnerDocument.CreateAttribute("userid")

        att_height.Value = Page_Height
        att_pagewidth.Value = Page_Width

        If Not Font_size.Contains("pt") Then
            Font_size += "pt"
        End If

        att_fontsize.Value = Font_size
        att_linewidth.Value = _pagewidth & "cm"
        att_userid.Value = HttpContext.Current.Session("username")

        typeNode.Attributes.Append(att_height)
        typeNode.Attributes.Append(att_pagewidth)
        typeNode.Attributes.Append(att_fontsize)
        typeNode.Attributes.Append(att_linewidth)
        typeNode.Attributes.Append(att_userid)

        Dim att_align As XmlAttribute
        Dim att_label As XmlAttribute
        Dim att_width As XmlAttribute
        Dim att_cellMode As XmlAttribute
        Dim att_label2 As XmlAttribute

        For x As Integer = 0 To rpt.Rows.Count - 1
            att_align = typeNode.OwnerDocument.CreateAttribute("align")
            att_label = typeNode.OwnerDocument.CreateAttribute("label")
            att_width = typeNode.OwnerDocument.CreateAttribute("width")

            att_cellMode = typeNode.OwnerDocument.CreateAttribute("cellMode")
            att_label2 = typeNode.OwnerDocument.CreateAttribute("label2")


            If typeNode.ChildNodes(x).Attributes.Item(0).Value.ToString.ToUpper = (rpt.Rows(x).Item("table_name").ToString.Trim & "_" & rpt.Rows(x).Item("column_name").ToString.Trim).ToString.ToUpper Then
                Select Case gU.decodeNullOrEmpty(rpt.Rows(x).Item("cold_align").ToString, "L")
                    Case "L"
                        tmpValue = "Left"
                    Case "R"
                        tmpValue = "Right"
                    Case "C"
                        tmpValue = "Center"
                    Case Else
                        tmpValue = "Left"
                End Select

                att_align.Value = tmpValue
                att_label.Value = gU.decodeNullOrEmpty(rpt.Rows(x).Item(fieldname).ToString.Trim, "")
                att_width.Value = gU.decodeNullOrEmpty(rpt.Rows(x).Item("cold_width").ToString.Trim, defColWidth) & "cm"

                typeNode.ChildNodes(x).Attributes.Append(att_align)
                typeNode.ChildNodes(x).Attributes.Append(att_label)
                typeNode.ChildNodes(x).Attributes.Append(att_label2)
                typeNode.ChildNodes(x).Attributes.Append(att_width)
                typeNode.ChildNodes(x).Attributes.Append(att_cellMode)
            End If
        Next

        xmlDomSchema.ImportNode(typeNode, True)
        xmlDomSchema.Save(physicalSchema)

        physicalXslt = HttpContext.Current.Server.MapPath(virtualXslt)
        Dim xml As String = TransformXml(xmlDomSchema.OuterXml, physicalXslt)
        physicalRdlc = HttpContext.Current.Server.MapPath(virtualRdlc)
        Dim xmlDomRdlc As New XmlDocument()

        xmlDomRdlc.LoadXml(xml)
        xmlDomRdlc.Save(physicalRdlc)

        Return virtualRdlc

    End Function

    Public Function BuildRDLC_SearchCol(ByVal data As DataTable, ByVal rpt As DataTable, ByVal name As String, ByVal table_name As String, _
                                        ByVal gLanguage As String, ByVal title As String, _
                                        ByVal Font_size As String, _
                                        ByVal Page_Width As String, _
                                        ByVal Page_Height As String, _
                                        ByVal borderStyle As String, _
                                        ByVal showHeader As String, _
                                        ByVal showFooter As String) As String

        Dim virtualFolder As String = storingPath & "rdlc"
        Dim virtualXslt As String = virtualFolder & "/xslt"
        Dim virtualTemp As String = virtualFolder & "/temp"

        If Not IO.Directory.Exists(HttpContext.Current.Server.MapPath(virtualFolder)) Then IO.Directory.CreateDirectory(HttpContext.Current.Server.MapPath(virtualFolder))
        If Not IO.Directory.Exists(HttpContext.Current.Server.MapPath(virtualXslt)) Then IO.Directory.CreateDirectory(HttpContext.Current.Server.MapPath(virtualXslt))
        If Not IO.Directory.Exists(HttpContext.Current.Server.MapPath(virtualTemp)) Then IO.Directory.CreateDirectory(HttpContext.Current.Server.MapPath(virtualTemp))

        Dim virtualRdlc As String = virtualTemp & "/" & table_name & Format(nowTime, "hhmmssfff") & ".rdlc"
        Dim virtualSchema As String = virtualTemp & "/" & table_name & Format(nowTime, "hhmmssfff") & ".schema"

        virtualXslt = virtualXslt & "/rdlc.xsl"

        data.TableName = table_name

        physicalSchema = HttpContext.Current.Server.MapPath(virtualSchema)

        data.WriteXmlSchema(physicalSchema)

        Dim xmlDomSchema As New XmlDocument()
        xmlDomSchema.Load(physicalSchema)
        xmlDomSchema.DocumentElement.SetAttribute("Name", "DT" & name)

        Dim typeNode As XmlElement = xmlDomSchema.SelectSingleNode("*/*/*/*/*/*/*")
        Dim att_title As XmlAttribute
        If title = "" Then title = name & " Report"
        att_title = typeNode.OwnerDocument.CreateAttribute("title")
        att_title.Value = title

        typeNode.Attributes.Append(att_title)

        Dim att_comp As XmlAttribute
        att_comp = typeNode.OwnerDocument.CreateAttribute("company")

        If HttpContext.Current.Session("comp_name") IsNot Nothing Then
            att_comp.Value = HttpContext.Current.Session("comp_name").ToString
        End If

        typeNode.Attributes.Append(att_comp)

        Dim att_border As XmlAttribute
        att_border = typeNode.OwnerDocument.CreateAttribute("border")

        att_border.Value = borderStyle

        typeNode.Attributes.Append(att_border)


        Dim att_sh As XmlAttribute
        att_sh = typeNode.OwnerDocument.CreateAttribute("showheader")

        att_sh.Value = If(showHeader = "Y", "true", "false")

        typeNode.Attributes.Append(att_sh)

        Dim att_sf As XmlAttribute
        att_sf = typeNode.OwnerDocument.CreateAttribute("showfooter")

        att_sf.Value = If(showFooter = "Y", "true", "false")

        typeNode.Attributes.Append(att_sf)

        Dim usedWidth As Double = 0
        Dim _pagewidth As Double = CDbl(Page_Width.Substring(0, Page_Width.Length - 2)) / 10

        _pagewidth = _pagewidth - 2

        Dim colGotWidthCnt As Integer = 0

        For i As Integer = 0 To rpt.Rows.Count - 1
            usedWidth += gU.decodeEmptyCdbl(rpt.Rows(i).Item("srcl_rpt_width").ToString.Trim, 0)
            If gU.decodeEmptyCdbl(rpt.Rows(i).Item("srcl_rpt_width").ToString.Trim, 0) > 0 Then
                colGotWidthCnt += 1
            End If
        Next

        Dim unuse_width As Double = _pagewidth - usedWidth
        Dim defColWidth As Double = unuse_width / (rpt.Rows.Count - colGotWidthCnt)

        Dim att_height As XmlAttribute = typeNode.OwnerDocument.CreateAttribute("pageheight")
        Dim att_pagewidth As XmlAttribute = typeNode.OwnerDocument.CreateAttribute("pagewidth")
        Dim att_fontsize As XmlAttribute = typeNode.OwnerDocument.CreateAttribute("fontsize")
        Dim att_linewidth As XmlAttribute = typeNode.OwnerDocument.CreateAttribute("linewidth")
        Dim att_userid As XmlAttribute = typeNode.OwnerDocument.CreateAttribute("userid")

        att_height.Value = Page_Height
        att_pagewidth.Value = Page_Width

        If Not Font_size.Contains("pt") Then
            Font_size += "pt"
        End If

        att_fontsize.Value = Font_size
        att_linewidth.Value = _pagewidth & "cm"
        att_userid.Value = HttpContext.Current.Session("usr_id")

        typeNode.Attributes.Append(att_height)
        typeNode.Attributes.Append(att_pagewidth)
        typeNode.Attributes.Append(att_fontsize)
        typeNode.Attributes.Append(att_linewidth)
        typeNode.Attributes.Append(att_userid)

        Dim att_align As XmlAttribute
        Dim att_label As XmlAttribute
        Dim att_width As XmlAttribute

        Dim att_hd_width As XmlAttribute
        Dim att_hd_align As XmlAttribute

        Dim att_hd_above_height As XmlAttribute
        Dim att_hd_below_height As XmlAttribute

        Dim att_label2 As XmlAttribute
        Dim att_cellMode As XmlAttribute
        Dim att_colspan As XmlAttribute

        Dim att_dateformat As XmlAttribute

        Dim mergeTillIdx As Integer = -1
        Dim colspan As Integer = 0

        For x As Integer = 0 To rpt.Rows.Count - 1
            att_align = typeNode.OwnerDocument.CreateAttribute("align")
            att_label = typeNode.OwnerDocument.CreateAttribute("label")
            att_width = typeNode.OwnerDocument.CreateAttribute("width")

            att_hd_width = typeNode.OwnerDocument.CreateAttribute("hd_width")
            att_hd_align = typeNode.OwnerDocument.CreateAttribute("hd_align")

            att_hd_above_height = typeNode.OwnerDocument.CreateAttribute("hd_above_height")
            att_hd_below_height = typeNode.OwnerDocument.CreateAttribute("hd_below_height")

            att_label2 = typeNode.OwnerDocument.CreateAttribute("label2")

            att_colspan = typeNode.OwnerDocument.CreateAttribute("colspan")
            att_cellMode = typeNode.OwnerDocument.CreateAttribute("cellMode")

            att_dateformat = typeNode.OwnerDocument.CreateAttribute("dateformat")

            For i As Integer = 0 To typeNode.ChildNodes.Count - 1
                Try
                    Dim isMatched As Boolean = False

                    If rpt.Rows(x).Item("cold_tabcol").ToString.Trim.Contains(".") Then
                        If typeNode.ChildNodes(i).Attributes.Item(0).Value.ToString.ToUpper = (rpt.Rows(x).Item("cold_tabcol").ToString.Trim.Split(".")(1)).ToString.ToUpper Then
                            isMatched = True
                        End If
                    Else
                        If typeNode.ChildNodes(i).Attributes.Item(0).Value.ToString.ToUpper = rpt.Rows(x).Item("cold_tabcol").ToString.Trim.ToUpper Then
                            isMatched = True
                        End If
                    End If

                    If isMatched Then
                        Select Case gU.decodeNullOrEmpty(rpt.Rows(x).Item("srcl_rpt_align").ToString, "L")
                            Case "L"
                                att_align.Value = "Left"
                            Case "R"
                                att_align.Value = "Right"
                            Case "C"
                                att_align.Value = "Center"
                            Case Else
                                att_align.Value = "Left"
                        End Select

                        Select Case "L"
                            Case "L"
                                att_hd_align.Value = "Left"
                            Case "R"
                                att_hd_align.Value = "Right"
                            Case "C"
                                att_hd_align.Value = "Center"
                            Case Else
                                att_hd_align.Value = "Left"
                        End Select

                        att_dateformat.Value = HttpContext.Current.Cache("DDFORMAT2")

                        att_label.Value = rpt.Rows(x).Item("srcl_label").ToString.Trim
                        att_width.Value = gU.decodeNullOrEmpty(rpt.Rows(x).Item("srcl_rpt_width").ToString.Trim, defColWidth) & "cm"
                        att_hd_width.Value = att_width.Value

                        att_hd_above_height.Value = "0.508cm"
                        att_hd_below_height.Value = "0.508cm"

                        att_label2.Value = ""

                        If att_label2.Value <> "" Then
                            att_cellMode.Value = "MergeDown"
                        End If

                        If mergeTillIdx > 0 And i <= mergeTillIdx - 1 Then
                            att_cellMode.Value = "MergePrevious"

                            If typeNode.ChildNodes(mergeTillIdx - colspan).Attributes("label2").Value.ToString <> "" Then
                                Dim mergeEle As XmlElement = typeNode.ChildNodes(mergeTillIdx - colspan).SelectSingleNode("MergeColumns")
                                Dim colsEle As XmlElement = xmlDomSchema.CreateElement("Columns")
                                Dim att_name As XmlAttribute = colsEle.OwnerDocument.CreateAttribute("name")
                                Dim att_align_hd As XmlAttribute = colsEle.OwnerDocument.CreateAttribute("align")
                                Dim att_label_hd As XmlAttribute = colsEle.OwnerDocument.CreateAttribute("label")
                                Dim att_width_hd As XmlAttribute = colsEle.OwnerDocument.CreateAttribute("width")
                                Dim att_left As XmlAttribute = colsEle.OwnerDocument.CreateAttribute("left")

                                Dim att_height_hd As XmlAttribute = colsEle.OwnerDocument.CreateAttribute("height")

                                att_name.Value = typeNode.ChildNodes(i).Attributes.Item(0).Value.ToString.ToUpper
                                att_align_hd.Value = att_hd_align.Value
                                att_label_hd.Value = att_label.Value
                                att_width_hd.Value = att_width.Value
                                att_height_hd.Value = "0.508cm"

                                Dim currentLeft As Double = 0

                                For j As Integer = 0 To mergeEle.ChildNodes.Count - 1
                                    currentLeft += CDbl(mergeEle.ChildNodes(j).Attributes("width").Value.ToString().Replace("cm", ""))
                                    'currentLeft += 0.0123
                                Next

                                typeNode.ChildNodes(mergeTillIdx - colspan).Attributes("hd_width").Value = currentLeft + CDbl(att_width.Value.ToString.Replace("cm", "")) & "cm"

                                att_left.Value = currentLeft & "cm"

                                colsEle.Attributes.Append(att_name)
                                colsEle.Attributes.Append(att_align_hd)
                                colsEle.Attributes.Append(att_label_hd)
                                colsEle.Attributes.Append(att_width_hd)
                                colsEle.Attributes.Append(att_left)

                                colsEle.Attributes.Append(att_height_hd)

                                mergeEle.AppendChild(colsEle)
                            End If

                            If i = mergeTillIdx Then
                                mergeTillIdx = -1
                                colspan = 0
                            End If
                        End If

                        typeNode.ChildNodes(x).Attributes.Append(att_align)
                        typeNode.ChildNodes(x).Attributes.Append(att_label)
                        typeNode.ChildNodes(x).Attributes.Append(att_label2)
                        typeNode.ChildNodes(x).Attributes.Append(att_width)

                        typeNode.ChildNodes(x).Attributes.Append(att_hd_width)
                        typeNode.ChildNodes(x).Attributes.Append(att_hd_align)

                        typeNode.ChildNodes(x).Attributes.Append(att_hd_above_height)
                        typeNode.ChildNodes(x).Attributes.Append(att_hd_below_height)

                        typeNode.ChildNodes(x).Attributes.Append(att_colspan)

                        typeNode.ChildNodes(x).Attributes.Append(att_cellMode)

                        typeNode.ChildNodes(x).Attributes.Append(att_dateformat)

                        Exit For
                    End If
                Catch ex As Exception

                End Try
            Next
        Next

        xmlDomSchema.ImportNode(typeNode, True)
        xmlDomSchema.Save(physicalSchema)

        physicalXslt = HttpContext.Current.Server.MapPath(virtualXslt)
        Dim xml As String = TransformXml(xmlDomSchema.OuterXml, physicalXslt)
        physicalRdlc = HttpContext.Current.Server.MapPath(virtualRdlc)
        Dim xmlDomRdlc As New XmlDocument()

        xmlDomRdlc.LoadXml(xml)
        xmlDomRdlc.Save(physicalRdlc)

        Return virtualRdlc

    End Function


    Public Sub deleteTempFiles(ByVal pRdlc As String, ByVal pSchema As String)
        Dim physicalTemp As String = HttpContext.Current.Server.MapPath(storingPath & "rdlc/temp")
        If File.Exists(pRdlc) Then
            File.Delete(pRdlc)
        End If

        If File.Exists(pSchema) Then
            File.Delete(pSchema)
        End If
    End Sub


    Public Function GetCode39(ByVal strSource As String) As Bitmap
        Dim x As Integer = 5    '左邊界
        Dim y As Integer = 0    '上邊界 
        Dim WidLength As Integer = 2    '粗BarCode長度
        Dim NarrowLength As Integer = 1 '細BarCode長度
        Dim BarCodeHeight As Integer = 24 'BarCode高度
        Dim intSourceLength As Integer = strSource.Length
        Dim strEncode As String = "010010100"   '編碼字串 初值為 起始符號 *
        Dim AlphaBet As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%*" 'Code39的字母

        'Code39的各字母對應碼
        Dim Code39 As String() = {"000110100", "100100001", "001100001", "101100000", "000110001", "100110000", "001110000", "000100101", _
                                  "100100100", "001100100", "100001001", "001001001", "101001000", "000011001", "100011000", "001011000", _
                                  "000001101", "100001100", "001001100", "000011100", "100000011", "001000011", "101000010", "000010011", _
                                  "100010010", "001010010", "000000111", "100000110", "001000110", "000010110", "110000001", "011000001", _
                                  "111000000", "010010001", "110010000", "011010000", "010000101", "110000100", "011000100", "010101000", _
                                  "010100010", "010001010", "000101010", "010010100"}

        strSource = strSource.ToUpper()

        '實作圖片 
        Dim objBitmap As Bitmap = New Bitmap(((WidLength * 3 + NarrowLength * 7) * (intSourceLength + 2)) + (x * 2), BarCodeHeight + (y * 2))

        Dim objGraphics As Graphics = Graphics.FromImage(objBitmap)  '宣告GDI+繪圖介面

        '填上底色

        objGraphics.FillRectangle(Brushes.White, 0, 0, objBitmap.Width, objBitmap.Height)

        For i As Integer = 0 To intSourceLength - 1
            '檢查是否有非法字元
            If (AlphaBet.IndexOf(strSource(i)) = -1 Or strSource(i) = "*") Then
                objGraphics.DrawString("含有非法字元", SystemFonts.DefaultFont, Brushes.Red, x, y)
                Return objBitmap
            End If

            '查表編碼

            strEncode = String.Format("{0}0{1}", strEncode, Code39(AlphaBet.IndexOf(strSource(i))))
        Next

        strEncode = String.Format("{0}0010010100", strEncode)   '補上結束符號 *       

        Dim intEncodeLength As Integer = strEncode.Length   '編碼後長度

        Dim intBarWidth As Integer

        For i As Integer = 0 To intEncodeLength - 1 '依碼畫出Code39 BarCode
            If strEncode(i) = "1" Then
                intBarWidth = WidLength
            Else
                intBarWidth = NarrowLength
            End If

            If i Mod 2 = 0 Then
                objGraphics.FillRectangle(Brushes.Black, x, y, intBarWidth, BarCodeHeight)
            Else
                objGraphics.FillRectangle(Brushes.White, x, y, intBarWidth, BarCodeHeight)
            End If
            x += intBarWidth
        Next

        Return objBitmap
    End Function

End Class
