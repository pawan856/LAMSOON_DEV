
Partial Class CustomControls_WebUserControl
    Inherits System.Web.UI.UserControl
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils

    Private BehaviorID As String = ""

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            BehaviorID = btnSearch.UniqueID
            BIN_ModalPopupExtender.BehaviorID = BehaviorID
        End If

    End Sub

    Protected Sub TEXTBOXUDP_Unload(sender As Object, e As System.EventArgs) Handles TEXTBOXUDP.Unload
        Dim methodInfo As System.Reflection.MethodInfo = GetType(ScriptManager).GetMethods(System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance).Where(Function(i) i.Name.Equals("System.Web.UI.IScriptManagerInternal.RegisterUpdatePanel")).First()
        methodInfo.Invoke(ScriptManager.GetCurrent(Page), New Object() {TryCast(sender, UpdatePanel)})

    End Sub

    Protected Sub btnSearch_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles btnSearch.Click

        Dim addSQL As String = ""
        Dim noSelect As Boolean = False

        If Not Session("LOOKUP_MAIN_WH") Is Nothing AndAlso Session("LOOKUP_MAIN_WH") <> "" Then
            addSQL = " AND WH_MAIN_WH='" & gU.dbEncode(Session("LOOKUP_MAIN_WH")) & "' "
            'noSelect = True
        End If

        uiFun.load_dropdown(WH_CODE, "Select wh_code, wh_name from wms_warehouse where imp_code='" & gU.dbEncode(Session("imp_code")) & "' " & addSQL & " order by wh_name", "wh_code", "wh_name", , , , noSelect)



        Dim newListItem As ListItem

        FL_NUM.Items.Clear()
        AR_CODE.Items.Clear()
        RK_CODE.Items.Clear()
        BN_CODE.Items.Clear()

        newListItem = New ListItem
        newListItem.Value = ""
        newListItem.Text = "SELECT"
        FL_NUM.Items.Add(newListItem)

        newListItem = New ListItem
        newListItem.Value = ""
        newListItem.Text = "SELECT"
        AR_CODE.Items.Add(newListItem)

        newListItem = New ListItem
        newListItem.Value = ""
        newListItem.Text = "SELECT"
        RK_CODE.Items.Add(newListItem)

        newListItem = New ListItem
        newListItem.Value = ""
        newListItem.Text = "SELECT"
        BN_CODE.Items.Add(newListItem)

        If Not Session("LOOKUP_WH_CODE") Is Nothing AndAlso Session("LOOKUP_WH_CODE") <> "" Then
            'WH_CODE.Enabled = False
            uiFun.load_dropdown(FL_NUM, "Select fl_num, fl_name from wms_wh_fl where imp_code='" & gU.dbEncode(Session("imp_code")) & "' AND WH_CODE='" & WH_CODE.SelectedValue & "' order by fl_name", "fl_num", "fl_name")
        End If

        btnOK.Enabled = False

        BIN_ModalPopupExtender.Show()

    End Sub

    Protected Sub btnOK_Click(sender As Object, e As System.EventArgs) Handles btnOK.Click


        Dim returnCode As String = ""
        Dim returnText As String = ""

        'locCode = WH_CODE.Value & "" & FL_NUM.Value.ToString.PadLeft(2, "0") & "" & AR_CODE.Value.PadLeft(3, "0") & "" & RK_CODE.Value.ToString.PadLeft(4, "0") & "" & nDT.Rows(i).Item(i1).ToString.PadLeft(3, "0")

        returnText &= WH_CODE.SelectedItem.Text
        returnCode &= WH_CODE.SelectedValue

        If FL_NUM.SelectedValue <> "" Then
            returnText &= " - " & FL_NUM.SelectedItem.Text
            returnCode &= FL_NUM.SelectedValue.ToString.PadLeft(2, "0")

            If AR_CODE.SelectedValue <> "" Then
                returnText &= " - " & AR_CODE.SelectedItem.Text
                returnCode &= AR_CODE.SelectedValue.PadLeft(3, "0")
            End If

            If RK_CODE.SelectedValue <> "" Then
                returnText &= " - " & RK_CODE.SelectedItem.Text
                returnCode &= RK_CODE.SelectedValue.PadLeft(4, "0")
            End If

            If BN_CODE.SelectedValue <> "" Then
                returnText &= " - " & BN_CODE.SelectedItem.Text
                returnCode &= BN_CODE.SelectedValue.PadLeft(3, "0")
            End If
        End If

        'dsp_LOCATION_CODE.Text = WH_CODE.SelectedItem.Text & " - " & FL_NUM.SelectedItem.Text & " - " & AR_CODE.SelectedItem.Text
        'LOCATION_CODE.Value = WH_CODE.SelectedValue & "" & FL_NUM.SelectedValue.ToString.PadLeft(2, "0") & "" & AR_CODE.SelectedValue.PadLeft(3, "0")
        dsp_LOCATION_CODE.Text = returnText
        LOCATION_CODE.Value = returnCode

        Dim jsStr As String = ""
        jsStr = "document.getElementById('LOCATION_CODE').value = document.getElementById('" & LOCATION_CODE.ClientID & "').value;"

        ScriptManager.RegisterStartupScript(TEXTBOXUDP, GetType(UpdatePanel), "closeModal", jsStr, True)

    End Sub

    Protected Sub WH_CODE_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles WH_CODE.SelectedIndexChanged
        Dim newListItem As ListItem
        BIN_ModalPopupExtender.Show()

        If WH_CODE.SelectedValue <> "" Then
            uiFun.load_dropdown(FL_NUM, "Select fl_num, fl_name from wms_wh_fl where imp_code='" & gU.dbEncode(Session("imp_code")) & "' and wh_code='" & gU.dbEncode(WH_CODE.SelectedValue) & "' order by fl_name", "fl_num", "fl_name")
            btnOK.Enabled = True
        Else
            FL_NUM.Items.Clear()
            newListItem = New ListItem
            newListItem.Value = ""
            newListItem.Text = "SELECT"
            FL_NUM.Items.Add(newListItem)
        End If



        AR_CODE.Items.Clear()
        newListItem = New ListItem
        newListItem.Value = ""
        newListItem.Text = "SELECT"
        AR_CODE.Items.Add(newListItem)

        RK_CODE.Items.Clear()
        newListItem = New ListItem
        newListItem.Value = ""
        newListItem.Text = "SELECT"
        RK_CODE.Items.Add(newListItem)

        BN_CODE.Items.Clear()
        newListItem = New ListItem
        newListItem.Value = ""
        newListItem.Text = "SELECT"
        BN_CODE.Items.Add(newListItem)


    End Sub

    Protected Sub FL_NUM_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles FL_NUM.SelectedIndexChanged
        If FL_NUM.SelectedValue <> "" Then
            uiFun.load_dropdown(AR_CODE, "Select ar_code, ar_name from wms_wh_area where imp_code='" & gU.dbEncode(Session("imp_code")) & "' and wh_code='" & gU.dbEncode(WH_CODE.SelectedValue) & "' and fl_num='" & gU.dbEncode(FL_NUM.SelectedValue) & "' order by ar_name", "ar_code", "ar_name")
            btnOK.Enabled = True
        Else
            Dim newListItem As New ListItem

            AR_CODE.Items.Clear()
            newListItem.Value = ""
            newListItem.Text = "SELECT"
            AR_CODE.Items.Add(newListItem)
            btnOK.Enabled = False
        End If

        Dim nListItem As ListItem

        RK_CODE.Items.Clear()
        nListItem = New ListItem
        nListItem.Value = ""
        nListItem.Text = "SELECT"
        RK_CODE.Items.Add(nListItem)

        BN_CODE.Items.Clear()
        nListItem = New ListItem
        nListItem.Value = ""
        nListItem.Text = "SELECT"
        BN_CODE.Items.Add(nListItem)

        BIN_ModalPopupExtender.Show()
    End Sub

    Protected Sub AR_CODE_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles AR_CODE.SelectedIndexChanged
        Dim nListItem As ListItem

        If AR_CODE.SelectedValue <> "" Then
            uiFun.load_dropdown(RK_CODE, "Select rk_code, rk_name from wms_wh_rack where imp_code='" & gU.dbEncode(Session("imp_code")) & "' and wh_code='" & gU.dbEncode(WH_CODE.SelectedValue) & "' and fl_num='" & gU.dbEncode(FL_NUM.SelectedValue) & "' and ar_code='" & gU.dbEncode(AR_CODE.SelectedValue) & "' order by rk_name", "rk_code", "rk_name")

            BN_CODE.Items.Clear()
            nListItem = New ListItem
            nListItem.Value = ""
            nListItem.Text = "SELECT"
            BN_CODE.Items.Add(nListItem)
            btnOK.Enabled = True
        Else

            RK_CODE.Items.Clear()
            nListItem = New ListItem
            nListItem.Value = ""
            nListItem.Text = "SELECT"
            RK_CODE.Items.Add(nListItem)

            BN_CODE.Items.Clear()
            nListItem = New ListItem
            nListItem.Value = ""
            nListItem.Text = "SELECT"
            BN_CODE.Items.Add(nListItem)
        End If

        BIN_ModalPopupExtender.Show()
    End Sub

    Protected Sub RK_CODE_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles RK_CODE.SelectedIndexChanged
        Dim nListItem As ListItem

        If RK_CODE.SelectedValue <> "" Then
            uiFun.load_dropdown(BN_CODE, "Select bn_code, BN_CSMS_CODE from wms_wh_bin where imp_code='" & gU.dbEncode(Session("imp_code")) & "' and wh_code='" & gU.dbEncode(WH_CODE.SelectedValue) & "' and fl_num='" & gU.dbEncode(FL_NUM.SelectedValue) & "' and ar_code='" & gU.dbEncode(AR_CODE.SelectedValue) & "' and rk_code='" & gU.dbEncode(RK_CODE.SelectedValue) & "' order by BN_CSMS_CODE", "bn_code", "BN_CSMS_CODE")
            btnOK.Enabled = True
        Else
            BN_CODE.Items.Clear()
            nListItem = New ListItem
            nListItem.Value = ""
            nListItem.Text = "SELECT"
            BN_CODE.Items.Add(nListItem)
        End If

        BIN_ModalPopupExtender.Show()
    End Sub

    Protected Sub BN_CODE_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles BN_CODE.SelectedIndexChanged
        If BN_CODE.SelectedValue <> "" Then
            btnOK.Enabled = True
        End If

        BIN_ModalPopupExtender.Show()
    End Sub
End Class
