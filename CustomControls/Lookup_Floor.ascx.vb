
Partial Class CustomControls_WebUserControl
    Inherits System.Web.UI.UserControl
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils

    Private BehaviorID As String = ""

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            BehaviorID = btnSearch.UniqueID
            FLOOR_ModalPopupExtender.BehaviorID = BehaviorID


        End If
        
    End Sub

    Protected Sub TEXTBOXUDP_Unload(sender As Object, e As System.EventArgs) Handles TEXTBOXUDP.Unload
        Dim methodInfo As System.Reflection.MethodInfo = GetType(ScriptManager).GetMethods(System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance).Where(Function(i) i.Name.Equals("System.Web.UI.IScriptManagerInternal.RegisterUpdatePanel")).First()
        methodInfo.Invoke(ScriptManager.GetCurrent(Page), New Object() {TryCast(sender, UpdatePanel)})

    End Sub

    Protected Sub btnSearch_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles btnSearch.Click

        Dim addSQL As String = ""
        Dim noSelect As Boolean = False

        If Not Session("CK_WH") Is Nothing AndAlso Session("CK_WH") <> "" Then
            addSQL = " AND WH_MAIN_WH='" & gU.dbEncode(Session("CK_WH")) & "' "
            noSelect = True
        End If

        uiFun.load_dropdown(WH_CODE, "Select wh_code, wh_name from wms_warehouse where imp_code='" & gU.dbEncode(Session("imp_code")) & "'" & addSQL & " order by wh_name", "wh_code", "wh_name", , , , noSelect)



        Dim newListItem As ListItem

        FL_NUM.Items.Clear()

        newListItem = New ListItem
        newListItem.Value = ""
        newListItem.Text = "SELECT"
        FL_NUM.Items.Add(newListItem)

        If Not Session("CK_WH") Is Nothing AndAlso Session("CK_WH") <> "" Then
            'WH_CODE.Enabled = False
            uiFun.load_dropdown(FL_NUM, "Select fl_num, fl_name from wms_wh_fl where imp_code='" & gU.dbEncode(Session("imp_code")) & "'  AND WH_CODE='" & WH_CODE.SelectedValue & "' order by fl_name", "fl_num", "fl_name")
        End If

        btnOK1.Enabled = False

        FLOOR_ModalPopupExtender.Show()

    End Sub

    Protected Sub btnOK1_Click(sender As Object, e As System.EventArgs) Handles btnOK1.Click
        dsp_FLOOR_CODE.Text = WH_CODE.SelectedItem.Text & " - " & FL_NUM.SelectedItem.Text
        FLOOR_CODE.Value = WH_CODE.SelectedValue & "" & FL_NUM.SelectedValue.ToString.PadLeft(2, "0")

        Dim jsStr As String = ""
        jsStr = "document.getElementById('FLOOR_CODE').value = document.getElementById('" & FLOOR_CODE.ClientID & "').value;"

        ScriptManager.RegisterStartupScript(TEXTBOXUDP, GetType(UpdatePanel), "closeModal", jsStr, True)

    End Sub

    Protected Sub WH_CODE_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles WH_CODE.SelectedIndexChanged
        Dim newListItem As ListItem

        If WH_CODE.SelectedValue <> "" Then
            uiFun.load_dropdown(FL_NUM, "Select fl_num, fl_name from wms_wh_fl where imp_code='" & gU.dbEncode(Session("imp_code")) & "' and wh_code='" & gU.dbEncode(WH_CODE.SelectedValue) & "' order by fl_name", "fl_num", "fl_name")
        Else
            FL_NUM.Items.Clear()
            newListItem = New ListItem
            newListItem.Value = ""
            newListItem.Text = "SELECT"
            FL_NUM.Items.Add(newListItem)
        End If

        FLOOR_ModalPopupExtender.Show()
    End Sub

    Protected Sub FL_NUM_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles FL_NUM.SelectedIndexChanged
        If FL_NUM.SelectedValue <> "" Then
            btnOK1.Enabled = True
        End If

        FLOOR_ModalPopupExtender.Show()
    End Sub

End Class
