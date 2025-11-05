
Partial Class CustomControls_YearRange
    Inherits System.Web.UI.UserControl
    Private startYr As Integer = 2012
    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        Dim currentYr As Integer = CInt(Now.Date.Year)
        Dim newListItem As ListItem
        If Not IsPostBack Then
            YrRange.Items.Clear()
            For i = startYr To currentYr

                newListItem = New ListItem

                newListItem.Value = i & "-" & i + 1
                newListItem.Text = i & " - " & i + 1

                YrRange.Items.Add(newListItem)
            Next

            YrRange.SelectedIndex = YrRange.Items.Count - 1

            Dim jsStr As String = ""

            jsStr = "document.getElementById('YrRange').value = document.getElementById('" & YrRange.ClientID & "').value;"

            Dim clientScript As ClientScriptManager = Me.Page.ClientScript

            clientScript.RegisterStartupScript(Me.GetType(), "", jsStr, True)

            YrRange.Attributes.Add("onchange", jsStr)
        End If
        

    End Sub
End Class
