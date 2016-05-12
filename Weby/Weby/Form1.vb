Public Class frmweby
    Dim IE As IEBrowser
    Private Sub frmweby_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        IE = New IEBrowser(Me)
    End Sub
    Private Delegate Sub setValuesDelegate(objcurelement As mshtml.IHTMLElement)
    Public Sub setValues(objcurelement As mshtml.IHTMLElement)
        If Me.InvokeRequired Then
            Me.Invoke(New setValuesDelegate(AddressOf setValues), objcurelement)
        Else
            Dim objop As ObjectProperties
            objop = New ObjectProperties()
            txtid.Text = objcurelement.id
            txtid.Text = objcurelement.id
            txttag.Text = objcurelement.tagName
            txtname.Text = objop.getElementName(objcurelement)
            txtclass.Text = objcurelement.className
            txtxpathrelative.Text = objop.getXpath(objcurelement, False)
            txtxpathabsolute.Text = objop.getXpath(objcurelement, True)
            txtcsspath.Text = objop.getCss(objcurelement)
            txtcsssubpath.Text = "css=" + objop.getCssSubPath(objcurelement)
        End If
    End Sub
    Private Delegate Sub addItemtoTreeDelegate(strvar As String)
    Public Sub addItemtoTree(strvar As String)
        If Me.InvokeRequired Then
            Me.Invoke(New addItemtoTreeDelegate(AddressOf addItemtoTree), strvar)
        Else
            treeobjectmap.Nodes.Add(strvar)
        End If
    End Sub

    Private Sub treeobjectmap_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles treeobjectmap.AfterSelect

    End Sub

    Private Sub btnspy_Click(sender As Object, e As EventArgs) Handles btnspy.Click
        If btnspy.Text = "Spy" Then
            IE.AddHandlers()
            btnspy.Text = "Stop Spy"
        Else
            btnspy.Text = "Spy"
        End If
    End Sub

    Private Sub btnIEWin_Click(sender As Object, e As EventArgs) Handles btnIEWin.Click
        IE.GetIEWindow()
    End Sub
End Class
