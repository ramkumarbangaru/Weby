Imports System.IO

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
            txttag.Text = objcurelement.tagName
            If Not (IsNothing(objcurelement.getAttribute("name")) Or IsDBNull(objcurelement.getAttribute("name"))) Then
                txtname.Text = objcurelement.getAttribute("name")
            Else
                txtname.Text = ""
            End If

            txtclass.Text = objcurelement.className
            txtxpathrelative.Text = objop.getXpath(objcurelement, False)
            txtxpathabsolute.Text = objop.getXpath(objcurelement, True)
            txtcsspath.Text = objop.getCss(objcurelement)
            txtcsssubpath.Text = "css=" + objop.getCssSubPath(objcurelement)
        End If
    End Sub
    Private Delegate Sub addItemtoTreeDelegate(strvar As String, strid As String, strname As String, strtagname As String, strclass As String, strxpathrelative As String, strxpathabsolute As String, strcsspath As String, strcsssubpath As String)
    Public Sub addItemtoTree(strvar As String, strid As String, strname As String, strtagname As String, strclass As String, strxpathrelative As String, strxpathabsolute As String, strcsspath As String, strcsssubpath As String)
        If Me.InvokeRequired Then
            Me.Invoke(New addItemtoTreeDelegate(AddressOf addItemtoTree), strvar, strid, strname, strtagname, strclass, strxpathrelative, strxpathabsolute, strcsspath, strcsssubpath)
        Else
            Dim currnode As TreeNode = treeobjectmap.Nodes.Add(strvar)
            currnode.Nodes.Add("ID : " & strid)
            currnode.Nodes.Add("Name : " & strname)
            currnode.Nodes.Add("Tag Name : " & strtagname)
            currnode.Nodes.Add("Class Name : " & strclass)
            currnode.Nodes.Add("XPATH Relative : " & strxpathrelative)
            currnode.Nodes.Add("XAPTH Absolute : " & strxpathabsolute)
            currnode.Nodes.Add("CSS Path : " & strcsspath)
            currnode.Nodes.Add("CSS Sub path : " & strcsssubpath)
        End If
    End Sub

    Private Sub treeobjectmap_NodeMouseClick(ByVal sender As Object,
    ByVal e As TreeNodeMouseClickEventArgs) _
    Handles treeobjectmap.NodeMouseClick
        If (IsNothing(e.Node.Parent)) Then
            txtid.Text = e.Node.Nodes(0).Text.Split(":")(1).Trim()
            txtname.Text = e.Node.Nodes(1).Text.Split(":")(1).Trim()
            txttag.Text = e.Node.Nodes(2).Text.Split(":")(1).Trim()
            txtclass.Text = e.Node.Nodes(3).Text.Split(":")(1).Trim()
            txtxpathrelative.Text = e.Node.Nodes(4).Text.Split(":")(1).Trim()
            txtxpathabsolute.Text = e.Node.Nodes(5).Text.Split(":")(1).Trim()
            txtcsspath.Text = e.Node.Nodes(6).Text.Split(":")(1).Trim()
            txtcsssubpath.Text = e.Node.Nodes(7).Text.Split(":")(1).Trim()
        End If
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

    Private Sub CSVToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CSVToolStripMenuItem.Click
        Dim saveFileDialog1 As New SaveFileDialog()
        Dim webobject As TreeNode
        saveFileDialog1.Filter = "Web Object File | *.csv"
        saveFileDialog1.Title = "Save the object file"
        saveFileDialog1.ShowDialog()
        If saveFileDialog1.FileName <> "" Then
            Using sw As StreamWriter = File.CreateText(saveFileDialog1.FileName)
                sw.WriteLine("ObjectName,ID,Name,Tag,Class,Xpath - Relative,Xapth - Absolute,CSS, CSS SubPath")
                For Each webobject In treeobjectmap.Nodes
                    sw.WriteLine(webobject.Text + "," + webobject.Nodes(0).Text.Split(":")(1).Trim() + "," + webobject.Nodes(1).Text.Split(":")(1).Trim() + "," + webobject.Nodes(2).Text.Split(":")(1).Trim() + "," + webobject.Nodes(3).Text.Split(":")(1).Trim() + "," + webobject.Nodes(4).Text.Split(":")(1).Trim() + "," + webobject.Nodes(5).Text.Split(":")(1).Trim() + "," + webobject.Nodes(6).Text.Split(":")(1).Trim() + "," + webobject.Nodes(7).Text.Split(":")(1).Trim())
                Next
            End Using
        End If
    End Sub

    Private Sub JavaPageFactoryObjectToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles JavaPageFactoryObjectToolStripMenuItem.Click
        Dim saveFileDialog1 As New SaveFileDialog()
        Dim webobject As TreeNode
        saveFileDialog1.Filter = "Web Object File | *.java"
        saveFileDialog1.Title = "Save the object file"
        saveFileDialog1.ShowDialog()
        If saveFileDialog1.FileName <> "" Then
            Using sw As StreamWriter = File.CreateText(saveFileDialog1.FileName)
                sw.WriteLine("import org.openqa.selenium.WebDriver;")
                sw.WriteLine("import org.openqa.selenium.WebElement;")
                sw.WriteLine("import org.openqa.selenium.support.FindBy;")
                sw.WriteLine("import org.openqa.selenium.support.PageFactory;")
                sw.WriteLine(vbNewLine & "public class WebObjects {")
                sw.WriteLine(vbNewLine & vbTab & "/**")
                sw.WriteLine(vbTab & "* All WebElements are identified by @FindBy annotation")
                sw.WriteLine(vbTab & "*/")
                sw.WriteLine(vbNewLine & vbTab & "WebDriver driver;")
                For Each webobject In treeobjectmap.Nodes
                    If webobject.Nodes(0).Text.Split(":")(1).Trim() <> "" Then
                        sw.WriteLine(vbNewLine & vbTab & "@FindBy(id=" & """" & webobject.Nodes(0).Text.Split(":")(1).Trim() & """" & ")")
                    Else
                        sw.WriteLine(vbNewLine & vbTab & "@FindBy(xpath=" & """" & webobject.Nodes(4).Text.Split(":")(1).Trim() & """" & ")")
                    End If
                    sw.WriteLine(vbTab & "WebElement " & webobject.Text & ";")
                Next
                sw.WriteLine(vbNewLine & "}")
            End Using
        End If
    End Sub
End Class
