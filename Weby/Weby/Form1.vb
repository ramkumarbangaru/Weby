Imports System.IO
Public Class frmweby
    Dim IE As IEBrowser
    Dim FlagSaved As Boolean
    Private Sub frmweby_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        IE = New IEBrowser(Me)
        treeobjectmap.ContextMenuStrip = ContextMenuStrip2
        For Each RootNode As TreeNode In treeobjectmap.Nodes
            RootNode.ContextMenuStrip = ContextMenuStrip1
        Next
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
            currnode.ContextMenuStrip = ContextMenuStrip1
        End If
        FlagSaved = False
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
            FlagSaved = True
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
            FlagSaved = True
        End If
    End Sub

    Private Sub exitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles exitToolStripMenuItem.Click
        If treeobjectmap.Nodes.Count = 0 Then
            FlagSaved = True
        End If
        If FlagSaved = False Then
            If MsgBox("Do you want to save unsaved changes?", vbYesNo, "Weby") = MsgBoxResult.No Then
                Application.Exit()
            End If
        Else
            Application.Exit()
        End If
    End Sub
    Private Sub frmby_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If treeobjectmap.Nodes.Count = 0 Then
            FlagSaved = True
        End If
        If FlagSaved = False Then
            If MsgBox("Do you want to save unsaved changes?", vbYesNo, "Weby") = MsgBoxResult.Yes Then
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub newToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles newToolStripMenuItem.Click
        If FlagSaved = False Then
            If MsgBox("Do you want to save unsaved changes?", vbYesNo, "Weby") = MsgBoxResult.No Then
                treeobjectmap.Nodes.Clear()
                FlagSaved = True
            End If
        Else
            treeobjectmap.Nodes.Clear()
            FlagSaved = True
        End If
    End Sub

    Private Sub openToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles openToolStripMenuItem.Click
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Web Object File | *csv;*.java;*.py"
        openFileDialog.Title = "Select the object file to open"
        openFileDialog.ShowDialog()
        Dim strFilename = openFileDialog.FileName
        If strFilename <> "" Then
            Dim strFileExtension = strFilename.Substring(InStrRev(strFilename, "."))
            Dim strcurrentline
            Using reader As StreamReader = New StreamReader(openFileDialog.FileName)
                strcurrentline = reader.ReadLine()
                Do While (Not strcurrentline Is Nothing)
                    If strFileExtension.ToLower() = "csv" Then
                        If Not (strcurrentline.ToString().Contains("ObjectName") And strcurrentline.ToString().Contains("Xpath")) Then 'First line  check
                            Dim flds
                            flds = strcurrentline.ToString().Split(",")
                            addItemtoTree(flds(0), flds(1), flds(2), flds(3), flds(4), flds(5), flds(6), flds(7), flds(8))
                        End If
                    ElseIf strFileExtension.ToLower() = "java" Then
                        If strcurrentline.ToString().Contains("FindBy(") Then
                            Dim flds = strcurrentline.ToString().Split("(")(1).Split(")")(0)
                            Dim locatortype = flds.Substring(0, InStr(flds, "=") - 1)
                            Dim lovatorvalue = flds.Substring(InStr(flds, "="))
                            Dim strsubline = reader.ReadLine()
                            While Not (strsubline.ToString().Contains("WebElement") Or strsubline Is Nothing)
                                strsubline = reader.ReadLine()
                            End While
                            If strsubline Is Nothing Then
                                Exit Do
                            End If
                            If strsubline.ToString().Contains("WebElement") Then
                                If locatortype.Trim().ToLower() = "id" Then
                                    addItemtoTree(strsubline.ToString().Split(" ")(1).Split(";")(0), lovatorvalue.Trim().Split("""")(1), "", "", "", "", "", "", "")
                                ElseIf locatortype.Trim().ToLower() = "xpath" Then
                                    addItemtoTree(strsubline.ToString().Split(" ")(1).Split(";")(0), "", "", "", "", lovatorvalue.Trim().Replace("""", ""), "", "", "")
                                End If
                            End If
                        End If
                    ElseIf strFileExtension.ToLower() = "py" Then
                        If strcurrentline.ToString().Contains("(By.") Then
                            Dim ObjName = strcurrentline.ToString().Split("=")(0).Trim()
                            Dim locatortype = strcurrentline.ToString().Split(".")(1).Split(",")(0).Trim()
                            Dim lovatorvalue = strcurrentline.ToString().Split(",")(1).Split(")")(0).Trim()
                            If locatortype.Trim().ToLower() = "id" Then
                                addItemtoTree(ObjName, lovatorvalue.Trim().Split("""")(1), "", "", "", "", "", "", "")
                            ElseIf locatortype.Trim().ToLower() = "xpath" Then
                                addItemtoTree(ObjName, "", "", "", "", lovatorvalue.Trim().Replace("""", ""), "", "", "")
                            End If
                        End If
                    End If
                    strcurrentline = reader.ReadLine()
                Loop
                FlagSaved = False
            End Using
        End If
    End Sub

    Private Sub PythonToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PythonToolStripMenuItem.Click
        Dim saveFileDialog1 As New SaveFileDialog()
        Dim webobject As TreeNode
        saveFileDialog1.Filter = "Web Object File | *.py"
        saveFileDialog1.Title = "Save the object file"
        saveFileDialog1.ShowDialog()
        If saveFileDialog1.FileName <> "" Then
            Using sw As StreamWriter = File.CreateText(saveFileDialog1.FileName)
                sw.WriteLine("from selenium.webdriver.common.by import By")
                sw.WriteLine(vbNewLine & "class WebObject(object):")
                sw.WriteLine(vbNewLine & vbTab & """""""")
                sw.WriteLine(vbTab & "Web Elements")
                sw.WriteLine(vbTab & """""""" & vbNewLine)
                For Each webobject In treeobjectmap.Nodes
                    If webobject.Nodes(0).Text.Split(":")(1).Trim() <> "" Then
                        sw.WriteLine(vbTab & webobject.Text & " = " & "(By.ID, """ & webobject.Nodes(0).Text.Split(":")(1).Trim() & """)")
                    Else
                        sw.WriteLine(vbTab & webobject.Text & " = " & "(By.XPATH, """ & webobject.Nodes(4).Text.Split(":")(1).Trim() & """)")
                    End If
                Next
                sw.WriteLine(vbNewLine)
            End Using
            FlagSaved = True
        End If
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem.Click
        If MsgBox("Do you want to delete this node?", vbYesNo, "Weby") = MsgBoxResult.Yes Then
            treeobjectmap.SelectedNode.Remove()
        End If
    End Sub

    Private Sub RenameToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RenameToolStripMenuItem.Click
        Dim frm = InputBox("Please Enter the new Name", "Rename", treeobjectmap.SelectedNode.Text)
        If frm.Trim() <> "" Then
            Dim SelectedNode As TreeNode = treeobjectmap.SelectedNode
            SelectedNode.Text = frm.Trim()
        End If
    End Sub

    Private Sub ClearAllToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles ClearAllToolStripMenuItem.Click
        If MsgBox("Do you want to clear all the nodes?", vbYesNo, "Weby") = MsgBoxResult.Yes Then
            treeobjectmap.Nodes.Clear()
        End If
    End Sub
End Class

