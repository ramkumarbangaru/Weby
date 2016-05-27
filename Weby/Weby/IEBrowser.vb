Imports System.Text.RegularExpressions
Imports System.Threading
Imports mshtml
Public Class IEBrowser
    REM Declaring objects with Events
    Dim WithEvents objhtmldoc As HTMLDocument
    Dim WithEvents objframehtmldoc As HTMLDocument
    Dim WithEvents objIE As SHDocVw.InternetExplorer
    Dim WithEvents objShellWindows As New SHDocVw.ShellWindows

    Dim objprevelement As IHTMLElement
    Dim strprevelementproperty As String
    Dim objcurelement As IHTMLElement
    Dim strframename As String
    Dim strframeid As String
    Dim strframeuniqueid As String
    Dim objweby As frmweby
    Dim htmlhandle As Boolean
    Dim frmhtmlhandle As Boolean
    Dim strborder

    Public Sub New(ByVal caller As frmweby)
        GetIEWindow()
        objweby = frmweby
        htmlhandle = False
        frmhtmlhandle = False
    End Sub

    Public Sub GetIEWindow()
        If htmlhandle = True Then
            objhtmldoc = objIE.Document
            RemoveHandler CType(objhtmldoc, HTMLDocumentEvents2_Event).onclick, AddressOf Document_onclick
            RemoveHandler CType(objhtmldoc, HTMLDocumentEvents2_Event).onmouseover, AddressOf Document_onmouseover
            RemoveHandler CType(objhtmldoc, HTMLDocumentEvents2_Event).oncontextmenu, AddressOf Document_oncontextmenu
            RemoveHandler CType(objhtmldoc, HTMLDocumentEvents2_Event).onmouseout, AddressOf Document_onmouseout
        End If

        objIE = Nothing
        objIE = GetIE()
        If IsNothing(objIE) Then
            Application.Exit()
        Else
            htmlhandle = False
            frmhtmlhandle = False
        End If
    End Sub
    Public Sub AddHandlers()
        objhtmldoc = objIE.Document
        If htmlhandle = False Then
            AddHandler CType(objhtmldoc, HTMLDocumentEvents2_Event).onclick, AddressOf Document_onclick
            AddHandler CType(objhtmldoc, HTMLDocumentEvents2_Event).onmouseover, AddressOf Document_onmouseover
            AddHandler CType(objhtmldoc, HTMLDocumentEvents2_Event).oncontextmenu, AddressOf Document_oncontextmenu
            AddHandler CType(objhtmldoc, HTMLDocumentEvents2_Event).onmouseout, AddressOf Document_onmouseout
            htmlhandle = True
        End If
    End Sub

    Private Sub objIE_DocumentComplete(pDisp As Object, ByRef URL As Object) Handles objIE.DocumentComplete
        AddHandlers()
    End Sub
    Private Sub Document_onmouseover(ByVal e As mshtml.IHTMLEventObj)
        If objweby.btnspy.Text = "Stop Spy" Then
            objprevelement = Nothing
            objcurelement = e.srcElement
            objhtmldoc = objIE.Document
            If InStr(1, LCase(objcurelement.tagName), "iframe") >= 1 Then
                Call setiFrame(objcurelement)
            ElseIf InStr(1, LCase(objcurelement.tagName), "frame") >= 1 Then
                Call setFrame(objcurelement)
            Else
                objprevelement = objcurelement
                strborder = objcurelement.style.border
                objcurelement.style.border = "solid 1px #ff0000"
                'objcurelement.style.setAttribute("border", "solid 1px #ff0000")
            End If

            Call objweby.setValues(objcurelement)
        End If
    End Sub

    Private Sub setiFrame(objcurelement As IHTMLElement)
        objframehtmldoc = objcurelement.document
        If frmhtmlhandle = False Then
            objhtmldoc = objIE.Document
            RemoveHandler CType(objhtmldoc, HTMLDocumentEvents2_Event).onclick, AddressOf Document_onclick
            RemoveHandler CType(objhtmldoc, HTMLDocumentEvents2_Event).onmouseover, AddressOf Document_onmouseover
            RemoveHandler CType(objhtmldoc, HTMLDocumentEvents2_Event).oncontextmenu, AddressOf Document_oncontextmenu
            RemoveHandler CType(objhtmldoc, HTMLDocumentEvents2_Event).onmouseout, AddressOf Document_onmouseout
            'AddHandler CType(objframehtmldoc, HTMLDocumentEvents2_Event).onclick, AddressOf Document_onclick
            'AddHandler CType(objframehtmldoc, HTMLDocumentEvents2_Event).onmouseover, AddressOf Document_onmouseover
            'AddHandler CType(objframehtmldoc, HTMLDocumentEvents2_Event).oncontextmenu, AddressOf Document_oncontextmenu
            'AddHandler CType(objframehtmldoc, HTMLDocumentEvents2_Event).onmouseout, AddressOf Document_onmouseout
        End If
    End Sub

    Private Function Document_onclick(ByVal e As mshtml.IHTMLEventObj) As Boolean
        Return True
    End Function
    Private Function Document_oncontextmenu(ByVal e As mshtml.IHTMLEventObj) As Boolean
        objcurelement = e.srcElement
        If objweby.btnspy.Text = "Stop Spy" Then
            If My.Computer.Keyboard.ShiftKeyDown Then
                objhtmldoc = objIE.Document
                GetAllElemtsfrompage(objhtmldoc)
            Else
                GetAllElementProperties(objcurelement, False)
            End If
            Return False
        Else
            Return True
        End If
    End Function

    Private Sub GetAllElementProperties(objelement As IHTMLElement, automationNaming As Boolean)
        Dim objop As ObjectProperties
        Static counter = 0
        objop = New ObjectProperties()

        Dim ObjName As String = ""
        If Not (IsNothing(objelement.getAttribute("Value")) Or IsDBNull(objelement.getAttribute("Value"))) Then
            ObjName = objelement.getAttribute("Value").ToString().Trim()
        ElseIf Not (IsNothing(objelement.getAttribute("OuterText")) Or IsDBNull(objelement.getAttribute("OuterText"))) Then
            ObjName = objelement.getAttribute("OuterText").ToString().Trim()
        ElseIf Not (IsNothing(objelement.getAttribute("InnerText")) Or IsDBNull(objelement.getAttribute("InnerText"))) Then
            ObjName = objelement.getAttribute("InnerText").ToString().Trim()
        ElseIf Not (IsNothing(objelement.getAttribute("Alt")) Or IsDBNull(objelement.getAttribute("Alt"))) Then
            ObjName = objelement.getAttribute("Alt").ToString().Trim()
        ElseIf Not (IsNothing(objelement.getAttribute("label")) Or IsDBNull(objelement.getAttribute("label"))) Then
            ObjName = objelement.getAttribute("label").ToString().Trim()
        ElseIf Not (IsNothing(objelement.getAttribute("title")) Or IsDBNull(objelement.getAttribute("title"))) Then
            ObjName = objelement.getAttribute("title").ToString().Trim()
        End If
        If ObjName.Trim() = "" Then
            If automationNaming = False Then
                ObjName = InputBox("Please Enter the Object Name")
            Else
                ObjName = "Objects_" & counter
                counter = counter + 1
            End If
        End If

        ObjName = ObjName.Replace(vbCrLf, "").Replace(vbNewLine, "")
        ObjName = Regex.Replace(ObjName, "[^A-Za-z0-9\-/]", "")

        If ObjName.Length() > 20 Then
            ObjName = ObjName.ToString().Replace(" ", "_").Substring(0, 20)
        Else
            ObjName = ObjName.ToString().Replace(" ", "_")
        End If

        If Not (IsNothing(objelement.getAttribute("name")) Or IsDBNull(objelement.getAttribute("name"))) Then
            objweby.addItemtoTree(objelement.tagName.ToLower() + "_" + ObjName, objelement.id, objelement.getAttribute("name"), objelement.tagName, objelement.className, objop.getXpath(objelement, False), objop.getXpath(objelement, True), objop.getCss(objelement), "css=" + objop.getCssSubPath(objelement))
        Else
            objweby.addItemtoTree(objelement.tagName.ToLower() + "_" + ObjName, objelement.id, "", objelement.tagName, objelement.className, objop.getXpath(objelement, False), objop.getXpath(objelement, True), objop.getCss(objelement), "css=" + objop.getCssSubPath(objelement))
        End If
    End Sub

    Private Sub GetAllElemtsfrompage(objhtmldoc As HTMLDocument)
        Dim allelements = objhtmldoc.all
        Dim eleTagName
        For Each element As IHTMLElement In allelements
            eleTagName = element.tagName.ToLower()
            If eleTagName = "input" Or eleTagName = "a" Or eleTagName = "table" Or eleTagName = "tr" Or eleTagName = "td" Or eleTagName = "button" Or eleTagName = "frame" Or eleTagName = "iframe" Or eleTagName = "select" Then
                GetAllElementProperties(element, True)
            End If
        Next
    End Sub

    Private Sub Document_onmouseout(ByVal e As mshtml.IHTMLEventObj)
        If objweby.btnspy.Text = "Stop Spy" Then
            If Not objprevelement Is Nothing Then
                objcurelement.style.border = strborder
                'objprevelement.style.setAttribute("border", "solid 0px #000000")
                objprevelement = objcurelement
            End If
        End If
    End Sub

    Private Sub setFrame(objcurelement As IHTMLElement)
        objframehtmldoc = objcurelement.document
        If frmhtmlhandle = False Then
            AddHandler CType(objframehtmldoc, HTMLDocumentEvents2_Event).onclick, AddressOf Document_onclick
            AddHandler CType(objframehtmldoc, HTMLDocumentEvents2_Event).onmouseover, AddressOf Document_onmouseover
            AddHandler CType(objframehtmldoc, HTMLDocumentEvents2_Event).oncontextmenu, AddressOf Document_oncontextmenu
            AddHandler CType(objframehtmldoc, HTMLDocumentEvents2_Event).onmouseout, AddressOf Document_onmouseout
        End If
    End Sub

    Private Function GetIE() As SHDocVw.InternetExplorer
        Dim objIEx As SHDocVw.InternetExplorer
        Dim objIEControls() As SHDocVw.InternetExplorer
        Dim intcounter = 0

        For Each objIEx In objShellWindows
            If objIEx.Name = "Internet Explorer" Then
                ReDim Preserve objIEControls(intcounter)
                objIEControls(intcounter) = objIEx
                intcounter = intcounter + 1
            End If
        Next

        If intcounter = 0 Then
            objIEx = New SHDocVw.InternetExplorer
            objIEx.Visible = True
            Return (objIEx)
        ElseIf intcounter = 1 Then
            Return (objIEControls(0))
        Else
            Dim strmsg As String
            strmsg = ""
            For i = 0 To objIEControls.Length - 1
                strmsg = strmsg & i & ") " & objIEControls(i).LocationName & " : " & objIEControls(i).LocationURL & vbNewLine & vbCrLf
            Next
            Dim intInputValue = InputBox(strmsg, "Select the Browser", 0)
            If intInputValue <> "" Then
                If intInputValue >= 0 And intInputValue < objIEControls.Length Then
                    Return (objIEControls(intInputValue))
                Else
                    MsgBox("Wrong selection!!! - Quiting")
                End If
            End If
        End If
        Return (Nothing)
    End Function
End Class
