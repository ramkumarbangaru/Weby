Imports mshtml
Public Class IEBrowser
    REM Declaring objects with Events
    Dim WithEvents objhtmldoc As HtmlDocument
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

    Public Sub New(ByVal caller As frmweby)
        GetIEWindow()
        objweby = frmweby
        htmlhandle = False
        frmhtmlhandle = False
    End Sub

    Public Sub GetIEWindow()
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
            Dim frameName As String
            Dim frameId As String
            Dim frameUniqueId As String
            objprevelement = Nothing
            objcurelement = e.srcElement

            If InStr(1, LCase(objcurelement.tagName), "frame") >= 1 Then

                frameName = objIE.Document.parentWindow.event.srcElement.Name
                frameId = objIE.Document.parentWindow.event.srcElement.ID
                frameUniqueId = objIE.Document.parentWindow.event.srcElement.uniqueID
                If frameName <> "" Then
                    Call setFrameByName(frameName)
                ElseIf frameId <> "" Then
                    Call setFrameById(frameId)
                Else
                    Call setFrame(objcurelement, frameUniqueId)
                End If
            Else
                frameName = ""
                frameId = ""
                frameUniqueId = ""
                objprevelement = objcurelement
                objcurelement.style.setAttribute("border", "solid 1px #ff0000")
            End If

            If InStr(1, LCase(objcurelement.tagName), "frameset") >= 1 Then
                objcurelement = objcurelement.document.parentWindow.event.srcElement
            End If

            Call objweby.setValues(objcurelement)
        End If
    End Sub
    Private Function Document_onclick(ByVal e As mshtml.IHTMLEventObj) As Boolean
        Return True
    End Function
    Private Function Document_oncontextmenu(ByVal e As mshtml.IHTMLEventObj) As Boolean
        If objweby.btnspy.Text = "Stop Spy" Then
            Dim objop As ObjectProperties
            objop = New ObjectProperties()
            objweby.addItemtoTree(objcurelement.tagName + "_" + InputBox("Please Enter the Object Name"), objcurelement.id, objop.getElementName(objcurelement), objcurelement.tagName, objcurelement.className, objop.getXpath(objcurelement, False), objop.getXpath(objcurelement, True), objop.getCss(objcurelement), "css=" + objop.getCssSubPath(objcurelement))
            Return False
        Else
            Return True
        End If
    End Function
    Private Sub Document_onmouseout(ByVal e As mshtml.IHTMLEventObj)
        If objweby.btnspy.Text = "Stop Spy" Then
            If Not objprevelement Is Nothing Then
                objprevelement.style.setAttribute("border", "solid 0px #000000")
                objprevelement = objcurelement
            End If
        End If
    End Sub

    Private Sub setFrameByName(frameName As String)
        For i = 0 To (objIE.Document.frames.Length - 1)
            If objIE.Document.frames(i).Name = frameName Then
                objframehtmldoc = objIE.Document.frames(i).Document
            End If
        Next
        If frmhtmlhandle = False Then
            AddHandler CType(objframehtmldoc, HTMLDocumentEvents2_Event).onclick, AddressOf Document_onclick
            AddHandler CType(objframehtmldoc, HTMLDocumentEvents2_Event).onmouseover, AddressOf Document_onmouseover
            AddHandler CType(objframehtmldoc, HTMLDocumentEvents2_Event).oncontextmenu, AddressOf Document_oncontextmenu
            AddHandler CType(objframehtmldoc, HTMLDocumentEvents2_Event).onmouseout, AddressOf Document_onmouseout
            frmhtmlhandle = True
        End If
    End Sub
    Private Sub setFrameById(frameId As String)
        Dim i
        For i = 0 To (objIE.Document.frames.Length - 1)
            If objIE.Document.frames(i).ID = frameId Then
                objframehtmldoc = objIE.Document.frames(i).Document
            End If
        Next
        If frmhtmlhandle = False Then
            AddHandler CType(objframehtmldoc, HTMLDocumentEvents2_Event).onclick, AddressOf Document_onclick
            AddHandler CType(objframehtmldoc, HTMLDocumentEvents2_Event).onmouseover, AddressOf Document_onmouseover
            AddHandler CType(objframehtmldoc, HTMLDocumentEvents2_Event).oncontextmenu, AddressOf Document_oncontextmenu
            AddHandler CType(objframehtmldoc, HTMLDocumentEvents2_Event).onmouseout, AddressOf Document_onmouseout
            frmhtmlhandle = True
        End If
    End Sub

    Private Sub setFrame(cur As Object, frameId As String)
        objframehtmldoc = objIE.Document.parentWindow.event.srcElement.Document
        If frmhtmlhandle = False Then
            AddHandler CType(objframehtmldoc, HTMLDocumentEvents2_Event).onclick, AddressOf Document_onclick
            AddHandler CType(objframehtmldoc, HTMLDocumentEvents2_Event).onmouseover, AddressOf Document_onmouseover
            AddHandler CType(objframehtmldoc, HTMLDocumentEvents2_Event).oncontextmenu, AddressOf Document_oncontextmenu
            AddHandler CType(objframehtmldoc, HTMLDocumentEvents2_Event).onmouseout, AddressOf Document_onmouseout
            frmhtmlhandle = True
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
