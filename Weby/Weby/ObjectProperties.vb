Public Class ObjectProperties
    Public Function getXpath(element As mshtml.IHTMLElement, Absolute As Boolean) As String
        Dim isBreak As Boolean
        Dim path() As String
        ReDim path(0)
        Dim i As Integer

        i = -1
        isBreak = False

        If element Is Nothing Then
            getXpath = ""
        End If

        Dim objcurrentnode As mshtml.IHTMLElement
        objcurrentnode = element


        While Not objcurrentnode Is Nothing
            i = i + 1
            ReDim Preserve path(i)
            Dim pe As String
            pe = getNode(objcurrentnode, Absolute)
            If pe <> "" Then
                path(i) = pe
                REM Found an ID, no need to go upper, absolute path is OK
                Dim tempvar
                If InStr(pe, "@id") >= 1 Or InStr(pe, "@name") >= 1 Then
                    tempvar = pe
                    If (Len(tempvar) - Len(Replace(tempvar, "-", "")) > 1) Then
                    Else
                        isBreak = True
                    End If
                End If

            End If
            If isBreak = False Then
                objcurrentnode = objcurrentnode.parentElement
            Else
                objcurrentnode = Nothing
            End If
        End While

        Dim ResultXPath
        ResultXPath = joininReverse(path, "/", i)
        If Absolute = False Then
            If Not InStr(ResultXPath, "html") >= 1 Then
                getXpath = "/" + joininReverse(path, "/", i)
            Else
                getXpath = Replace(ResultXPath, "/html/body/", "//")
            End If
        Else
            getXpath = joininReverse(path, "/", i)
        End If
    End Function

    Public Function getNode(node As mshtml.IHTMLElement, Absolute As Boolean) As String
        Dim nodeExpr As String
        Dim isBreak As Boolean
        Dim isContinue As Boolean

        nodeExpr = ""
        nodeExpr = LCase(node.tagName)
        isBreak = False
        isContinue = True


        If nodeExpr = "" Then
            getNode = ""
        End If

        If Not Absolute Then
            Dim elementName
            elementName = getElementName(node)
            If Not node.id = "" Then
                If InStr(node.id, "-") > 0 Then
                    If UBound(Split(node.id, "-")) < 1 Then
                        nodeExpr = nodeExpr + "[@id='" + node.id + "']"
                        REM IDs are supposed to be unique, so they are a good starting point.
                        isContinue = False
                    End If
                Else
                    nodeExpr = nodeExpr + "[@id='" + node.id + "']"
                    REM IDs are supposed to be unique, so they are a good starting point.
                    isContinue = False
                End If
            ElseIf Not elementName = "" Then
                nodeExpr = nodeExpr + "[@name='" + elementName + "']"
                REM IDs are supposed to be unique, so they are a good starting point.
                isContinue = False
            End If
        End If

        If isContinue = True Then
            REM Find rank of node among its type in the parent
            Dim rank As Integer
            rank = 1
            Dim nodeDom As mshtml.IHTMLDOMNode
            nodeDom = node
            Dim psDom As mshtml.IHTMLDOMNode
            psDom = nodeDom.previousSibling


            While Not psDom Is Nothing
                If psDom.nodeName = node.tagName Then
                    rank = rank + 1
                End If

                psDom = psDom.previousSibling
            End While

            If rank > 1 Then
                nodeExpr = nodeExpr + "[" + Trim(Str(rank)) + "]"
            Else
                REM First node of its kind at this level. Are there any others?
                Dim nsDom As mshtml.IHTMLDOMNode
                nsDom = nodeDom.nextSibling

                While Not nsDom Is Nothing
                    If nsDom.nodeName = node.tagName Then
                        REM Yes, mark it as being the first one
                        nodeExpr = nodeExpr + "[1]"
                        isBreak = True
                    End If
                    If isBreak = True Then
                        nsDom = Nothing
                    Else
                        nsDom = nsDom.nextSibling
                    End If
                End While
            End If
        End If
        getNode = nodeExpr
    End Function


    Public Function joininReverse(items As Object, delimiter As String, totalSiblingCount As Integer) As String
        Dim actXpath As String
        Dim siblingPosition As Integer
        actXpath = ""

        For siblingPosition = 0 To totalSiblingCount
            If Not items(totalSiblingCount - siblingPosition) = "" Then
                actXpath = actXpath + delimiter + items(totalSiblingCount - siblingPosition)
            End If
        Next siblingPosition
        joininReverse = actXpath
    End Function



    Public Function getNodeNumber(currentnode As mshtml.IHTMLElement) As String
        Dim rank
        rank = 0
        Dim nodeDom As mshtml.IHTMLDOMNode
        nodeDom = currentnode
        Dim psDom As mshtml.IHTMLDOMNode
        psDom = nodeDom.previousSibling


        While Not psDom Is Nothing
            If psDom.nodeName = currentnode.tagName Then
                rank = rank + 1
            End If
            psDom = psDom.previousSibling
        End While
        getNodeNumber = Str(rank)
    End Function

    Public Function getCssSubPath(currentnode As mshtml.IHTMLElement) As String
        Dim attr
        Dim attrValue As Object
        Dim cssAttributes(6) As String
        cssAttributes(0) = "id"
        cssAttributes(1) = "name"
        cssAttributes(2) = "class"
        cssAttributes(3) = "type"
        cssAttributes(4) = "alt"
        cssAttributes(5) = "title"
        cssAttributes(6) = "value"

        For Each attr In cssAttributes

            attrValue = currentnode.getAttribute(attr)
            If (IsDBNull(attrValue)) Then
                If (attrValue.ToString <> "") Then
                    If attr = "id" Or attr = "name" Then
                        getCssSubPath = "#" + attrValue
                        Exit Function
                    End If
                    If attr = "class" Then
                        getCssSubPath = LCase(currentnode.tagName) + "." + attrValue.Replace(" ", ".").Replace("..", ".")
                        Exit Function
                    End If
                    getCssSubPath = LCase(currentnode.tagName) + "[" + attr + "=' + value + ']"
                    Exit Function
                End If
            End If
        Next
        If (CInt(getNodeNumber(currentnode)) > 0) Then
            getCssSubPath = LCase(currentnode.tagName) + ":nth-of-type(" + Trim(getNodeNumber(currentnode)) + ")"
        Else
            getCssSubPath = LCase(currentnode.tagName)
        End If
    End Function

    Public Function getCss(currentnode As mshtml.IHTMLElement) As String
        Dim subpath As String
        Dim current As mshtml.IHTMLElement
        current = currentnode
        subpath = getCssSubPath(current)


        While LCase(current.tagName) <> "html"
            subpath = getCssSubPath(current.parentElement) + " > " + subpath
            current = current.parentElement
        End While
        getCss = "css=" + subpath
    End Function

    Public Function getElementName(objcurelement As mshtml.IHTMLElement) As String

        On Error Resume Next
        Dim nameTagPosition As Integer
        Dim elementName As String
        Dim htmlStr As String
        htmlStr = objcurelement.outerHTML
        htmlStr = Mid(htmlStr, 2, InStr(2, htmlStr, ">") - 2)
        nameTagPosition = InStr(1, htmlStr, "name=")
        If nameTagPosition > 1 Then
            elementName = objcurelement.getAttribute("name")
        Else
            elementName = ""
        End If

        getElementName = elementName
    End Function
End Class
