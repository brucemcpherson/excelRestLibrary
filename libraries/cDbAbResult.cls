'gistThat@mcpher.com :do not modify this line - see ramblings.mcpher.com for details: updated on 8/18/2014 3:54:12 PM : from manifest:5055578 gist https://gist.github.com/brucemcpherson/b72f7a826d553fb46cfa/raw/cDbabResult.cls
Option Explicit
' v1.10
' this is a dbab result

Private pResult As cJobject

Public Property Get handleError() As String
    handleError = pResult.child("handleError").value
End Property
Public Property Get handleCode() As Long
    handleCode = pResult.child("handleCode").value
End Property
Public Property Get handleKeys() As cJobject
    Set handleKeys = pResult.child("handleKeys")
End Property
Public Property Get driverKeys() As cJobject
    Set driverKeys = pResult.child("driverKeys")
End Property
Public Property Get data() As cJobject
    Set data = pResult.child("data")
End Property
Public Property Get response() As cJobject
    Set response = pResult
End Property
Public Property Get length() As Long
    length = 0
    If (isSomething(data)) Then
        length = data.children.count
    End If
End Property
Public Property Get count() As Long
    Dim c As cJobject
    count = 0
    If (isSomething(data)) Then
        If (data.hasChildren) Then
            Set c = data.children(1).child("count")
            If (isSomething(c)) Then
                count = c.value
            Else
                count = length
            End If
        End If
    End If
End Property
Public Function setResult(browser As cBrowser) As cDbAbResult
    Set pResult = JSONParse(browser.Text, False)
    If (pResult Is Nothing) Then
        MsgBox ("invalid json data:" + browser.Text)
    End If
    Set setResult = Me
End Function

Public Function tearDown()
    If (isSomething(pResult)) Then
        pResult.tearDown
    End If
End Function

