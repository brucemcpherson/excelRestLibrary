'gistThat@mcpher.com :do not modify this line - see ramblings.mcpher.com for details: updated on 8/18/2014 3:54:11 PM : from manifest:5055578 gist https://gist.github.com/brucemcpherson/b72f7a826d553fb46cfa/raw/cDbab.cls
Option Explicit
' v1.0
' this one manages interaction with dbAbstraction on Google Apps Script
Private pDbId As String
Private pSiloId As String
Private pResult As cDbAbResult
Private poAuth2 As cOauth2
Private pEndPoint As String
Private pDbName As String
Private pBrowser As cBrowser
Private pNoCache As Long
Private pPeanut As String
Private pConstraints As cJobject
Public Function constraints(json As String) As String
' for example
  'stuff.age': handler.constraints("['GT' ,25],['LTE',60]]")
' needs to become
  'stuff.age':{'__CONSTR$KEY$':[{'constraint':'$gt','value':25},{'constraint':'$lte','value':60}]}
    Dim s As cStringChunker, job As cJobject, jo As cJobject
    Set s = New cStringChunker

     
    Set job = JSONParse(json)
    s.add "{'__CONSTR$KEY$':["
    For Each jo In job.children
        With s.add("{'value':")
            .add (maybeQuote(jo.children(2)))
            .add (",'constraint':'")
            .add(pConstraints.toString(jo.children(1).value)).add ("'")
            .add ("},")
        End With
    Next jo
    
    constraints = s.chopIf(",").add("]}").content
    
End Function
Private Function maybeQuote(jo As cJobject) As Variant
    Dim v As Variant
    If (jo.isArrayRoot) Then
        maybeQuote = jo.stringify
    Else
        v = jo.value
        
        If TypeName(v) = "string" Then
            maybeQuote = "'" & v & "'"
        Else
            maybeQuote = v
        End If
    End If
End Function

' the endpoint - your Google Apps Script webapp url
Public Function setEndPoint(endPoint As String) As cDbAb
    pEndPoint = endPoint
    Set setEndPoint = Me
End Function
Public Function getEndPoint() As String
    getEndPoint = pEndPoint
End Function
' the result of the last fetch
Public Function getResult() As cDbAbResult
    Set getResult = pResult
End Function

' the siloid is roughly equivalent to a tablename
Public Function setSiloId(id As String) As cDbAb
    pSiloId = id
    Set setSiloId = Me
End Function
Public Function getSiloId() As String
    getSiloId = pSiloId
End Function

' this is the oauth2 object used to provide the accesstoken
Public Function setOauth2(oauth2 As cOauth2) As cDbAb
    Set poAuth2 = oauth2
    Set setOauth2 = Me
End Function

' the dbid is roughtly equivalent to the database name
Public Function setDbId(id As String) As cDbAb
    pDbId = id
    Set setDbId = Me
End Function
Public Function getDbId() As String
    getDbId = pDbId
End Function

' any special id to use in google analytics when the api call is serviced
Public Function setPeanut(id As String) As cDbAb
    pPeanut = id
    Set setPeanut = Me
End Function

Private Function getQueryString(Optional queryJSON As String = vbNullString) As String
    Dim queryOb As cJobject
    If (queryJSON <> vbNullString) Then
        Set queryOb = JSONParse(queryJSON)
    End If
    If (isSomething(queryOb)) Then
        getQueryString = "&query=" & URLEncode(queryOb.stringify)
        queryOb.tearDown
    Else
        getQueryString = vbNullString
    End If
    
End Function
Private Function getParamString(Optional paramJSON As String = vbNullString) As String
    Dim paramOb As cJobject
    If (paramJSON <> vbNullString) Then
        Set paramOb = JSONParse(paramJSON)
    End If
    If (isSomething(paramOb)) Then
        getParamString = "&params=" & URLEncode(paramOb.stringify)
        paramOb.tearDown
    Else
        getParamString = vbNullString
    End If
    
End Function
Private Function getAuthHeader() As String
    getAuthHeader = poAuth2.authHeader
End Function
' the dbname is the name of the type of db .. eg SHEET
Public Function setDbName(dbName As String) As cDbAb
    pDbName = dbName
    Set setDbName = Me
End Function
Public Function getDbName() As String
    getDbName = pDbName
End Function
Public Property Get browser() As cBrowser
    Set browser = pBrowser
End Property
Public Function setNoCache(noCache As Long) As cDbAb
    pNoCache = noCache
    Set setNoCache = Me
End Function
Private Function makeUrl(action As String, Optional noCache As Long = 0, Optional keepid As Boolean = False, _
            Optional queryJSON As String = vbNullString, Optional paramsJSON As String = vbNullString) As String
    Dim s As New cStringChunker
    s.add(getEndPoint()) _
     .add("?").add("driver=").add(getDbName()) _
     .add("&").add("action=").add(action) _
     .add("&").add("siloid=").add(getSiloId()) _
     .add("&").add("dbid=").add(getDbId) _
     .add("&").add("nocache=").add(CStr(noCache)) _
     .add("&").add("keepid=").add(CLng(keepid) * -1) _
     .add("&").add("peanut=").add(CStr(pPeanut)) _
     .add(getQueryString(queryJSON)) _
     .add getParamString(paramsJSON)

    makeUrl = s.content
    
End Function
' dbabstraction save
' @param {cJobect} obs the data to save
' @return {cDbAbResult} the result
Public Function save(obs As cJobject) As cDbAbResult
    Set save = execute("save", "POST", , , obs)
    
End Function
Public Function query(Optional queryJSON As String = vbNullString, _
        Optional paramsJSON As String = vbNullString, _
        Optional noCache As Long = 0, _
        Optional keepid As Boolean = False) As cDbAbResult

    Set query = execute("query", "GET", queryJSON, paramsJSON, , , noCache, keepid)

End Function
Public Function update(keys As cJobject, obs As cJobject) As cDbAbResult
    Set update = execute("update", "POST", , , obs, keys, 1, 0)
End Function
Public Function remove(Optional queryJSON As String = vbNullString, _
        Optional paramsJSON As String = vbNullString) As cDbAbResult
    Set remove = execute("remove", "POST", queryJSON, paramsJSON, , , 1, 0)
End Function
Public Function count(Optional queryJSON As String = vbNullString, _
        Optional paramsJSON As String = vbNullString, _
        Optional noCache As Long = 0) As cDbAbResult

    Set count = execute("count", "GET", queryJSON, paramsJSON, , , noCache, 0)
End Function
Public Function getObjects(keys As cJobject, Optional noCache As Long = 0, _
                Optional keepid As Boolean = False) As cDbAbResult
    ' normally called get, but vba reserved name
    Set getObjects = execute("get", "POST", , , , keys, noCache, keepid)
    
End Function
Private Function execute(action As String, _
        Optional method As String = "GET", _
        Optional queryJSON As String = vbNullString, _
        Optional paramsJSON As String = vbNullString, _
        Optional data As cJobject = Nothing, _
        Optional keys As cJobject = Nothing, _
        Optional noCache As Long = 0, _
        Optional keepid As Boolean = False) As cDbAbResult
    Dim result As String, payload As String, url As String, s As cStringChunker
    Set s = New cStringChunker
    If (pNoCache > 0) Then noCache = 1
    url = makeUrl(action, noCache, keepid, queryJSON, paramsJSON)

    If (method = "GET") Then
        pBrowser.httpGET url, , , , , getAuthHeader(), , method
    Else
        If (isSomething(keys)) Then
            s.add (keys.stringify)
            If (isSomething(data)) Then s.chopIf("}").add(",").add (Mid(data.stringify, 2))
        
        ElseIf (isSomething(data)) Then
            s.add (data.stringify)
        
        End If

        payload = s.content

        pBrowser.httpPost url, payload, True, getAuthHeader(), , method
    End If
    Set pResult = New cDbAbResult
    pResult.setResult pBrowser
    Set execute = pResult
End Function

Private Sub Class_Initialize()

    Set pBrowser = New cBrowser
    pNoCache = 0
    pPeanut = getUserHash()
    Set pConstraints = JSONParse( _
    "{'LT':'$lt','GTE':'$gte', 'GT':'$gt', 'NE':'$ne', 'IN':'$in','NIN':'$nin','EQ':'$eq','LTE':'$lte'}")
    
End Sub

Public Function tearDown()
    If isSomething(pConstraints) Then
        pConstraints.tearDown
    End If
    pBrowser.tearDown
    If (isSomething(pResult)) Then
        pResult.tearDown
    End If
End Function

