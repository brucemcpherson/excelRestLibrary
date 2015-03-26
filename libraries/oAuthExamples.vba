'gistThat@mcpher.com :do not modify this line - see ramblings.mcpher.com for details: updated on 8/18/2014 3:54:15 PM : from manifest:5055578 gist https://gist.github.com/brucemcpherson/6937450/raw/oAuthExamples.vba
Option Explicit
' oauth examples
' v1.2
' convienience function for auth..
Public Function getGoogled(scope As String, _
                                Optional replacementpackage As cJobject = Nothing, _
                                Optional clientID As String = vbNullString, _
                                Optional clientSecret As String = vbNullString, _
                                Optional complain As Boolean = True, _
                                Optional cloneFromeScope As String = vbNullString) As cOauth2
    Dim o2 As cOauth2
    Set o2 = New cOauth2
    With o2.googleAuth(scope, replacementpackage, clientID, clientSecret, complain, cloneFromeScope)
        If Not .hasToken And complain Then
            MsgBox ("Failed to authorize to google for scope " & scope & ":denied code " & o2.denied)
        End If
    End With
    
    Set getGoogled = o2
End Function
Private Sub testOauth2()
    Dim myConsole As cJobject
    ' if you are calling for the first time ever you can either provide your
    ' clientid/clientsecret - or pass the the jsonparse retrieved from the google app console
    ' normally all this stuff comes from encrpted registry store
    
    ' first ever
    'Set myConsole = makeMyGoogleConsole
    'With getGoogled("analytics", myConsole)
    '    Debug.Print .authHeader
   '     .tearDown
   ' End With

    'or you can do first ever like this
    'With getGoogled("viz", , "xxxxx.apps.googleusercontent.com", "xxxxx")
    '    Debug.Print .authHeader
    '    .tearDown
    'End With
    
    ' all other times this is what is needed
    With getGoogled("drive")
        Debug.Print .authHeader

        .tearDown
    End With
    ' lets auth and have a look at the contents
    'Debug.Print objectStringify(getGoogled("drive"))
    
    ' all other times this is what is needed
    With getGoogled("analytics")
        Debug.Print .authHeader
        .tearDown
    End With
    
    ' here's an example of cloning credentials from a different scope for 1st time in
    With getGoogled("urlshortener", , , , , "drive")
        Debug.Print .authHeader
        .tearDown
    End With
    
    With getGoogled("urlshortener")
        Debug.Print .authHeader
        .tearDown
    End With
    
    ' if you made one, clean it up
    If Not myConsole Is Nothing Then
        myConsole.tearDown
    End If
End Sub

Private Function makeMyGoogleConsole() As cJobject
    Dim consoleJSON As String
 
     consoleJSON = _
    "{'installed':{'auth_uri':'https://accounts.google.com/o/oauth2/auth'," & _
      "'client_secret':'xxxxxxxx'," & _
      "'token_uri':'https://accounts.google.com/o/oauth2/token'," & _
      "'client_email':'','redirect_uris':['urn:ietf:wg:oauth:2.0:oob','oob']," & _
      "'client_x509_cert_url':'','client_id':'xxxxxxx.apps.googleusercontent.com'," & _
      "'auth_provider_x509_cert_url':'https://www.googleapis.com/oauth2/v1/certs'}}"
      
      Set makeMyGoogleConsole = JSONParse(consoleJSON)

End Function
   
Private Sub showLinkedinConsole()
    Dim url As String, cb As cBrowser
    Set cb = New cBrowser
    
    ' see http://excelramblings.blogspot.co.uk/2012/10/somewhere-to-keep-those-api-keys-google.html
    ' for how to store credentials in a google lockbox
    url = "https://script.google.com/a/macros/mcpher.com/s/" & _
    "AKfycbza96-Mpa47jlqXoPosk64bUfR8T7zO5POZMYyN45InrvX8gm28/exec" & _
    "?action=show&entry=linkedinauth"
    
    With getGoogled("drive")
        If .hasToken Then
            Debug.Print cb.httpGET(url, , , , , .authHeader)
        Else
            MsgBox ("failed to authenticate: " & .denied)
        End If
        .tearDown
    End With

    cb.tearDown
    
End Sub
' db abstraction tests
Private Function every(a As Variant) As Boolean
    Dim i As Long, good As Boolean
    If (IsArray(a)) Then
        For i = LBound(a) To UBound(a)
            If (Not a(i)) Then
                good = False
                Exit For
            End If
            good = True
        Next i
    Else
        good = a
    End If
    every = good
End Function
Private Function assert(what As Variant, message As cJobject, n As String) As Boolean

    Dim fatal As Boolean, m As String, good As Boolean
    fatal = True
    good = every(what)
    m = "assertion:" & n
    If (Not good) Then
        m = m & ":failed:" & message.stringify
        If fatal Then
            Debug.Print m
            Debug.Assert good
        End If
    End If
    assert = good
    Debug.Print m
End Function
Private Function setUpTest(dbtype As String, siloid As String, dbid As String) As cDbAb
    ' get authorized - using drive scope
    Dim oauth2 As cOauth2, handler As cDbAb
    Dim ds As cDataSet, testData As cJobject

    ' set up handler & set end point
    Set oauth2 = getGoogled("drive")
    Set handler = New cDbAb
    With handler.setOauth2(oauth2)
        .setEndPoint ("https://script.google.com/macros/s/AKfycbyfapkJpd4UOhiqLOJOGBb11nG4BTru_Bw8bZQ49eQSTfL2vbU/exec")
        .setDbId ("13ccFPXI0L8-ZViHlv8qoVspotUcnX8v0ZFeY4nUP574")
        .setNoCache (1)
        .setDbId (dbid)
        .setSiloId (siloid)
        .setDbName dbtype
        
    End With
    Set setUpTest = handler

End Function
Public Function testDbAb()

    ' get authorized - using drive scope
    Dim handler As cDbAb
    Dim ds As cDataSet, testData As cJobject, result As cDbAbResult, testSheet As String

    'get testdata from a google sheet & write it to an excel sheet
    testSheet = "dbab"
    Set handler = setUpTest("sheet", "customers", "13ccFPXI0L8-ZViHlv8qoVspotUcnX8v0ZFeY4nUP574")
    Set result = handler.query()
    assert result.handleCode >= 0, result.response, "getting testData from google sheet"
    Set ds = makeSheetFromJob(result.data, testSheet)
    Set testData = ds.jObject(, , , , "data")

    ' set up handler & set end point
    Set handler = setUpTest("datastore", ds.name, "xliberationdatastore")
    lotsoftests handler, testData
    handler.tearDown
    
    Set handler = setUpTest("mongolab", ds.name, "xliberation")
    lotsoftests handler, testData
    handler.tearDown
    
    Set handler = setUpTest("parse", ds.name, "xliberation")
    lotsoftests handler, testData
    handler.tearDown
    
    Set handler = setUpTest("drive", ds.name, "/datahandler/driverdrive")
    lotsoftests handler, testData
    handler.tearDown
    
    Set handler = setUpTest("sheet", ds.name, "13ccFPXI0L8-ZViHlv8qoVspotUcnX8v0ZFeY4nUP574")
    lotsoftests handler, testData
    handler.tearDown
    

    ds.tearDown
End Function

Private Sub lotsoftests(handler As cDbAb, testData As cJobject)
   ' remove from last time
    Dim result As cDbAbResult, r2 As cDbAbResult
    Dim x As Long, job As cJobject
    Set result = handler.remove()
    
    Debug.Print "Starting " & handler.getDbName
    assert result.handleCode >= 0, result.response, "removing initial"
    
    ' save the new data
    Set result = handler.save(testData)
    assert result.handleCode >= 0, result.response, "saving initial"
    
    ' query and make sure it matches what was saved
    Set result = handler.query()
    assert Array(result.handleCode >= 0, result.count = testData.children.count), _
            result.response, "querying initial"
    
    '--------query everything with limit
    Set result = handler.query(, "{'limit':2}")
    assert Array(result.handleCode >= 0, result.count = 2), _
            result.response, "limit test(" & result.count & ")"
            
    '------Sort Reverse
    Set result = handler.query(, "{'sort':'-name'}")
    assert Array(result.handleCode >= 0, result.count = testData.children.count), _
            result.response, "querysort(" & result.count & ")"
 
    '------Sort Reverse/skip
    Set result = handler.query(, "{'sort':'-name','skip':3}")
    assert Array(result.handleCode >= 0, result.count = testData.children.count - 3), _
            result.response, "querysort+skip(" & result.count & ")"
            
    '------query simple nosql
    Set result = handler.query("{'name':'ethel'}")
    x = 0
    For Each job In testData.children
        x = x + -1 * CLng(job.child("name").value = "ethel")
    Next job
    assert Array(result.handleCode >= 0, result.count = x), _
            result.response, "filterdot0(" & result.count & ")"

    '------query multi level
    Set result = handler.query("{'stuff':{'sex':'female'}}")
    x = 0
    For Each job In testData.children
        x = x + -1 * CLng(job.child("stuff.sex").value = "female")
    Next job
    assert Array(result.handleCode >= 0, result.count = x), _
            result.response, "filter(" & result.count & ")"
            
     '------queries in
    Set result = handler.query("{'name':" & _
        handler.constraints("[['IN',['ethel','fred']]]") & "}", , True)
    x = 0
    For Each job In testData.children
        x = x + -1 * (job.toString("name") = "ethel" Or job.toString("name") = "fred")
    Next job
    assert Array(result.handleCode >= 0, result.count = x), _
            result.response, "filterdotc4 (" & result.count & ")"
                  
            
    '------first complex constraints
    Set result = handler.query("{'stuff.age':" & _
        handler.constraints("[['GT',25],['LTE',60]]") & "}")
    ' checking results kind of long winded in vba
    x = 0
    For Each job In testData.children
        x = x + -1 * CLng(job.child("stuff.age").value > 25 And job.child("stuff.age").value <= 60)
    Next job
    assert Array(result.handleCode >= 0, result.count = x), _
            result.response, "querying initial complex(" & result.count & ")"


    '------query single constraint
    Set result = handler.query("{'stuff':{'age':" & _
        handler.constraints("[['GT',25]]") & "}}")
    x = 0
    For Each job In testData.children
        x = x + -1 * CDbl(job.child("stuff.age").value > 25)
    Next job
    assert Array(result.handleCode >= 0, result.count = x), _
           result.response, "filterdotc1 (" & result.count & ")"

  '------2 queries same constraint
    Set result = handler.query("{'stuff':{'age':" & _
        handler.constraints("[['GT',25],['LT',60]]") & "}}")
    x = 0
    For Each job In testData.children
        x = x + -1 * CDbl(job.child("stuff.age").value > 25 And job.child("stuff.age").value < 60)
    Next job
    assert Array(result.handleCode >= 0, result.count = x), _
            result.response, "filterdotc2 (" & result.count & ")"

  '------2 queries same constraint
    Set result = handler.query("{'stuff':{'sex':'male', 'age':" & _
        handler.constraints("[['GTE',25],['LT',60]]") & "}}", , True)
    x = 0
    For Each job In testData.children
        x = x + -1 * (job.child("stuff.age").value >= 25 And job.child("stuff.age").value < 60 _
            And job.child("stuff.sex").value = "male")
    Next job
    assert Array(result.handleCode >= 0, result.count = x), _
            result.response, "filterdotc3 (" & result.count & ")"

  '------queries in +
    Set result = handler.query( _
        "{'name':" & handler.constraints("[['IN',['john','mary']]]") & _
        ",'stuff.sex':'male','stuff.age':" & handler.constraints("[['GT',25]]") & "}")
    x = 0
    For Each job In testData.children
        x = x + -1 * (job.child("stuff.sex").value = "male" And job.child("stuff.age").value > 25 And _
            (job.toString("name") = "john" Or job.toString("name") = "mary"))
    Next job
    assert Array(result.handleCode >= 0, result.count = x), _
            result.response, "filterdotc5 (" & result.count & "/" & x & ")"
            
    '------query single constraint, get keys
    Set result = handler.query( _
        "{'stuff.age':" & handler.constraints("[['GT',25]]") & "}", _
        "{'limit':1}", , True)
    x = 1
    assert Array(result.handleCode >= 0, result.handleKeys.children.count = 1), _
            result.response, "limitkeycheck1 (" & result.count & ")"

    '-------testing Get -- known as getobjects because get is reserved in vba
    Set r2 = handler.getObjects(result.handleKeys)
    x = 0
    For Each job In r2.data.children
        x = x + -1 * CDbl(job.child("stuff.age").value > 25)
    Next job
    assert Array(r2.handleCode >= 0, r2.count = 1, x = r2.count), _
            result.response, "get1 (" & r2.count & ")"
            
    '------retest constraint
    Set result = handler.query("{'stuff':{'age':" & _
        handler.constraints("[['GT',60]]") & "}}")
    x = 0
    For Each job In testData.children
        x = x + -1 * CDbl(job.child("stuff.age").value > 60)
    Next job
    assert Array(result.handleCode >= 0, result.count = x), _
           result.response, "repeat test easy (" & result.count & ")"
           
    '------get ready for update test
    Set result = handler.query("{'stuff.sex':'male'}", , 1, 1)
    x = 0
    For Each job In testData.children
        x = x + -1 * CDbl(job.child("stuff.sex").value = "male")
    Next job
    assert Array(result.handleCode >= 0, result.handleKeys.children.count = x), _
           result.response, "does male work(" & result.count & ")"
   
    '----- do the update
    'first update the data with a new field
    For Each job In result.data.children
        job.add "stuff.man", job.child("stuff.sex").value = "male"
    Next job
    ' now update it
    Set r2 = handler.update(result.handleKeys, result.data)
    assert Array(r2.handleCode = 0), _
           r2.response, "update 2 (" & r2.count & ")"
           
  '------check previous query still works
    Set result = handler.query( _
        "{'name':" & handler.constraints("[['IN',['john','mary']]]") & _
        ",'stuff.sex':'male','stuff.age':" & handler.constraints("[['GT',25]]") & "}")
    x = 0
    For Each job In testData.children
        x = x + -1 * (job.child("stuff.sex").value = "male" And job.child("stuff.age").value > 25 And _
            (job.toString("name") = "john" Or job.toString("name") = "mary"))
    Next job
    assert Array(result.handleCode >= 0, result.count = x), _
            result.response, "repeat test after update (" & result.count & "/" & x & ")"
            
    ' query again and make sure it matches what was saved
    Set result = handler.query()
    assert Array(result.handleCode >= 0, result.count = testData.children.count), _
            result.response, "repeat querying initial"
         
    ' try counting
    Set result = handler.count()
    assert Array(result.handleCode >= 0, result.count = testData.children.count), _
            result.response, "count 1"
            
    ' try complicated counting
    Set result = handler.count( _
        "{'name':" & handler.constraints("[['IN',['john','mary']]]") & _
        ",'stuff.sex':'male','stuff.age':" & handler.constraints("[['GT',25]]") & "}")
    x = 0
    For Each job In testData.children
        x = x + -1 * (job.child("stuff.sex").value = "male" And job.child("stuff.age").value > 25 And _
            (job.toString("name") = "john" Or job.toString("name") = "mary"))
    Next job
    assert Array(result.handleCode >= 0, result.count = x), _
            result.response, "complex counting (" & result.count & "/" & x & ")"
   
    
    '--------------some more
    Set result = handler.query( _
        "{ 'stuff.sex':'male','stuff.age':" & handler.constraints("[['GT',59]]") & "}")
    x = 0
    For Each job In testData.children
        x = x + -1 * (job.child("stuff.sex").value = "male" And job.child("stuff.age").value >= 60)
    Next job
    assert Array(result.handleCode >= 0, result.count = x), _
            result.response, "normal 0 (" & result.count & "/" & x & ")"
            
 
    
    '--------------make sure we're getting the right id with complex constaints
    Set result = handler.query( _
        "{'stuff.age':" & handler.constraints("[['GT',25],['LTE',60]]") & "}")
    x = 0
    For Each job In testData.children
        x = x + -1 * (job.child("stuff.age").value > 25 And job.child("stuff.age").value <= 60)
    Next job
    assert Array(result.handleCode >= 0, result.count = x), _
            result.response, "repeat test constraint (" & result.count & "/" & x & ")"
            
    '--------------try OR
    Set result = handler.query( _
        "[{'stuff.age':" & handler.constraints("[['LT',26]]") & ",'stuff.sex':'male'}," & _
         "{'stuff.age':" & handler.constraints("[['GTE',60]]") & ",'stuff.sex':'male'}]")
    x = 0
    For Each job In testData.children
        x = x + -1 * (job.child("stuff.sex").value = "male" And (job.child("stuff.age").value < 26 Or job.child("stuff.age").value >= 60))
    Next job
    assert Array(result.handleCode >= 0, result.count = x), _
            result.response, "OR 1 (" & result.count & "/" & x & ")"
            
            
    '------------show all the males
    Set r2 = handler.query("{'stuff.sex':'male'}", , 1, 1)
    x = 0
    For Each job In testData.children
        x = x + -1 * CDbl(job.child("stuff.sex").value = "male")
    Next job
    assert Array(r2.handleCode >= 0, r2.handleKeys.children.count = x), _
           r2.response, "show the males(" & r2.count & ")"

  '------------remove all the males
    Set result = handler.remove("{'stuff.sex':'male'}")
    assert Array(result.handleCode >= 0), _
           result.response, "remove the males(" & result.count & ")"
           
   '-----------make sure they are gone
    Set result = handler.query()
    x = 0
    For Each job In testData.children
        x = x + -1 * CDbl(job.child("stuff.sex").value <> "male")
    Next job
    assert Array(result.handleCode >= 0, result.handleKeys.children.count = x), _
           result.response, "check after delete males(" & result.count & ")"

   '-----------add them back in
    Set result = handler.save(r2.data)
    assert Array(result.handleCode >= 0), _
           result.response, "add them back(" & result.count & ")"
           
    '--------check they got added
    Set result = handler.query("{'stuff.man':true}")
    x = 0
    For Each job In testData.children
        x = x + -1 * CDbl(job.child("stuff.sex").value = "male")
    Next job
    assert Array(result.handleCode >= 0, result.handleKeys.children.count = x), _
           result.response, "check after adding them back(" & result.count & ")"
    
    '-------sort and save
    Set result = handler.query(, "{'sort':'-serial'}")
    assert Array(result.handleCode >= 0, result.count = testData.children.count), _
            result.response, "sorting serial"
           
    '----- mark as good and save
    For Each job In result.data.children
        job.add "good", True
    Next job
    Set r2 = handler.save(result.data)
    assert Array(r2.handleCode >= 0), _
            r2.response, "adding goods"

    '-------check we have twice th records
    Set result = handler.count()
    assert Array(result.handleCode >= 0, result.count = testData.children.count * 2), _
            result.response, "doubled data"
            
    '------delete the ones we added
    Set result = handler.remove("{'good':true}")
    assert Array(result.handleCode >= 0), _
            result.response, "doubled data"

    '------check original length
    Set result = handler.count()
    assert Array(result.handleCode >= 0, result.count = testData.children.count), _
            result.response, "check final count"
            
    Debug.Print "Finished " & handler.getDbName
End Sub


  


