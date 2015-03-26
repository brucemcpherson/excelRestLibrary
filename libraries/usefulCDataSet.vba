Option Explicit
' v0.1 23.3.15
Public Function makeSheetFromJob(job As cJobject, sheetName As String) As cDataSet
    Dim ds As cDataSet, target As Range, dc As cCell, jo As cJobject


    ' clear the target sheet
    Set ds = New cDataSet
    
    ' need something there to load
    Set target = Sheets(sheetName).Range("a1")
    If IsEmpty(target.value) Then
        target.value = "dummy"
    End If
    ds.load target.Worksheet.name
    
    ' create headings based on all data found
    makeSheetHeadingsFromJob job, ds
    ds.tearDown
    
    ' now reload with new headings
    Set ds = New cDataSet
    ds.load (target.Worksheet.name)
    
    ' now populate the data
    With ds.headingRow
        For Each jo In job.children
            For Each dc In .headings
                If (isSomething(jo.child(dc.value))) Then
                    .where.Resize(1, 1).Offset(jo.childIndex, dc.column - 1).value = jo.child(dc.value).value
                End If
            Next dc
            If (jo.childIndex Mod 1000 = 0) Then
                Debug.Print "done "; jo.childIndex; " rows"
            End If
        Next jo
    End With
    
    ' clean
    Dim dsNew As cDataSet
    Set dsNew = New cDataSet
    dsNew.load (ds.name)
    ds.tearDown
    Set makeSheetFromJob = dsNew
    
End Function
Public Sub makeSheetHeadingsFromJob(jo As cJobject, ds As cDataSet)
    Dim jobHead As cJobject, job As cJobject, joc As cJobject, jod As cJobject
    Set jobHead = New cJobject
    

    ' first step, identify the headings
    ' this will also take care of situation when each item doesnt have the same children
    Set jobHead = jobHead.init(Nothing)
    For Each job In jo.children
        Set jobHead = rescurseSheetHeadersFromJob(job, jobHead)
    Next job
    
    ' let's clear all existing
    If (isSomething(ds.where)) Then
        ds.where.ClearContents
    End If

    ds.headingRow.where.ClearContents
    
    ' now the heading
    With firstCell(ds.headingRow.where)
        For Each job In jobHead.children
            .Offset(, job.childIndex - 1).value = Replace(job.key, "___", ".")
        Next job
    End With

    jobHead.tearDown
End Sub

Private Function rescurseSheetHeadersFromJob(job As cJobject, _
            jobHead As cJobject, Optional k As String = vbNullString) As cJobject
    Dim joc As cJobject, s As String

    ' the trick here is to collapse to a single depth- we'll replace the underscores with . later
    If job.hasChildren Then
        If k <> vbNullString Then k = k + "___"
        For Each joc In job.children
            rescurseSheetHeadersFromJob joc, jobHead, k + joc.key
        Next joc
    Else
        If k = vbNullString Then k = job.key
        If (Not IsEmpty(job.value)) Then
            jobHead.add k
        End If
    End If
    
    Set rescurseSheetHeadersFromJob = jobHead
End Function

Private Function addD3TreeItem(meOb As cJobject, ds As cDataSet, label As String, key As String, parentkey As String, _
    Optional drd As cDataRow = Nothing) As cJobject
    Dim cj As cJobject, dr As cDataRow, cc As cCell
    ' does parent key exist?
    Set cj = meOb.find(parentkey)
    If (cj Is Nothing) Then
        Set dr = findD3Parent(ds, parentkey)
        If Not dr Is Nothing Then
            Set cj = addD3TreeItem(meOb, ds, label, parentkey, cleanDot(dr.cell("Parent key").toString), dr)
        End If
    End If
    If cj Is Nothing Then
        MsgBox ("could not find " & key & " " & parentkey)
    Else
        With cj.add(key)
            .add "label", label
            ' anything else on this row?
            If Not drd Is Nothing Then
                For Each cc In drd.columns
                    If (cc.myKey <> "key" And cc.myKey <> "label" And _
                        cc.myKey <> "parent key" And Not IsEmpty(cc.value)) Then
                        .add cc.myKey, cc.value
                    End If
                Next cc
            End If
        End With
    End If
    Set addD3TreeItem = cj
End Function
Private Function findD3Parent(ds As cDataSet, parentkey) As cDataRow
    Dim dr As cDataRow
    For Each dr In ds.rows
        If cleanDot(dr.cell("key").toString) = parentkey Then
            Set findD3Parent = dr
            Exit Function
        End If
    Next dr
    
End Function
Private Function cleanDot(s As String) As String
    '. has special meaning for cJobject so if present in key, then remove
    cleanDot = makeKey(Replace(s, ".", "_ _"))
End Function
Public Function makeD3Tree(meOb As cJobject, ds As cDataSet, dsOptions As cDataSet, Optional options As String = "options") As cJobject
    ' this one will take a list of Name/Parents and make a structured cJobject out of it
    Dim dr As cDataRow, cj As cJobject, parent As String, name As String, c3 As cJobject, ct As cJobject, t As cJobject
    Const container = "contents"
    If Not ds.headingRow.validate(True, "Label", "Parent Key", "Key") Then Exit Function
    Set cj = meOb.add("D3Root")
    
    For Each dr In ds.rows
        Set ct = addD3TreeItem(cj, ds, _
            dr.cell("label").toString, _
            cleanDot(dr.cell("key").toString), _
            cleanDot(dr.cell("Parent key").toString), dr)
    Next dr
    ' now lets tweak that to a d3 format
    Set c3 = New cJobject
    
    With c3.init(Nothing)
        ' add an options branch
        With .add("options")
            For Each dr In dsOptions.rows
                If dr.cell("value").toString <> vbNullString Then
                    .add dr.cell(options).toString, _
                            dr.cell("value").toString
                End If
            Next dr
        End With
        
        
        ' add a branch for data
        Set t = .add("data")
        t.add "label", dsOptions.cell("root", "value").toString
        makeD3 t, cj.children(1)
       
    End With
    Set makeD3Tree = c3
End Function
Private Function makeD3(meOb As cJobject, cj As cJobject) As cJobject
    Dim cjc As cJobject, t As cJobject

    If cj.hasChildren Then
        Set t = meOb.add("children").addArray.add
        For Each cjc In cj.children
            makeD3 t, cjc
        Next cjc
       
    Else
        meOb.add cj.key, cj.value
    End If
    
    Set makeD3 = meOb
End Function




