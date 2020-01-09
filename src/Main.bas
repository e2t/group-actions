Attribute VB_Name = "Main"
'Written in 2015 by Eduard E. Tikhenko <aquaried@gmail.com>
'
'To the extent possible under law, the author(s) have dedicated all copyright
'and related and neighboring rights to this software to the public domain
'worldwide. This software is distributed without any warranty.
'You should have received a copy of the CC0 Public Domain Dedication along
'with this software.
'If not, see <http://creativecommons.org/publicdomain/zero/1.0/>

Option Explicit

Type IAction
    rebuild As Boolean
    save As Boolean
    close As Boolean
End Type

Type ISubject
    drawings As Boolean
    parts As Boolean
    assemblies As Boolean
End Type

Dim swApp As Object
Dim documentsCount As Integer
Dim documentFormat As String

Sub Main()
    Set swApp = Application.SldWorks
    Form.Show
End Sub

Function GetAction() As IAction
    With GetAction
        .close = Form.closeBox.value
        .rebuild = Form.rebuildBox.value
        .save = Form.saveBox.value
    End With
End Function

Function GetSubject() As ISubject
    With GetSubject
        .assemblies = Form.asmBox.value
        .drawings = Form.drawBox.value
        .parts = Form.partBox.value
    End With
End Function

Function IsAnySubject(subject As ISubject) As Boolean
    IsAnySubject = subject.assemblies Or subject.drawings Or subject.parts
End Function

Sub StatusInc(index As Integer)
    Form.statusLab.Caption = Format(index, documentFormat) & "/" & documentsCount
    Form.Repaint
End Sub

Function Execute() 'mask for button
    Dim action As IAction
    Dim subject As ISubject
    Dim i As Integer
    Dim doc_ As Variant
    Dim doc As ModelDoc2
    
    subject = GetSubject
    If Not IsAnySubject(subject) Then Exit Function
    
    documentsCount = swApp.GetDocumentCount
    If documentsCount > 0 Then
        documentFormat = ""
        For i = 1 To Int(Log(documentsCount) / Log(10)) + 1
            documentFormat = documentFormat & "0"
        Next
    Else
        Exit Function
    End If
    
    action = GetAction
    
    If action.rebuild Then
        i = 0
        For Each doc_ In swApp.GetDocuments
            i = i + 1
            StatusInc i
            Set doc = doc_
            If IsMaking(doc, subject) Then RebuildThis doc
        Next
    End If
        
    If action.save Then
        i = 0
        For Each doc_ In swApp.GetDocuments
            i = i + 1
            StatusInc i
            Set doc = doc_
            If IsMaking(doc, subject) Then SaveThis doc
        Next
    End If
        
    If action.close Then
        i = 0
        For Each doc_ In swApp.GetDocuments
            i = i + 1
            StatusInc i
            Set doc = doc_
            CloseThisDocIfNeed doc, subject
        Next
    End If
    
End Function

Function IsMaking(doc As ModelDoc2, subject As ISubject) As Boolean
    IsMaking = (subject.assemblies And doc.GetType = swDocASSEMBLY) _
               Or (subject.drawings And doc.GetType = swDocDRAWING) _
               Or (subject.parts And doc.GetType = swDocPART)
End Function

Sub RebuildThis(doc As ModelDoc2)
    doc.ForceRebuild3 True
End Sub

Sub SaveThis(doc As ModelDoc2)
    Dim errors As swFileSaveError_e
    Dim warnings As swFileSaveWarning_e
    
    doc.SetReadOnlyState False  'must be first!
    doc.Save3 swSaveAsOptions_Silent, errors, warnings  ' AsBool if needed
End Sub

Sub CloseThisDocIfNeed(doc As ModelDoc2, subject As ISubject)
    On Error Resume Next
    If IsMaking(doc, subject) Then swApp.CloseDoc doc.GetPathName
End Sub
