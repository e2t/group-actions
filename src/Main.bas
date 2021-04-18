Attribute VB_Name = "Main"
Option Explicit

Type IAction
    Rebuild As Boolean
    Save As Boolean
    Close As Boolean
    DisableReadOnly As Boolean
    ReduceQuality As Boolean
    ChangeUnitsIPS As Boolean
    ChangeUnitsMmKg As Boolean
    MakeFastener As Boolean
End Type

Type ISubject
    Drawings As Boolean
    Parts As Boolean
    Assemblies As Boolean
    RootOnly As Boolean
End Type

Public gCurrentFolderPath As String

Dim swApp As Object
Dim gFSO As FileSystemObject
Dim documentsCount As Integer
Dim documentFormat As String

Sub Main()

  Dim сurrentDoc As ModelDoc2
  
  Set swApp = Application.SldWorks
  Set gFSO = New FileSystemObject
  Set сurrentDoc = swApp.ActiveDoc
  If Not сurrentDoc Is Nothing Then
    gCurrentFolderPath = LCase(gFSO.GetParentFolderName(сurrentDoc.GetPathName))
  End If
  Form.Show
  
End Sub

Function GetAction() As IAction

  GetAction.Close = Form.closeBox.value
  GetAction.Rebuild = Form.rebuildBox.value
  GetAction.Save = Form.saveBox.value
  GetAction.DisableReadOnly = Form.RoChk.value
  GetAction.ReduceQuality = Form.qualityBox.value
  GetAction.ChangeUnitsIPS = Form.IPSBox.value
  GetAction.ChangeUnitsMmKg = Form.MmKgBox.value
  GetAction.MakeFastener = Form.fastenerBox.value
  
End Function

Function GetSubject() As ISubject

  GetSubject.Assemblies = Form.asmBox.value
  GetSubject.Drawings = Form.drawBox.value
  GetSubject.Parts = Form.partBox.value
  GetSubject.RootOnly = Form.RootOnlyBox.value
  
End Function

Function IsAnySubject(Subject As ISubject) As Boolean

  IsAnySubject = Subject.Assemblies Or Subject.Drawings Or Subject.Parts
  
End Function

Sub StatusInc(index As Integer)

  Form.statusLab.Caption = Format(index, documentFormat) & "/" & documentsCount
  Form.Repaint
  
End Sub

Function Execute() 'mask for button

  Dim Action As IAction
  Dim Subject As ISubject
  Dim i As Integer
  Dim doc_ As Variant
  Dim doc As ModelDoc2
  
  Subject = GetSubject
  If Not IsAnySubject(Subject) Then Exit Function
  
  documentsCount = swApp.GetDocumentCount
  If documentsCount > 0 Then
    documentFormat = ""
    For i = 1 To Int(Log(documentsCount) / Log(10)) + 1
      documentFormat = documentFormat & "0"
    Next
  Else
    Exit Function
  End If
  
  Action = GetAction
  
  If Action.DisableReadOnly Then
    i = 0
    For Each doc_ In swApp.GetDocuments
      i = i + 1
      StatusInc i
      Set doc = doc_
      If IsMaking(doc, Subject) Then DisableReadOnly doc
    Next
  End If
  
  If Action.ReduceQuality Then
    i = 0
    For Each doc_ In swApp.GetDocuments
      i = i + 1
      StatusInc i
      Set doc = doc_
      If IsMaking(doc, Subject) Then ChangeDocumentSettings doc.Extension
    Next
  End If
  
  If Action.ChangeUnitsIPS Then
    i = 0
    For Each doc_ In swApp.GetDocuments
      i = i + 1
      StatusInc i
      Set doc = doc_
      If IsMaking(doc, Subject) Then
        SetIPSUnits doc
        ShowDualDimensions doc
      End If
    Next
  End If
  
  If Action.ChangeUnitsMmKg Then
    i = 0
    For Each doc_ In swApp.GetDocuments
      i = i + 1
      StatusInc i
      Set doc = doc_
      If IsMaking(doc, Subject) Then
        SetMmKgUnits doc
        HideDualDimensions doc
      End If
    Next
  End If
  
  If Action.MakeFastener Then
    i = 0
    For Each doc_ In swApp.GetDocuments
      i = i + 1
      StatusInc i
      Set doc = doc_
      If IsMaking(doc, Subject) Then SetIsFastener doc.Extension.CustomPropertyManager("")
    Next
  End If
  
  If Action.Rebuild Then
    i = 0
    For Each doc_ In swApp.GetDocuments
      i = i + 1
      StatusInc i
      Set doc = doc_
      If IsMaking(doc, Subject) Then RebuildThis doc
    Next
  End If
      
  If Action.Save Then
    i = 0
    For Each doc_ In swApp.GetDocuments
      i = i + 1
      StatusInc i
      Set doc = doc_
      If IsMaking(doc, Subject) Then SaveThis doc
    Next
  End If
      
  If Action.Close Then
    i = 0
    For Each doc_ In swApp.GetDocuments
      i = i + 1
      StatusInc i
      Set doc = doc_
      CloseThisDocIfNeed doc, Subject
    Next
  End If
    
End Function

Function IsMaking(doc As ModelDoc2, Subject As ISubject) As Boolean
  
  Dim IsCorrectType, IsCorrectPath As Boolean

  IsCorrectType = (Subject.Assemblies And doc.GetType = swDocASSEMBLY) _
               Or (Subject.Drawings And doc.GetType = swDocDRAWING) _
               Or (Subject.Parts And doc.GetType = swDocPART)
  If Subject.RootOnly Then
    IsCorrectPath = LCase(doc.GetPathName) Like gCurrentFolderPath + "\*"
  Else
    IsCorrectPath = True
  End If
  IsMaking = IsCorrectType And IsCorrectPath
  
End Function

Sub RebuildThis(doc As ModelDoc2)

  doc.ForceRebuild3 True
  
End Sub

Sub DisableReadOnly(doc As ModelDoc2)

  doc.SetReadOnlyState False
  
End Sub

Sub SaveThis(doc As ModelDoc2)

  Dim errors As swFileSaveError_e
  Dim warnings As swFileSaveWarning_e
  
  doc.Save3 swSaveAsOptions_Silent, errors, warnings  ' AsBool if needed
  
End Sub

Sub CloseThisDocIfNeed(doc As ModelDoc2, Subject As ISubject)

  On Error Resume Next
  If IsMaking(doc, Subject) Then swApp.CloseDoc doc.GetPathName
  
End Sub

Function ChangeDocumentSettings(modelExt As ModelDocExtension) 'masked

  '''Image Quality
  
  'Shaded and draft quality HLR/HLV resolution - Quality
  modelExt.SetUserPreferenceInteger swUserPreferenceIntegerValue_e.swImageQualityShaded, _
                                    swUserPreferenceOption_e.swDetailingNoOptionSpecified, _
                                    swImageQualityShaded_e.swShadedImageQualityCoarse
                                    
  'Shaded and draft quality HLR/HLV resolution - Optimize edge length (higher quality, but slower
  modelExt.SetUserPreferenceToggle swUserPreferenceToggle_e.swImageQualityUseHighQualityEdgeSize, _
                                   swUserPreferenceOption_e.swDetailingNoOptionSpecified, False
                                   
  'Wireframe and high quality HLR/HLV resolution - quality of wireframe display output
  modelExt.SetUserPreferenceInteger swUserPreferenceIntegerValue_e.swImageQualityWireframe, _
                                    swUserPreferenceOption_e.swDetailingNoOptionSpecified, _
                                    swImageQualityWireframe_e.swWireframeImageQualityOptimal
                                    
  'Wireframe and high quality HLR/HLV resolution - Precisely render overlapping geometry (higher quality, but slower)
  modelExt.SetUserPreferenceToggle swUserPreferenceToggle_e.swPreciseRenderingOfOverlappingGeometry, _
                                   swUserPreferenceOption_e.swDetailingNoOptionSpecified, False
                                   
End Function

Sub ShowDualDimensions(doc As ModelDoc2)

  doc.Extension.SetUserPreferenceToggle swDetailingDualDimensions, swDetailingDimension, True
  doc.Extension.SetUserPreferenceToggle swDetailingShowUnitsForDualDisplay, swDetailingDimension, True
  doc.Extension.SetUserPreferenceInteger swDetailingDualDimPosition, swDetailingDimension, swDualDimensionsOnRight
    
End Sub

Sub HideDualDimensions(doc As ModelDoc2)

  doc.Extension.SetUserPreferenceToggle swDetailingDualDimensions, swDetailingDimension, False
    
End Sub

Sub SetIPSUnits(doc As ModelDoc2) 'mask for button

  '''Глобальная предустановка
  doc.Extension.SetUserPreferenceInteger swUnitSystem, swDetailingNoOptionSpecified, swUnitSystem_IPS

  '''Основные единицы длины
  'doc.Extension.SetUserPreferenceInteger swUnitsLinear, swDetailingNoOptionSpecified, swMM
  doc.Extension.SetUserPreferenceInteger swUnitsLinearDecimalDisplay, swDetailingNoOptionSpecified, swFRACTION
  doc.Extension.SetUserPreferenceInteger swUnitsLinearDecimalPlaces, swDetailingNoOptionSpecified, 4
  doc.Extension.SetUserPreferenceInteger swUnitsLinearFractionDenominator, swDetailingNoOptionSpecified, 64
  doc.Extension.SetUserPreferenceInteger swUnitsLinearFeetAndInchesFormat, swDetailingNoOptionSpecified, False
  'doc.Extension.SetUserPreferenceInteger swUnitsLinearRoundToNearestFraction, swDetailingNoOptionSpecified, True
  
  '''Двойные единицы длины
  doc.Extension.SetUserPreferenceInteger swUnitsDualLinear, swDetailingNoOptionSpecified, swMM
  doc.Extension.SetUserPreferenceInteger swUnitsDualLinearDecimalDisplay, swDetailingNoOptionSpecified, swDECIMAL
  doc.Extension.SetUserPreferenceInteger swUnitsDualLinearDecimalPlaces, swDetailingNoOptionSpecified, 2
  'doc.Extension.SetUserPreferenceInteger swUnitsDualLinearFractionDenominator, swDetailingNoOptionSpecified, 64
  'doc.Extension.SetUserPreferenceInteger swUnitsDualLinearRoundToNearestFraction, swDetailingNoOptionSpecified, True
  'doc.Extension.SetUserPreferenceInteger swUnitsDualLinearFeetAndInchesFormat, swDetailingNoOptionSpecified, False
  
  '''Угловые единицы
  doc.Extension.SetUserPreferenceInteger swUnitsAngular, swDetailingNoOptionSpecified, swDEGREES
  doc.Extension.SetUserPreferenceInteger swUnitsAngularDecimalPlaces, swDetailingNoOptionSpecified, 2
  
  '''Единицы массы
  'doc.Extension.SetUserPreferenceInteger swUnitsMassPropLength, swDetailingNoOptionSpecified, swMM
  doc.Extension.SetUserPreferenceInteger swUnitsMassPropDecimalPlaces, swDetailingNoOptionSpecified, 2
  'doc.Extension.SetUserPreferenceInteger swUnitsMassPropMass, swDetailingNoOptionSpecified, swUnitsMassPropMass_Kilograms
  'doc.Extension.SetUserPreferenceInteger swUnitsMassPropVolume, swDetailingNoOptionSpecified, swUnitsMassPropVolume_Meters3
    
End Sub

Sub SetMmKgUnits(doc As ModelDoc2) 'mask for button

  '''Глобальная предустановка
  doc.Extension.SetUserPreferenceInteger swUnitSystem, swDetailingNoOptionSpecified, swUnitSystem_Custom

  '''Основные единицы длины
  doc.Extension.SetUserPreferenceInteger swUnitsLinear, swDetailingNoOptionSpecified, swMM
  doc.Extension.SetUserPreferenceInteger swUnitsLinearDecimalDisplay, swDetailingNoOptionSpecified, swDECIMAL
  doc.Extension.SetUserPreferenceInteger swUnitsLinearDecimalPlaces, swDetailingNoOptionSpecified, 2
  'doc.Extension.SetUserPreferenceInteger swUnitsLinearFractionDenominator, swDetailingNoOptionSpecified, 64
  'doc.Extension.SetUserPreferenceInteger swUnitsLinearFeetAndInchesFormat, swDetailingNoOptionSpecified, False
  'doc.Extension.SetUserPreferenceInteger swUnitsLinearRoundToNearestFraction, swDetailingNoOptionSpecified, True
  
  '''Двойные единицы длины
  doc.Extension.SetUserPreferenceInteger swUnitsDualLinear, swDetailingNoOptionSpecified, swINCHES
  doc.Extension.SetUserPreferenceInteger swUnitsDualLinearDecimalDisplay, swDetailingNoOptionSpecified, swFRACTION
  doc.Extension.SetUserPreferenceInteger swUnitsDualLinearDecimalPlaces, swDetailingNoOptionSpecified, 4
  doc.Extension.SetUserPreferenceInteger swUnitsDualLinearFractionDenominator, swDetailingNoOptionSpecified, 64
  doc.Extension.SetUserPreferenceInteger swUnitsDualLinearRoundToNearestFraction, swDetailingNoOptionSpecified, False
  'doc.Extension.SetUserPreferenceInteger swUnitsDualLinearFeetAndInchesFormat, swDetailingNoOptionSpecified, False
  
  '''Угловые единицы
  doc.Extension.SetUserPreferenceInteger swUnitsAngular, swDetailingNoOptionSpecified, swDEGREES
  doc.Extension.SetUserPreferenceInteger swUnitsAngularDecimalPlaces, swDetailingNoOptionSpecified, 2
  
  '''Единицы массы
  doc.Extension.SetUserPreferenceInteger swUnitsMassPropLength, swDetailingNoOptionSpecified, swMM
  doc.Extension.SetUserPreferenceInteger swUnitsMassPropDecimalPlaces, swDetailingNoOptionSpecified, 2
  doc.Extension.SetUserPreferenceInteger swUnitsMassPropMass, swDetailingNoOptionSpecified, swUnitsMassPropMass_Kilograms
  doc.Extension.SetUserPreferenceInteger swUnitsMassPropVolume, swDetailingNoOptionSpecified, swUnitsMassPropVolume_Meters3
    
End Sub

Sub SetIsFastener(mgr As CustomPropertyManager)

  Const pIsFastener = "IsFastener"
  
  mgr.Delete2 pIsFastener
  mgr.Add3 pIsFastener, swCustomInfoNumber, "1", swCustomPropertyDeleteAndAdd

End Sub
