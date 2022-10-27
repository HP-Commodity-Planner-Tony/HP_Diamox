Attribute VB_Name = "A_Home"
Sub Mode_Selector()

StartTime = Timer

If ThisWorkbook.Worksheets("Home").Cells(7, 3) = "A. Import DSP (B/S view)" Then

    Call Z_General.Z1_DeleteWorksheet
    Call B_Mode_Selector.B1_ImportDSP_BSview

ElseIf ThisWorkbook.Worksheets("Home").Cells(7, 3) = "B. WIF-DSP Comparison for BU IDM" Then
    
    Call Z_General.Z1_DeleteWorksheet
    Call B_Mode_Selector.B2_L10Comparison

ElseIf ThisWorkbook.Worksheets("Home").Cells(7, 3) = "C. Coppsheet-DSP Comparison for DT FIR" Then
    
    Call Z_General.Z1_DeleteWorksheet
    Call B_Mode_Selector.B3_FIRComparison
  
ElseIf ThisWorkbook.Worksheets("Home").Cells(7, 3) = "D. WOW WIF Comparison for BU IDM" Then
    
    Call Z_General.Z1_DeleteWorksheet
    Call B_Mode_Selector.B4_WOWWIF
    
End If


SecondsElapsed = Round(Timer - StartTime, 2)
MsgBox "This code completed in " & SecondsElapsed & " seconds"

End Sub
Sub Modifier_Selector()

StartTime = Timer

If ThisWorkbook.Worksheets("Home").Cells(12, 3) = "a. Copy Schenker FD to FD Override" Then
    
    If E_ErrorCheck.E1_CheckCalculation = False Then Exit Sub
    Call Y_Modifier_Selector.y1_Import_SchenkerFD

ElseIf ThisWorkbook.Worksheets("Home").Cells(12, 3) = "b. Modify Virtual Supply from Modifier Table" Then
    
    If E_ErrorCheck.E1_CheckCalculation = False Then Exit Sub
    Call Y_Modifier_Selector.y2_VSCalculation
    
ElseIf ThisWorkbook.Worksheets("Home").Cells(12, 3) = "c. Modify FD Override from Modifier Table" Then
    
    If E_ErrorCheck.E1_CheckCalculation = False Then Exit Sub
    Call Y_Modifier_Selector.y3_FDCalculation
    
ElseIf ThisWorkbook.Worksheets("Home").Cells(12, 3) = "d. Keep Prior WIF less consumption" Then
    
    If E_ErrorCheck.E1_CheckCalculation = False Then Exit Sub
    Call Y_Modifier_Selector.y4_keep_prior_wif

ElseIf ThisWorkbook.Worksheets("Home").Cells(12, 3) = "e. Copy to FDOverride from FD Table" Then
    
    If E_ErrorCheck.E1_CheckCalculation = False Then Exit Sub
    Call Y_Modifier_Selector.y5_FDtable_FDO

ElseIf ThisWorkbook.Worksheets("Home").Cells(12, 3) = "f. Copy to Virtual Supply from FD Table" Then
    
    If E_ErrorCheck.E1_CheckCalculation = False Then Exit Sub
    Call Y_Modifier_Selector.y6_FDtable_VS
    
ElseIf ThisWorkbook.Worksheets("Home").Cells(12, 3) = "g. CRP fair share from CRP Table" Then
    
    If E_ErrorCheck.E1_CheckCalculation = False Then Exit Sub
    Call Y_Modifier_Selector.y7_CRP_fair_share

ElseIf ThisWorkbook.Worksheets("Home").Cells(12, 3) = "h. Keep Previous Commit" Then
    
    If E_ErrorCheck.E1_CheckCalculation = False Then Exit Sub
    Call Y_Modifier_Selector.y8_keep_previous_commit

End If

SecondsElapsed = Round(Timer - StartTime, 2)
MsgBox "This code completed in " & SecondsElapsed & " seconds"

End Sub

Sub Upload_Selector()


StartTime = Timer


If ThisWorkbook.Worksheets("Home").Cells(17, 3) = "1. Create CSV (B/S)" Then
    
    If E_ErrorCheck.E1_CheckCalculation = False Then Exit Sub
    Call X_Upload_Selector.x1_Create_csv_calculation
 
ElseIf ThisWorkbook.Worksheets("Home").Cells(17, 3) = "2. Copy FDoverride and Virtual Supply to imported DSP (B/S)" Then

    Call X_Upload_Selector.x4_Copy_FDOverride_to_DSP

ElseIf ThisWorkbook.Worksheets("Home").Cells(17, 3) = "3. DSP lookup - L10 BUIDM (WIF)" Then

    Call X_Upload_Selector.x3_L10_lookup

ElseIf ThisWorkbook.Worksheets("Home").Cells(17, 3) = "4. DSP lookup - L10 RCTO (WIF)" Then
    
    Call X_Upload_Selector.x2_RCTO_lookup

ElseIf ThisWorkbook.Worksheets("Home").Cells(17, 3) = "5. Copy FIR commitment to DSP (FIR)" Then

    Call X_Upload_Selector.x5_FIR_Copp_to_DSP

ElseIf ThisWorkbook.Worksheets("Home").Cells(17, 3) = "6. Copy WIF Calculation to DSP (BUIDM)" Then

    Call X_Upload_Selector.x6_L10_lookup_Cal

End If


SecondsElapsed = Round(Timer - StartTime, 2)
MsgBox "This code completed in " & SecondsElapsed & " seconds"

End Sub

Sub Function_Selector()

If ThisWorkbook.Worksheets("Home").Cells(7, 8) = "Turn On DateChecker (Default)" Then

    Call Z_General.Z2_DateChecker
    MsgBox "Date checker has been enabled"

ElseIf ThisWorkbook.Worksheets("Home").Cells(7, 8) = "Turn Off DateChecker" Then

    Call Z_General.Z2_DateChecker
    MsgBox "Date checker has been disabled"

End If


End Sub

Sub Instruction_Selector()


StartTime = Timer


If ThisWorkbook.Worksheets("Home").Cells(12, 8) = "Show Explanation" Then

    ThisWorkbook.Worksheets("Explanation").Visible = True

ElseIf ThisWorkbook.Worksheets("Home").Cells(12, 8) = "I am Developer" Then

    ThisWorkbook.Worksheets("Developer").Visible = True

End If


SecondsElapsed = Round(Timer - StartTime, 2)
MsgBox "This code completed in " & SecondsElapsed & " seconds"

End Sub

