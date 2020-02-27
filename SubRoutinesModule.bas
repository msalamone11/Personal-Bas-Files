Attribute VB_Name = "SubRoutinesModule"
Option Explicit

Sub ActivateWebFormField(WebDocument As HTMLDocument, WebFormFieldId As String)
    
    WebDocument.getElementById(WebFormFieldId).disabled = False 'This will make the field editable
    WebDocument.getElementById(WebFormFieldId).className = "text_field" 'Turns the field from Gray to white
    
End Sub

Sub AppendToATextFile(fileName As String, stringToAppendToTextFile As String)

    Dim strFileExists As String
    
    On Error Resume Next
        
        'fileName is the full file path of the file such as "\\QDCNS0002\TMP_Data$\TMPDept\Knowledge_Base\Logs\Specimen-In-Transit-Form\Specimen-In-Transit-Form-Log.txt"
        strFileExists = Dir(fileName) 'Dir(fileName) will return the file name if it exists. If file doesn't exist it will return Nothing.
        
        If strFileExists <> "" Then 'If strFileExists does not equal Nothing
            
            Open fileName For Append As #1
            Write #1, stringToAppendToTextFile
            Close #1
        
        End If

End Sub

Sub ClickIdTypeKeyStrokePressEnter(WebDocument As HTMLDocument, WebId As String, KeystrokesToType As String)

    WebDocument.getElementById(WebId).Click
    Application.SendKeys KeystrokesToType
    Application.SendKeys "{Enter}"
    Application.Wait Now + #12:00:01 AM# 'Wait 1 seconds

End Sub

Sub DeleteButtons()

    Sheets("Specimen In Transit Form").Select
    ActiveSheet.Shapes.Range(Array("GetFromQlsButton")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("ResetFormButton")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("SendFormButton")).Select
    Selection.Delete
    
End Sub

Sub GetFromQls()

    Dim UserName As String
    Dim FilePath As String 'C:\Users\" & userName & "\NCS-Automated-Forms\Specimen-In-Transit-Screen-Scraper\Specimen-In-Transit-Screen-Scraper.exe
    
    FilePath = ""
    UserName = GetUserName
    UserName = AdjustUserName(UserName) 'Some people's names in their C:\Users\%UserProfile% do not reflect what is returned by the GetUserName function.
        
    OpenAnyFile (FilePath)

End Sub

Sub HideEntireRangeRowAndOneBelow(RangeToHide As String)

    Range(RangeToHide).EntireRow.Hidden = True
    Range(RangeToHide).Offset(1, 0).EntireRow.Hidden = True

End Sub

Sub HideRanges(CommaSeparatedStringOfRanges As String)

    Dim arrWsNames() As String
    Dim Item As Variant
    
    arrWsNames = Split(CommaSeparatedStringOfRanges, ",")
    
    For Each Item In arrWsNames
        HideEntireRangeRowAndOneBelow (Item)
    Next
    
End Sub

Sub InputBoxThatExitsIfUserSelectsCancelForNamedRange(NamedRange As String, InputBoxPrompt As String, InputBoxTitle As String, Optional InputBoxDefaultPrompt As String)

    Range(NamedRange) = InputBox(InputBoxPrompt, InputBoxTitle, InputBoxDefaultPrompt)
    
    If Range(NamedRange) = "" Then
    
        End
        
    End If

End Sub

Sub MarkTheseRangesBlank(CommaSeparatedStringOfRanges As String)

    Dim arrWsNames() As String
    Dim Item As Variant
    
    arrWsNames = Split(CommaSeparatedStringOfRanges, ",")
    
    For Each Item In arrWsNames
        Range(Item) = ""
    Next

End Sub

Sub MarkTheseRangesBlankAndShow(CommaSeparatedStringOfRanges As String)

    Dim arrWsNames() As String
    Dim Item As Variant
    
    arrWsNames = Split(CommaSeparatedStringOfRanges, ",")
    
    For Each Item In arrWsNames
        Range(Item) = ""
        ShowEntireRangeRowAndOneBelow (Item)
    Next

End Sub

Sub MarkTheseRangesNa(CommaSeparatedListOfRangesToBeMarkedNA As String)

    Dim arrWsNames() As String
    Dim Item As Variant
    
    arrWsNames = Split(CommaSeparatedListOfRangesToBeMarkedNA, ",")
    
    For Each Item In arrWsNames
        Range(Item) = "N/A"
    Next

End Sub

Sub MarkTheseRangesNaAndHide(CommaSeparatedListOfRangesToMarkNaAndHide As String)

    Dim arrWsNames() As String
    Dim Item As Variant
    
    arrWsNames = Split(CommaSeparatedListOfRangesToMarkNaAndHide, ",")
    
    For Each Item In arrWsNames
        
        HideEntireRangeRowAndOneBelow (Item)
        Range(Item) = "N/A"
        
    Next

End Sub

Sub OpenAnyFile(FilePath As String)

    Dim fileX As Object
    
    Set fileX = CreateObject("Shell.Application")
    
    fileX.Open (FilePath)

End Sub

Sub ResetForm(rangesToClear As Variant)

    Application.ScreenUpdating = False
    
    Dim Item As Variant

    For Each Item In rangesToClear
        
        Range(Item) = ""
        
    Next
    
    'ActiveWindow.ScrollColumn = 1
    'ActiveWindow.ScrollRow = 1
    
    Application.ScreenUpdating = True
    
    'rangesToClear needs to be an Array of strings declared like this in another sub - Dim rangesToClear() As Variant

End Sub

Sub SendForm()
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False  'or True
      
    Dim FileExtStr As String
    Dim FileFormatNum As Long
    Dim Sourcewb As Workbook
    Dim Destwb As Workbook
    Dim TempFilePath As String
    Dim TempFileName As String
    Dim OutApp As Object
    Dim OutMail As Object
    Dim AccessionNumber As String
    Dim Body As String
    Dim DFax, DEmail, SEmail, Subj As String
    Dim sBody As Range
    Dim rng As Range
        
    Worksheets("Specimen In Transit Form").Unprotect Password:=""
    
    Set rng = Nothing
    Set rng = Range("EntireForm").SpecialCells(xlCellTypeVisible)
    
    Range("Date").Value = Format(Now(), "mm/dd/yyyy")
    Range("CsrName") = Application.WorksheetFunction.Proper(GetUserName)
    
    Set sBody = Nothing
    
    Sheets("Specimen In Transit Form").Select
    
    AccessionNumber = Range("AccessionNumber").Value
    
    Subj = AccessionNumber & " - Specimen In Transit Form"
    
    Sheets("Specimen In Transit Form").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    
    Set Sourcewb = ActiveWorkbook
    
    'Copy the sheet to a new workbook
    ActiveSheet.Copy
    Set Destwb = ActiveWorkbook
    
    'Call DeleteExtraText
    Call DeleteButtons

    'Set Sourcewb = ActiveWorkbook
    '
    '    'Copy the sheet to a new workbook
    '    ActiveSheet.Copy
    '    Set Destwb = ActiveWorkbook
    
    'Determine the Excel version and file extension/format
    With Destwb
    If Val(Application.Version) < 12 Then
        'You use Excel 2000-2003
        FileExtStr = ".xls": FileFormatNum = -4143
    Else
        FileExtStr = ".xlsx"
        FileFormatNum = 51
    
    End If
    End With

    'Save the new workbook/Mail it/Delete it
    TempFilePath = Environ$("temp") & "\"
    TempFileName = AccessionNumber & " - Specimen In Transit Form"
        
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    With Destwb
        .SaveAs TempFilePath & TempFileName & FileExtStr, _
                FileFormat:=FileFormatNum
        On Error Resume Next
        With OutMail
            .To = ""
            .CC = ""
            .BCC = ""
            .Subject = Subj
            .Attachments.Add Destwb.FullName
            'You can add other files also like this
            '.Attachments.Add ("C:\test.txt")
            '.Send
            .Display
            .HTMLBody = RangetoHTML(rng)
        End With
    End With
   
    Set OutMail = Nothing
    Set OutApp = Nothing
    
    Destwb.Close False
    
    Sourcewb.Activate
    
    Range("D9").Select
    Call ResetForm
    
    Worksheets("Specimen In Transit Form").Protect Password:=""
    
    'Application.Quit
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub

Sub ShowEntireRangeRowAndOneBelow(RangeToShow As String)

    Range(RangeToShow).EntireRow.Hidden = False
    Range(RangeToShow).Offset(1, 0).EntireRow.Hidden = False

End Sub

Sub ShowRanges(CommaSeparatedListOfRangesToShow As String)

    Dim arrWsNames() As String
    Dim Item As Variant
    
    arrWsNames = Split(CommaSeparatedListOfRangesToShow, ",")
    
    For Each Item In arrWsNames
        ShowEntireRangeRowAndOneBelow (Item)
    Next
    
End Sub

Sub ShowRangesAndMakeThemBlank(CommaSeparatedStringOfRanges As String)

    Dim arrWsNames() As String
    Dim Item As Variant
    
    arrWsNames = Split(CommaSeparatedStringOfRanges, ",")
    
    For Each Item In arrWsNames
        ShowEntireRangeRowAndOneBelow (Item)
        Range(Item) = ""
    Next
    
End Sub

Sub TipsMessageBox()

    Dim MsgBoxResponse As Integer

    MsgBoxResponse = MsgBox("Do you need tips to help identify or investigate the error that occurred?", vbYesNo, "Tips")
            
    If MsgBoxResponse = 6 Then
        
        MsgBox "Make sure you view the Req for 'Copy To' instructions." & vbNewLine & vbNewLine & _
            "In QLS, inside the Accession, check C2 line 18 for 'Copy To'" & vbNewLine & vbNewLine & _
            "View N for 'Electronically delivered list'", vbOKOnly, "Tips To Help Identify Or Investigate The Error"
        
    End If
    
End Sub

Sub ValidateBlankFormRanges(CommaSeparatedStringOfRanges As String)

    Dim arrWsNames() As String
    Dim Item As Variant
    
    arrWsNames = Split(CommaSeparatedStringOfRanges, ",")
    
    For Each Item In arrWsNames
        
        If Range(Item) = "" Then
        
            Range(Item).Select
            MsgBox ("You've missed a required field. Please fill in this field and then try again.")
            
            Debug.Print ("Caught a blank form field: " & Item) 'Use your Immediate window to see this output.
            
            End
        
        End If
        
    Next
    
    'Example of how to call this Sub: 'Call ValidateForm("CsrName,CallersName,AccountName,AccountNumber,AccessionNumberReqNumber," & _
                                                         "PatientsName,PatientsDob,Laboratory,TestName1,TestCode1")

End Sub
