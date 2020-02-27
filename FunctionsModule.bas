Attribute VB_Name = "FunctionsModule"
Option Explicit 'This Module is specifically reserved for Functions

Declare PtrSafe Function WNetGetUser Lib "mpr.dll" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long 'Used by GetUserName
Const NoError = 0 'The Function call was successful

Function AdjustUserName(UserName As String)
    
    If UserName = "penny.b.cummings" Then
        AdjustUserName = "Penny.A.Barton"
    Else
        AdjustUserName = UserName
    End If
    
End Function

Function CopyVal(CR As String) 'Copies the value of the cell only. Use if the cell has a formula in it.
  Range(CR).Select
  Selection.Copy
  Range(CR).Select
  Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
End Function

Function GetEmailAddressByBusinessUnit(businessUnit As String) As String
    
    Select Case businessUnit
    
        Case "Atlanta"
            GetEmailAddressByBusinessUnit = "NoctalkATL@QuestDiagnostics.com"
        Case "Miami", "Florida"
            GetEmailAddressByBusinessUnit = "NoctalkTampa@QuestDiagnostics.com"
        Case "Auburn Hills", "Cincinnati", "Lenexa", "Wood Dale"
            GetEmailAddressByBusinessUnit = "NoctalkMidwest@QuestDiagnostics.com"
        Case "Greensboro" 'What Business Unit is Greensboro?
            GetEmailAddressByBusinessUnit = "NCSTalkGreensboro@QuestDiagnostics.com"
        Case "Cambridge", "Wallingford"
            GetEmailAddressByBusinessUnit = "NoctalkNEL@QuestDiagnostics.com"
        Case Else
            GetEmailAddressByBusinessUnit = ""
            
    End Select
    
End Function

Function GetEmailAddressByLaboratory(laboratory As String) As String
    
    Select Case laboratory
    
        Case "Albuquerque"
            GetEmailAddressByLaboratory = "DGXDALPROCESSING@questdiagnostics.com"
        Case "Atlanta"
            GetEmailAddressByLaboratory = ""
        Case "Auburn Hills"
            GetEmailAddressByLaboratory = ""
        Case "Baltimore"
            GetEmailAddressByLaboratory = ""
        Case "Cincinnati"
            GetEmailAddressByLaboratory = ""
        Case "Dallas"
            GetEmailAddressByLaboratory = "DGXDALPROCESSING@questdiagnostics.com"
        Case "Denver"
            GetEmailAddressByLaboratory = ""
        Case "DLO"
            GetEmailAddressByLaboratory = ""
        Case "Greensboro"
            GetEmailAddressByLaboratory = ""
        Case "Houston"
            GetEmailAddressByLaboratory = "DGXHOUPROCESSING@questdiagnostics.com"
        Case "Las Vegas"
            GetEmailAddressByLaboratory = ""
        Case "Lenexa"
            GetEmailAddressByLaboratory = ""
        Case "MACL"
            GetEmailAddressByLaboratory = ""
        Case "Marlborough"
            GetEmailAddressByLaboratory = ""
        Case "Miami"
            GetEmailAddressByLaboratory = ""
        Case "New Orleans"
            GetEmailAddressByLaboratory = ""
        Case "Philadelphia"
            GetEmailAddressByLaboratory = ""
        Case "Pittsburgh"
            GetEmailAddressByLaboratory = ""
        Case "Puerto Rico"
            GetEmailAddressByLaboratory = ""
        Case "Sacramento"
            GetEmailAddressByLaboratory = ""
        Case "Seattle"
            GetEmailAddressByLaboratory = ""
        Case "Solstas"
            GetEmailAddressByLaboratory = ""
        Case "Syosset"
            GetEmailAddressByLaboratory = ""
        Case "Tampa"
            GetEmailAddressByLaboratory = ""
        Case "Teterboro"
            GetEmailAddressByLaboratory = ""
        Case "Wallingford"
            GetEmailAddressByLaboratory = ""
        Case "West Hills"
            GetEmailAddressByLaboratory = ""
        Case "Wood Dale"
            GetEmailAddressByLaboratory = ""
        Case Else
            GetEmailAddressByLaboratory = ""
        
    End Select
    
End Function

Function RegionByBusinessUnit(businessUnit As String) As String

    Select Case businessUnit
    
        Case "Baltimore", "Philadelphia", "Syosset", "Teterboro"
            Region = "East"
            
        Case "Auburn Hills", "Cincinnati", "Wood Dale"
            Region = "Great Lakes"
            
        Case "Denver", "Lenexa"
            Region = "Midwest"
            
        Case "MACL"
            Region = "MACL"
        
        Case "Marlborough", "Pittsburgh", "Wallingford"
            Region = "North"
            
        Case "Puerto Rico"
            Region = "Puerto Rico"
            
        Case "Atlanta", "Solstas"
            Region = "South"
        
        Case "Miami", "Tampa"
            Region = "Southeast"
            
        Case "Albuquerque", "Dallas", "DLO", "Houston", "New Orleans"
            Region = "Southwest"
            
        Case "Las Vegas", "Sacramento", "Seattle", "West Hills"
            Region = "West"
            
        Case Else
            Region = ""
            
    End Select

End Function

Function GetUserName()

    Const lpnLength As Integer = 255 'Buffer size for the return string.
    
    Dim status As Integer 'Get return buffer space.
    Dim lpName, lpUserName As String 'For getting user information.

    lpUserName = Space$(lpnLength + 1) 'Assign the buffer size constant to lpUserName.
    status = WNetGetUser(lpName, lpUserName, lpnLength) 'Get the log-on name of the person using product.
    
    ' See whether error occurred.
    If status = NoError Then
       ' This line removes the null character. Strings in C are null-
       ' terminated. Strings in Visual Basic are not null-terminated.
       ' The null character must be removed from the C strings to be used
       ' cleanly in Visual Basic.
       lpUserName = Left$(lpUserName, InStr(lpUserName, Chr(0)) - 1)
    Else
       MsgBox "Unable to get the name." 'An error occurred.
       End
    End If
    
    lpUserName = AdjustUserName(lpUserName)

    GetUserName = lpUserName 'Display the name of the person logged on to the machine.

End Function

Function RangetoHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2016
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         fileName:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function


