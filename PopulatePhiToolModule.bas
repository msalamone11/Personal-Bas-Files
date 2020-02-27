Attribute VB_Name = "PopulatePhiToolModule"
'This Module is specifically reserved for populating data to the PHITT Tool

Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdSHow As Long) As Long 'Allows you to maximize the Internet Explorer window

Sub PopulatePhiTool()

    Const SW_SHOWMAXIMIZED = 3
    Dim IE As Object
    Dim doc As HTMLDocument
    
    Set IE = CreateObject("InternetExplorer.Application")
    
    Dim UserName As String
    Dim Password As String
    
    UserName = GetUserName
    Password = InputBoxDK("Enter your PHITT Tool Password", "PHITT Tool Password")
    'Password = InputBox("Enter your password")
    
    'Launch PHITT Website
    IE.Visible = True
    IE.navigate "http://apexdb.qdx.com:7777/pls/apexp/f?p=220:LOGIN:7605821446442072:::::"
    ShowWindow IE.hwnd, SW_SHOWMAXIMIZED 'Maximize The Window

    'Wait for the website to load all the way
    Do While IE.Busy
        Application.Wait DateAdd("s", 1, Now)
    Loop
    
    Set doc = IE.document

    doc.getElementById("P101_USERNAME").Value = UserName
    doc.getElementById("P101_PASSWORD").Value = Password
    
    Dim the_input_elements
    Set the_input_elements = doc.getElementsByClassName("questButton2")
    
    For Each input_element In the_input_elements
        
        If input_element.getAttribute("value") = "Login" Then
            input_element.Click
            Exit For
        End If
        
    Next
    
    'Wait for the website to load all the way
    
    Do While IE.Busy
        Application.Wait DateAdd("s", 1, Now)
    Loop
    
        'SelectYourRegion
    
    If Range("SelectYourRegion") <> "" Then
    
        Set SelectYourRegion = doc.getElementById("P1_BUSINESS_UNIT_DESCR")
        
        Select Case Range("SelectYourRegion").Value
        
            Case "East"
                Call ClickIdTypeKeyStrokePressEnter(doc, "P1_BUSINESS_UNIT_DESCR", "E")
            Case "Great Lakes"
                Call ClickIdTypeKeyStrokePressEnter(doc, "P1_BUSINESS_UNIT_DESCR", "G")
            Case "MACL"
                Call ClickIdTypeKeyStrokePressEnter(doc, "P1_BUSINESS_UNIT_DESCR", "M")
            Case "Midwest"
                Call ClickIdTypeKeyStrokePressEnter(doc, "P1_BUSINESS_UNIT_DESCR", "Midwest")
            Case "North"
                Call ClickIdTypeKeyStrokePressEnter(doc, "P1_BUSINESS_UNIT_DESCR", "N")
            Case "Puerto Rico"
                Call ClickIdTypeKeyStrokePressEnter(doc, "P1_BUSINESS_UNIT_DESCR", "P")
            Case "Southeast"
                Call ClickIdTypeKeyStrokePressEnter(doc, "P1_BUSINESS_UNIT_DESCR", "Southeast")
            Case "Southwest"
                Call ClickIdTypeKeyStrokePressEnter(doc, "P1_BUSINESS_UNIT_DESCR", "Southwest")
            Case "West"
                Call ClickIdTypeKeyStrokePressEnter(doc, "P1_BUSINESS_UNIT_DESCR", "W")
        End Select
        
    End If
    
    'AccessionNumber - "Invoice/Accession/Requisition" on PHITT Tool
    
    If Range("AccessionNumber") <> "" Then
    
        doc.getElementById("P1_DISCOVERED_DATA_TYPE_IDENT").Value = Range("AccessionNumber").Value & " / Problem Tracking Number: " & Range("ProblemTrackingNumber").Value
    
    End If
    
    'AddressLine1PhiWasMailedTo - Not needed on PHITT Tool
    
    'AskTheCallerForFirstAndLastName
    
    'CallersName - "Reported By Name" on the PHITT Tool
    
    doc.getElementById("P1_REPORTED_BY_NAME").Value = Range("CallersName")
    
    'CityPhiWasMailedTo - Not needed on PHITT Tool
    
    'ContactPreference
    
    Call ActivateWebFormField(doc, "P1_REPORTED_BY_CONTACT")
    
    If Range("ContactPreference").Value = "Phone Number" Then
        
        doc.getElementById("P1_REPORTED_BY_CONTACT_TYPE").Value = "PHONE"
        
    ElseIf Range("ContactPreference").Value = "Email Address" Then
        doc.getElementById("P1_REPORTED_BY_CONTACT_TYPE").Value = "EMAIL"
        
    ElseIf Range("ContactPreference").Value = "Phone and Email" Then
        doc.getElementById("P1_REPORTED_BY_CONTACT_TYPE").Value = "PHONE_EMAIL"
        
    ElseIf Range("ContactPreference").Value = "Unavailable" Then
        doc.getElementById("P1_REPORTED_BY_CONTACT_TYPE").Value = "UNAVAILABLE"
        
    ElseIf Range("ContactPreference").Value = "Refused" Then
        doc.getElementById("P1_REPORTED_BY_CONTACT_TYPE").Value = "REFUSED"
        
    ElseIf Range("ContactPreference").Value = "N/A" Then
        doc.getElementById("P1_REPORTED_BY_CONTACT_TYPE").Value = "0"
        
    End If
    
    'ContactPreference2
    
    If Range("ContactPreference") <> "N/A" Then
    
        If Range("DiscoveredBy") = "Business" Then
        
            doc.getElementById("P1_REPORTED_BY_CONTACT").Value = Range("ContactPreference2").Value & " -  " & _
                                                                                                                Range("BusinessName") & ", " & _
                                                                                                                Range("BusinessAddress") & ", " & _
                                                                                                                Range("BusinessCity") & ", " & _
                                                                                                                Range("BusinessState") & ", " & _
                                                                                                                Range("BusinessZipCode")
        
        Else:
        
            doc.getElementById("P1_REPORTED_BY_CONTACT").Value = Range("ContactPreference2").Value
        
        End If
    
        
    
    End If
    
    'DidCallerStateItWasHackerOrContractor - "PHI Accessed By An Unauthorized Entity, Contractor Or Hacker"
    
    If Range("DidCallerStateItWasHackerOrContractor") <> "N/A" Then
    
        If Range("DidCallerStateItWasHackerOrContractor") = "No" Then
        
            doc.getElementById("P1_CAT_UNAUTHORIZED_ACCESS_I").Value = "N"
            
        ElseIf Range("DidCallerStateItWasHackerOrContractor") = "Yes" Then
        
            doc.getElementById("P1_CAT_UNAUTHORIZED_ACCESS_I").Value = "Y"
        
        End If
        
    End If
    
    'DidTheyConfirmThePhiWasShredded
    
    'DidYouAskTheCallerToReturnDestroyOrRemovePHI - "Was PHI Returned, Destroyed, Or Removed?" on PHITT Tool
    
    If Range("DidYouAskTheCallerToReturnDestroyOrRemovePHI").Value = "Yes" Then
    
        doc.getElementById("P1_PHI_RETURNED_DESTROYED_I").Value = "Y"
        
    ElseIf Range("DidYouAskTheCallerToReturnDestroyOrRemovePHI").Value = "No" Then
    
        doc.getElementById("P1_PHI_RETURNED_DESTROYED_I").Value = "N"
        
    End If
    
    'DiscoveredBy
    
    If Range("DiscoveredBy") <> "" And Range("DiscoveredBy") <> "N/A" Then
    
        Select Case Range("DiscoveredBy").Value
        
            Case "Applicant"
                doc.getElementById("P1_DISCOVERED_BY_ROLE_CD").Value = "APPLCNT"
            Case "Business"
                doc.getElementById("P1_DISCOVERED_BY_ROLE_CD").Value = "BUSINESS"
                doc.getElementById("P1_DISCOVERED_BY_NAME").Value = Range("BusinessName").Value & ", " & _
                                                                                                                  Range("BusinessAddress").Value & ", " & _
                                                                                                                  Range("BusinessCity").Value & ", " & _
                                                                                                                  Range("BusinessState").Value & ", " & _
                                                                                                                  Range("BusinessZipCode").Value
            Case "Client"
                doc.getElementById("P1_DISCOVERED_BY_ROLE_CD").Value = "CLIENT"
            Case "Employee"
                doc.getElementById("P1_DISCOVERED_BY_ROLE_CD").Value = "EMP"
            Case "Participant"
                doc.getElementById("P1_DISCOVERED_BY_ROLE_CD").Value = "PARTCPNT"
            Case "Patient"
                doc.getElementById("P1_DISCOVERED_BY_ROLE_CD").Value = "PATIENT"
            Case "Payer"
                doc.getElementById("P1_DISCOVERED_BY_ROLE_CD").Value = "PAYER"
            Case "Private Party"
                doc.getElementById("P1_DISCOVERED_BY_ROLE_CD").Value = "PRIVATE"
            Case "Provider"
                doc.getElementById("P1_DISCOVERED_BY_ROLE_CD").Value = "PROVIDER"
            Case "Other"
                doc.getElementById("P1_DISCOVERED_BY_ROLE_CD").Value = "OTHER"
    
        End Select
    
    End If
    
    'DoesTheCallerHaveAnAccountNumber
    
    If Range("DoesTheCallerHaveAnAccountNumber") <> "" Then
    
        If Range("DoesTheCallerHaveAnAccountNumber") <> "No" Then
        
            doc.getElementById("P1_DISCOVERED_BY_ROLE_CD").Value = "OTHER"
        
        End If
    
    End If
    
    'DoYouKnowThisPatient - Not on PHITT Tool
    
    'DoYouKnowThisPatientOther ???
    
    'FaxNumber - Not needed on PHITT Tool
    
    'HaveTheResultsBeenOpenedOrViewed - PHI sent to/found by Unintented Recipient  on PHITT Tool
    
    If Range("HaveTheResultsBeenOpenedOrViewed") <> "" And Range("HowDidYouReceiveTheseResults") <> "Hacked" And Range("HowDidYouReceiveTheseResults") <> "Stolen" Then
    
        If Range("HaveTheResultsBeenOpenedOrViewed").Value = "Yes" Then
        
            Call ActivateWebFormField(doc, "P1_CAT_IR_STATUS_CD")
    
            doc.getElementById("P1_CAT_INCORRECT_RECIPIENT_I").Value = "Y"       'PHI sent to/found by Unintented Recipient is "Yes"
            doc.getElementById("P1_CAT_IR_STATUS_CD").Value = "OPEN" 'StatusOfPHI - This is "Status Of PHI" on the PHITT Tool
            doc.getElementById("P1_CAT_LOST_STOLEN_HARDWARE_I").Value = "N" 'PHI On Lost/Stolen Hardware Or Mobile Media on PHITT Tool
            doc.getElementById("P1_CAT_UNAUTHORIZED_ACCESS_I").Value = "N" 'PHI Accessed By An Unauthorized Entity, Contractor Or Hacker on PHITT Tool
            
            'HowDoYouKnowThisPatient

            If Range("HowDoYouKnowThisPatient") <> "" And Range("HowDoYouKnowThisPatient") <> "N/A" Then
            
                Select Case Range("HowDoYouKnowThisPatient").Value
                
                    Case "Coworker"
                        doc.getElementById("P1_CAT_KNOWN_RELATIONSHIP").Value = "Coworker"
                    Case "Family Member"
                        doc.getElementById("P1_CAT_KNOWN_RELATIONSHIP").Value = "Family Member"
                    Case "Insurance Guarantor"
                        doc.getElementById("P1_CAT_KNOWN_RELATIONSHIP").Value = "Insurance Guarantor"
                    Case "Neighbor"
                        doc.getElementById("P1_CAT_KNOWN_RELATIONSHIP").Value = "Neighbor"
                    Case "Other"
                        doc.getElementById("P1_CAT_KNOWN_RELATIONSHIP").Value = "OTHER"
            
                End Select
                
                'DueTo
    
                If Range("DueTo") <> "" And Range("DueTo") <> "N/A" Then
                
                Call ActivateWebFormField(doc, "P1_CAT_IR_REASON_CD")
                    
                    Select Case Range("DueTo").Value
                    
                        Case "Client Error"
                            doc.getElementById("P1_CAT_IR_REASON_CD").Value = "CLIENT"
                        Case "Employee Error"
                            doc.getElementById("P1_CAT_IR_REASON_CD").Value = "EMP"
                            Call ActivateWebFormField(doc, "P1_CAT_DEPT_FUNCTION_OTHER")
                        Case "Unknown"
                            doc.getElementById("P1_CAT_IR_REASON_CD").Value = "UNK"
                        Case "Vendor Error"
                            doc.getElementById("P1_CAT_IR_REASON_CD").Value = "VENDOR"
                        Case "Other"
                            doc.getElementById("P1_CAT_IR_REASON_CD").Value = "OTHER"
                            Call ActivateWebFormField(doc, "P1_CAT_IR_REASON_OTHER")
                
                    End Select
                
                End If
            
            End If
            
        Else: 'PHI sent to/found by Unintented Recipient is either "No" or "N/A" on the PHI Form
            
            doc.getElementById("P1_CAT_INCORRECT_RECIPIENT_I").Value = "N"
            
        End If
    
    End If
    
    'HowDidYouReceiveTheseResults
    
    If Range("HowDidYouReceiveTheseResults") <> "" Then
    
        Call ActivateWebFormField(doc, "P1_CAT_IR_DEL_MODE_CD")
    
        If Range("HowDidYouReceiveTheseResults") = "Electronic" Then
        
            doc.getElementById("P1_CAT_IR_DEL_MODE_CD").Value = "ELECT"
            
        ElseIf Range("HowDidYouReceiveTheseResults") = "Email" Then
        
            doc.getElementById("P1_CAT_IR_DEL_MODE_CD").Value = "EMAIL"
            
        ElseIf Range("HowDidYouReceiveTheseResults") = "Found in Public" Then
            
            doc.getElementById("P1_CAT_IR_STATUS_CD").Value = "FP"
            doc.getElementById("P1_CAT_IR_DEL_MODE_CD").Value = "OTHER"
            doc.getElementById("P1_CAT_IR_DEL_MODE_OTHER").Value = Range("AskTheCallerToExplain")
            
        ElseIf Range("HowDidYouReceiveTheseResults") = "Hacked" Then
            
            doc.getElementById("P1_CAT_INCORRECT_RECIPIENT_I").Value = "N"
            doc.getElementById("P1_CAT_LOST_STOLEN_HARDWARE_I").Value = "Y"
            
            If Range("DidCallerStateItWasHackerOrContractor") = "Yes" Then
            
                doc.getElementById("P1_CAT_UNAUTHORIZED_ACCESS_I").Value = "Y"
                
            Else:
            
                doc.getElementById("P1_CAT_UNAUTHORIZED_ACCESS_I").Value = "N"
            
            End If
            
        ElseIf Range("HowDidYouReceiveTheseResults") = "Mail" Then
            
            doc.getElementById("P1_CAT_IR_DEL_MODE_CD").Value = "MAIL"
            
        ElseIf Range("HowDidYouReceiveTheseResults") = "MyQuest" Then
            
            doc.getElementById("P1_CAT_IR_DEL_MODE_CD").Value = "MYQUEST"
            
        ElseIf Range("HowDidYouReceiveTheseResults") = "Printer" Then
            
            doc.getElementById("P1_CAT_IR_DEL_MODE_CD").Value = "PRINTER"
            
        ElseIf Range("HowDidYouReceiveTheseResults") = "Social Media" Then
            
            doc.getElementById("P1_CAT_IR_STATUS_CD").Value = "SOCIAL MEDIA"
            doc.getElementById("P1_CAT_IR_DEL_MODE_CD").Value = "OTHER"
            doc.getElementById("P1_CAT_IR_DEL_MODE_OTHER").Value = Range("AskTheCallerToExplain")
            
        ElseIf Range("HowDidYouReceiveTheseResults") = "Stolen" Then
            
            doc.getElementById("P1_CAT_INCORRECT_RECIPIENT_I").Value = "N"
            doc.getElementById("P1_CAT_LOST_STOLEN_HARDWARE_I").Value = "Y"
            
        ElseIf Range("HowDidYouReceiveTheseResults") = "Verbal" Then
            
            doc.getElementById("P1_CAT_IR_STATUS_CD").Value = "ORAL"
            
            doc.getElementById("P1_CAT_IR_DEL_MODE_CD").Value = "ORAL"
            
        ElseIf Range("HowDidYouReceiveTheseResults") = "Other" Then
            
            doc.getElementById("P1_CAT_IR_DEL_MODE_CD").Value = "OTHER"
            doc.getElementById("P1_CAT_IR_DEL_MODE_OTHER").Value = Range("AskTheCallerToExplain")
        
        End If
        
        If doc.getElementById("P1_CAT_IR_DEL_MODE_CD").Value = "OTHER" Then
        
            Call ActivateWebFormField(doc, "P1_CAT_IR_DEL_MODE_OTHER")
        
        End If
    
    End If
    
    'AccountNumber
    
    If Range("AccountNumber") <> "" Then
    
        If Range("AccountNumber") = "No" Then
        
            doc.getElementById("P1_CLIENT_ACCOUNT").Value = "N/A"
            
        Else:
        
            doc.getElementById("P1_CLIENT_ACCOUNT").Value = Range("AccountNumber").Value
        
        End If
    
    End If
    
    'IncorrectEmailAddress - Not needed on PHITT Tool

    'PatientName - "Breached Patient Name(s)" on PHITT Tool
    
    doc.getElementById("P1_PATIENT_NAMES").Value = Range("PatientName").Value
    
    'PatientsStateOfResidence
    
    doc.getElementById("P1_RESIDENCE").Value = Range("PatientsStateOfResidence").Value
    
    'TestNames
    
    If Range("TestNames") <> "" Then
    
        doc.getElementById("P1_INFORMATION_DISCLOSED_0").Checked = True
        doc.getElementById("P1_TEST_INVOLVED").Value = Range("TestNames").Value
    
    End If
    
    'IsQuestAtRisk
    
    If Range("IsQuestAtRisk") = "Yes" Then
    
        doc.getElementById("P1_PHI_BEHAVIOR_RISK_I").Value = "Y"
        
    Else:
    
        doc.getElementById("P1_PHI_BEHAVIOR_RISK_I").Value = "N"
    
    End If
    
    'OtherPertinentInformation
    
    doc.getElementById("P1_EVENT_DESCRIPTION").Value = Range("OtherPertinentInformation")
    
    'SelectYourBusinessUnit
    
    If Range("SelectYourBusinessUnit") <> "" Then
    
        Select Case Range("SelectYourBusinessUnit").Value
        
            Case "Albuquerque"
                doc.getElementById("P1_LOCATION").Value = "Albuquerque"
            Case "Atlanta"
                doc.getElementById("P1_LOCATION").Value = "Atlanta"
            Case "Auburn Hills"
                doc.getElementById("P1_LOCATION").Value = "Auburn Hills"
            Case "Baltimore"
                doc.getElementById("P1_LOCATION").Value = "Baltimore"
            Case "Cincinnati"
                doc.getElementById("P1_LOCATION").Value = "Cincinnati"
            Case "Dallas"
                doc.getElementById("P1_LOCATION").Value = "Dallas"
            Case "Denver"
                doc.getElementById("P1_LOCATION").Value = "Denver"
            Case "DLO"
                doc.getElementById("P1_LOCATION").Value = "DLO"
            Case "Houston"
                doc.getElementById("P1_LOCATION").Value = "Houston"
            Case "Las Vegas"
                doc.getElementById("P1_LOCATION").Value = "Las Vegas"
            Case "Lenexa"
                doc.getElementById("P1_LOCATION").Value = "St Louis"
            Case "MACL"
                doc.getElementById("P1_LOCATION").Value = "MACL"
            Case "Marlborough"
                doc.getElementById("P1_LOCATION").Value = "Marlborough"
            Case "Miami"
                doc.getElementById("P1_LOCATION").Value = "Miami"
            Case "New Orleans"
                doc.getElementById("P1_LOCATION").Value = "New Orleans"
            Case "Philadelphia"
                doc.getElementById("P1_LOCATION").Value = "Philadelphia"
            Case "Pittsburgh"
                doc.getElementById("P1_LOCATION").Value = "Pittsburgh"
            Case "Puerto Rico"
                doc.getElementById("P1_LOCATION").Value = "Puerto Rico"
            Case "Sacramento"
                doc.getElementById("P1_LOCATION").Value = "Sacramento"
            Case "Seattle"
                doc.getElementById("P1_LOCATION").Value = "Seattle"
            Case "Solstas"
                doc.getElementById("P1_LOCATION").Value = "Solstas"
            Case "Syosset"
                doc.getElementById("P1_LOCATION").Value = "NewJersey/New York"
            Case "Tampa"
                doc.getElementById("P1_LOCATION").Value = "Tampa"
            Case "Teterboro"
                doc.getElementById("P1_LOCATION").Value = "NewJersey/New York"
            Case "Wallingford"
                doc.getElementById("P1_LOCATION").Value = "Wallingford"
            Case "West Hills"
                doc.getElementById("P1_LOCATION").Value = "West Hills"
            Case "Wood Dale"
                doc.getElementById("P1_LOCATION").Value = "Chicago"
        
        End Select
        
    End If
    
    'StatePhiWasMailedTo - Not needed on PHITT Tool
    
    'TypeOfDataDiscovered
    
    If Range("TypeOfDataDiscovered") = "Bills" Then
    
        doc.getElementById("P1_DISCOVERED_DATA_TYPE_CD_3").Checked = True
    
    ElseIf Range("TypeOfDataDiscovered") = "Electronic PHI (USB, etc.)" Then
    
        doc.getElementById("P1_DISCOVERED_DATA_TYPE_CD_7").Checked = True
    
    ElseIf Range("TypeOfDataDiscovered") = "Eligibility File" Then
    
        doc.getElementById("P1_DISCOVERED_DATA_TYPE_CD_1").Checked = True
    
    ElseIf Range("TypeOfDataDiscovered") = "Results" Then
    
        doc.getElementById("P1_DISCOVERED_DATA_TYPE_CD_4").Checked = True
        Call ActivateWebFormField(doc, "P1_IF_RESULTS")
        
        If Range("WereTheResultsNormalOrAbnormal") <> "" And Range("WereTheResultsNormalOrAbnormal") <> "N/A" Then
            doc.getElementById("P1_IF_RESULTS").Value = Range("WereTheResultsNormalOrAbnormal").Value
        End If
        
    ElseIf Range("TypeOfDataDiscovered") = "Requisition(s)" Then
    
        doc.getElementById("P1_DISCOVERED_DATA_TYPE_CD_5").Checked = True
    
    ElseIf Range("TypeOfDataDiscovered") = "Specimens" Then
    
        doc.getElementById("P1_DISCOVERED_DATA_TYPE_CD_6").Checked = True
    
    ElseIf Range("TypeOfDataDiscovered") = "Vaccine Form" Then
    
        doc.getElementById("P1_DISCOVERED_DATA_TYPE_CD_0").Checked = True
    
    ElseIf Range("TypeOfDataDiscovered") = "Wellness File" Then
    
        doc.getElementById("P1_DISCOVERED_DATA_TYPE_CD_2").Checked = True
    
    ElseIf Range("TypeOfDataDiscovered") = "Other" Then
    
        doc.getElementById("P1_DISCOVERED_DATA_TYPE_CD_8").Checked = True
    
    End If
    
    'TodaysDate
    
    If Range("TodaysDate") <> "" Then
    
        doc.getElementById("P1_DISCOVERED_DATE").Value = Format(Range("TodaysDate"), "mm/dd/yyyy")
    
    End If
    
    'OrderingPhysiciansName
    
    If Range("OrderingPhysiciansName") <> "" Then
    
        doc.getElementById("P1_ORDERING_PHYSICIAN").Value = Range("OrderingPhysiciansName")
    
    End If
    
    'Were500OrMorePatientsEffected
    
    If Range("Were500OrMorePatientsEffected") <> "" Then
    
        If Range("Were500OrMorePatientsEffected") = "Yes" Then
        
            doc.getElementById("P1_MORE_THAN_500_PATIENTS_I").Value = "Y"
        
        Else:
        
            doc.getElementById("P1_MORE_THAN_500_PATIENTS_I").Value = "N"
            
        End If
    
    End If
    
    'WereTheResultsSharedWithAnotherParty
    
    If Range("WereTheResultsSharedWithAnotherParty") = "No" Then
    
        doc.getElementById("P1_PHI_PROTECTED_NO_DISCLOSE_I").Value = "Y"
        
    Else:
    
        doc.getElementById("P1_PHI_PROTECTED_NO_DISCLOSE_I").Value = "N"
        Call ActivateWebFormField(doc, "P1_PHI_NO_DISCLOSE_COMMENT")
        doc.getElementById("P1_PHI_NO_DISCLOSE_COMMENT").Value = Range("AskTheCallerForFirstAndLastName")
    
    End If
    
    'WhatWasTheDateYouReceivedThisReport - "Date Incident Occurred" on PHITT Tool
    
    doc.getElementById("P1_OCCURRED_DATE").Value = Format(Range("WhatWasTheDateYouReceivedThisReport"), "mm/dd/yyyy")
    
    'Completed By Name on PHITT Tool
    
    doc.getElementById("P1_COMPLETED_BY").Value = WorksheetFunction.Proper(GetUserName())
    
    'Completed By Contact Info on PHITT Tool
    
    doc.getElementById("P1_COMPLETED_BY_CONTACT").Value = WorksheetFunction.Proper(GetUserName()) & "@QuestDiagnostics.com"
    
    'EmployeeErrorType
    
    If Range("EmployeeErrorType") <> "" Then
        
        Select Case Range("EmployeeErrorType").Value
        
            Case "Account Number On Document Submitted Was Not Used"
                doc.getElementById("P1_EMP_ERROR_TYPE").Value = "A3"
                
            Case "An Extra Report/Bill Was Stuffed In An Envelop"
                doc.getElementById("P1_EMP_ERROR_TYPE").Value = "M1"
            
            Case "Email Not Sent Secured"
                doc.getElementById("P1_EMP_ERROR_TYPE").Value = "E1"
                
            Case "Entered A Fax Number In Error"
                doc.getElementById("P1_EMP_ERROR_TYPE").Value = "M1"
                
            Case "Entered A Wrong Fax Number"
                doc.getElementById("P1_EMP_ERROR_TYPE").Value = "F5"
                
            Case "Incorrect Copy To Account Was Entered Or Selected"
                doc.getElementById("P1_EMP_ERROR_TYPE").Value = "A4"
                
            Case "Other"
                doc.getElementById("P1_EMP_ERROR_TYPE").Value = "A5"
                
            Case "Other Fax Error"
                doc.getElementById("P1_EMP_ERROR_TYPE").Value = "F6"
                
            Case "Other Mail Error"
                doc.getElementById("P1_EMP_ERROR_TYPE").Value = "M6"
                
            Case "Patient Demographics Used Were Incorrect"
                doc.getElementById("P1_EMP_ERROR_TYPE").Value = "A7"
                
            Case "Report Was Delivered To The Wrong Client"
                doc.getElementById("P1_EMP_ERROR_TYPE").Value = "M4"
                
            Case "Selected The Wrong Patient"
                doc.getElementById("P1_EMP_ERROR_TYPE").Value = "M2"
                
            Case "Specimen Mishandled, Lost, Dropped"
                doc.getElementById("P1_EMP_ERROR_TYPE").Value = "M5"
                
            Case "Typographical Error (E.G. Client Number Off By 1, 2 Or 3 Digits)"
                doc.getElementById("P1_EMP_ERROR_TYPE").Value = "F1"
                
            Case "Wrong Account Number"
                doc.getElementById("P1_EMP_ERROR_TYPE").Value = "A1"
                
            Case "Wrong Account Was Selected"
                doc.getElementById("P1_EMP_ERROR_TYPE").Value = "F3"
                
            Case "Wrong Copy To Was Used"
                doc.getElementById("P1_EMP_ERROR_TYPE").Value = "F2"
                
            Case "Wrong Email Address Used"
                doc.getElementById("P1_EMP_ERROR_TYPE").Value = "E3"
            
        End Select
        
    End If
    
    'EmployeeErrorTypeOther
    
    If Range("EmployeeErrorType") <> "N/A" Then
    
        doc.getElementById("P1_EMP_ERROR_OTHER").Value = Range("EmployeeErrorTypeOther")
    
    End If
    
End Sub
