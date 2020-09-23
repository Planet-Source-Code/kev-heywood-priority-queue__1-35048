VERSION 5.00
Begin VB.Form frmPQueue 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patient Priority Queue"
   ClientHeight    =   6735
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8190
   Icon            =   "frmPQueue.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraPatientsInput 
      Height          =   1695
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8175
      Begin VB.TextBox txtSurname 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         MaxLength       =   14
         TabIndex        =   7
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox txtPriority 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5160
         MaxLength       =   1
         TabIndex        =   6
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtHrs 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6720
         MaxLength       =   2
         TabIndex        =   5
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtMins 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7320
         MaxLength       =   2
         TabIndex        =   4
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add to Patient List"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save to File"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6720
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblSname 
         Caption         =   "Patients Surname"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblPs 
         Caption         =   "Priority Status"
         Height          =   375
         Left            =   4440
         TabIndex        =   12
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblToa 
         Caption         =   "Time of Arrival"
         Height          =   495
         Left            =   6000
         TabIndex        =   11
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblHrs 
         Caption         =   "Hrs"
         Height          =   255
         Left            =   6840
         TabIndex        =   10
         Top             =   760
         Width           =   375
      End
      Begin VB.Label lblMins 
         Caption         =   "Mins"
         Height          =   255
         Left            =   7440
         TabIndex        =   9
         Top             =   760
         Width           =   375
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PATIENT INPUT DATA"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   2400
         TabIndex        =   8
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   8180
      Begin VB.ListBox lstPQ 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         ItemData        =   "frmPQueue.frx":030A
         Left            =   75
         List            =   "frmPQueue.frx":030C
         TabIndex        =   15
         Top             =   2880
         Width           =   8010
      End
      Begin VB.ListBox lstPList 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         ItemData        =   "frmPQueue.frx":030E
         Left            =   3320
         List            =   "frmPQueue.frx":0310
         TabIndex        =   14
         Top             =   360
         Width           =   4770
      End
      Begin VB.Label lblDelete 
         Caption         =   "Double Click on an Item to Delete"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5640
         TabIndex        =   24
         Top             =   120
         Width           =   2445
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   240
         X2              =   7800
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Treatment Ends"
         Height          =   495
         Left            =   6720
         TabIndex        =   23
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Treatment Starts"
         Height          =   495
         Left            =   5880
         TabIndex        =   22
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Doctor"
         Height          =   255
         Left            =   4440
         TabIndex        =   21
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Time Arrived"
         Height          =   255
         Left            =   2520
         TabIndex        =   20
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Priority"
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Patient Name"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2640
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   1455
         Left            =   240
         Picture         =   "frmPQueue.frx":0312
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label lblQueue 
         Caption         =   "Patient Priority Queue"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label lblPList 
         Caption         =   "Patient List"
         Height          =   255
         Left            =   3360
         TabIndex        =   16
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About "
      Begin VB.Menu mnuAboutApp 
         Caption         =   "   Priority Queue Help Application"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAboutKH 
         Caption         =   "     Written by - Kev Heywood"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAboutCopyRight 
         Caption         =   "      Copyright (c) May 2002"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAboutEmail 
         Caption         =   "  kevan@000h.freeserve.co.uk"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmPQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************
'                           HELP WITH PRIORITY QUEUES
'Imagine a number of patients waiting in an emergency room of a local hospital.
'Priority is related to the seriousness of the patient's condition.
'
'                 (1 = most serious, 4 = least serious).
'
'                    We assume that a doctor spends
'
'           .  60 minutes with a patient whose priority is 1
'
'           .  45 minutes with a patient whose priority is 2
'
'           .  30 minutes with a patient whose priority is 3
'
'           .  15 minutes with a patient whose priority is 4

'The patient's treatment is not interrupted, even if a more serious case arrives.
'Two Doctors are on call.  Dr. Alpha starts seeing patients at 6:00 and Dr. Beta
'starts at 6:30. The programs goal is to list the patients in the order in which
'they will be treated.
'
'
'
'           All the best - Kev Heywood - Email kevan@000h.freeserve.co.uk
'***********************************************************************************
Option Explicit

Private Type Patients       'Define Patient Type
    SN As String            'Surname
    PS As Integer           'Priority Status
    TA As Double            'Time Arrival
End Type

Dim ERPats() As Patients    'Dynamic Patient Array
Dim SortTA As Variant       'Dynamic Index Sort Array on Time Arrival
Dim pAdd As Long            'Main index counter
Dim TimeInMins As Single    'Conversion of hours+minutes to minutes
Dim FileNum As Integer      'Next free file number
Dim appPath As String       'Application Path
Dim DataName As String      'Data FileName
Dim ErrMsg(4) As String     'Array of Error messages
Dim ListChanged As Boolean  'Assign True when the Patients Array changes

Private Sub Form_Load()
    'Centre the form
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    appPath = App.Path + "\"    'Patient File Location
    DataName = "ERList.Dat"     'Patient FileName
    'Error messages
    ErrMsg(1) = "Name": ErrMsg(2) = "Priority": ErrMsg(3) = "Hour": ErrMsg(4) = "Minute"
    'Deal with 'No data file Present' Scenario
    On Error Resume Next
    If FileLen(appPath + DataName) Then Call LoadFile
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Dim response As Integer 'Return value indicating which button the user clicked.
    Dim uloop As Integer    'Simple Loop variable
    
    If ListChanged Then 'Check to see if contents have changed
        response = MsgBox("Do you want to save the changes you made?" + vbCrLf + vbCrLf _
        , vbYesNoCancel + vbInformation, "SAVE CHANGES")
        If response = vbYes Then
            Call cmdSave_Click  'Save Changes
        ElseIf response = vbCancel Then
            Cancel = 1: Exit Sub    'Stop the form and application from closing.
        End If
    End If
    
    'Unload the form and close the application
    Unload Me
    End
End Sub

Private Sub mnuFileExit_Click()
    Unload Me   'Close the application
End Sub

Private Sub txtSurname_KeyPress(KeyAscii As Integer)
    'On Return Key move focus to next TextBox
    If KeyAscii = 13 Then txtPriority.SetFocus
End Sub

Private Sub txtSurname_LostFocus()
    'On Focus lost ensure first character of name is a Capital
    txtSurname = UCase(Mid(txtSurname, 1, 1)) + Mid(txtSurname, 2, Len(txtSurname))
End Sub

Private Sub txtPriority_KeyPress(KeyAscii As Integer)
    'On Return Key move focus to next TextBox
    If KeyAscii = 13 Then txtHrs.SetFocus
End Sub

Private Sub txtHrs_KeyPress(KeyAscii As Integer)
    'On Return Key move focus to next TextBox
    If KeyAscii = 13 Then txtMins.SetFocus
End Sub

Private Sub txtMins_KeyPress(KeyAscii As Integer)
    'On Return Key Add to Lists
    If KeyAscii = 13 Then Call cmdAdd_Click
End Sub

Private Sub cmdAdd_Click()
    Dim anyErrors As Integer    'Return value from Errors function
    anyErrors = Errors          'Ascertain an error value
    If anyErrors > 0 Then
        'Notify user of error on exit the sub
        MsgBox ErrMsg(anyErrors) + " Field Incorrect", vbCritical, "User Error"
        Exit Sub
    End If
   
    TimeInMins = (Val(txtHrs) * 60) + Val(txtMins)  'Calculate Arrival Time in minutes
    
    pAdd = pAdd + 1                 'Increment main index counter
    ReDim Preserve ERPats(pAdd)     'Increase array size and preserve the exisiting index
    With ERPats(pAdd)               'Assign the text fields to the properties of the Type
        .SN = txtSurname            'Surname
        .PS = Val(txtPriority)      'Priority Status
        .TA = TimeInMins            'Arrival Time (In Minutes)
    End With
    cmdSave.Enabled = True          'Allow user to Save changes to the file
    txtSurname = "": txtPriority = "": txtHrs = "": txtMins = ""    'Set textboxes to Null
    txtSurname.SetFocus             'Set focus to Surname textbox ready for next addition
    Call PopPatientList             'Update Patient List
    Call PopPriorityList            'Update Priority List
    ListChanged = True              'Set to record that a change has taken place
End Sub

Private Sub cmdSave_Click()
    Dim loopER As Long      'Main Index loop counter
    FileNum = FreeFile      'Get next File number
    'Save the contents of the Patient Array
    Open appPath + DataName For Output As FileNum
        For loopER = 1 To pAdd
            Write #FileNum, ERPats(loopER).SN, ERPats(loopER).PS, ERPats(loopER).TA
        Next loopER
    Close FileNum
    cmdSave.Enabled = False     'All changes saved to file so disable control
    ListChanged = False         'Set to record that no change has taken place
End Sub

Private Function Errors() As Integer
    'User Input Error Capture Function
    If txtSurname = "" Then Errors = 1: txtSurname.SetFocus: Exit Function
    If Val(txtPriority) < 1 Or Val(txtPriority) > 4 Then
        Errors = 2: txtPriority = "": txtPriority.SetFocus: Exit Function
    End If
    Dim IsNumber As Boolean 'Check if number
    IsNumber = IsNumeric(txtHrs)
    'If it's not a number or number is out of range then raise Error
    If Not IsNumber Or Val(txtHrs) < 0 Or Val(txtHrs) > 23 Then
        Errors = 3: txtHrs = "": txtHrs.SetFocus: Exit Function
    End If
    IsNumber = IsNumeric(txtMins) 'Check if number
    'If it's not a number or number is out of range then raise Error
    If Not IsNumber Or Val(txtMins) < 0 Or Val(txtMins) > 59 Then
        Errors = 4: txtMins = "": txtMins.SetFocus: Exit Function
    End If
End Function

Private Sub LoadFile()
    FileNum = FreeFile      'Get next File number
    'Load the contents of the data file into the Patient Array
    Open appPath + DataName For Input As FileNum
        While Not EOF(FileNum)
            pAdd = pAdd + 1     'Increment main index array counter
            ReDim Preserve ERPats(pAdd) 'Increase array and preserve the exisiting index
            Input #FileNum, ERPats(pAdd).SN, ERPats(pAdd).PS, ERPats(pAdd).TA
        Wend
    Close FileNum
    Call PopPatientList     'Update the Patient List
    Call PopPriorityList    'Update the Priority List
End Sub

Private Sub PopPatientList()
    Dim loopER As Long          'Main Index loop counter
    Dim joinData As String      'Concatenation of list Data
    Dim TabSN As String         'Surname Padding spaces
    lstPList.Clear              'Clear the Patient List
    For loopER = 1 To UBound(ERPats)
        TabSN = String(16 - Len(ERPats(loopER).SN), Chr$(32)) 'Surname + padding
        joinData = ERPats(loopER).SN + TabSN + Format(ERPats(loopER).PS) + _
        vbTab + Format(Int(ERPats(loopER).TA / 60), "0#") + ":" _
        + Format(ERPats(loopER).TA Mod 60, "0#")
        lstPList.AddItem joinData
    Next loopER
    'Highlight the last item in the list
    If lstPList.ListCount - 1 > 0 Then lstPList.Selected(lstPList.ListCount - 1) = True
End Sub

Private Sub PopPriorityList()
    SortTA = ArriveTimeSort()           'Acquire the Sort Time Arrival Index
    Dim loopER As Long                  'Main Index loop counter
    Dim joinData As String              'Concatenation of list Data
    Dim DrA As String                   'First Doctor
    Dim DrB As String                   'Second Doctor
    Dim dr1 As Single                   'First Doctor's Patient/Time duration
    Dim dr2 As Single                   'Second Doctor's Patient/Time duration
    Dim timeP(4) As Integer             'Priority Time Array
    Dim TabSN As String                 'Surname Padding spaces
    lstPQ.Clear                         'Clear the Priority List
    DrA = "Dr Alpha": DrB = "Dr Beta"   'Assign Doctor's names
    timeP(1) = 60: timeP(2) = 45: timeP(3) = 30: timeP(4) = 15 'Assign Priority Times
    dr1 = 360: dr2 = 390                'Doctor's Start Times
    For loopER = 1 To UBound(ERPats)    'Loop through the main index array
        TabSN = String(16 - Len(ERPats(SortTA(loopER)).SN), Chr$(32)) 'Surname + padding
        If dr1 <= dr2 Then 'Doctor Alpha becomes free before or at the same time as Dr Beta
            If ERPats(SortTA(loopER)).TA > dr1 Then dr1 = ERPats(SortTA(loopER)).TA
            joinData = ERPats(SortTA(loopER)).SN + TabSN + Format(ERPats(SortTA(loopER)).PS) + _
            vbTab + Format(Int(ERPats(SortTA(loopER)).TA / 60), "0#") + ":" _
            + Format(ERPats(SortTA(loopER)).TA Mod 60, "0#") + _
            vbTab + vbTab + DrA + vbTab + Format(Int(dr1 / 60), "0#") + ":" + _
            Format(dr1 Mod 60, "0#") + vbTab
            dr1 = dr1 + timeP(ERPats(SortTA(loopER)).PS)    'Record Time spent with Patient
            joinData = joinData + Format(Int(dr1 / 60), "0#") + ":" + Format(dr1 Mod 60, "0#")
            lstPQ.AddItem joinData  'Add to Priority List
        Else 'Doctor Beta becomes free before Dr Alpha
            If ERPats(SortTA(loopER)).TA > dr2 Then dr2 = ERPats(SortTA(loopER)).TA
            joinData = ERPats(SortTA(loopER)).SN + TabSN + Format(ERPats(SortTA(loopER)).PS) + _
            vbTab + Format(Int(ERPats(SortTA(loopER)).TA / 60), "0#") + ":" _
            + Format(ERPats(SortTA(loopER)).TA Mod 60, "0#") + _
            vbTab + vbTab + DrB + vbTab + vbTab + Format(Int(dr2 / 60), "0#") + ":" + _
            Format(dr2 Mod 60, "0#") + vbTab
            dr2 = dr2 + timeP(ERPats(SortTA(loopER)).PS)    'Record Time spent with Patient
            joinData = joinData + Format(Int(dr2 / 60), "0#") + ":" + Format(dr2 Mod 60, "0#")
            lstPQ.AddItem joinData  'Add to Priority List
        End If
    Next loopER
End Sub

Private Function ArriveTimeSort() As Variant
    'Returns an index of pointers To the array
    Dim OuterLoop As Long       'Outside Loop
    Dim InnerLoop As Long       'Inside Loop
    Dim tempIndex() As Long     'Temporary dynamic array
    Dim LB As Long              'Lower bounds of Patient Array
    Dim UB As Long              'Upper bounds of Patient Array
    
    LB = LBound(ERPats) + 1     'Assign Lower Bound Base 1
    UB = UBound(ERPats)         'Assign Upper Bound
    
    
    ReDim tempIndex(UB)         'Increase array size
    For InnerLoop = LB To UB
        tempIndex(InnerLoop) = InnerLoop    'Assign values to temporary array index
    Next

    For OuterLoop = UB To LB Step -1            'Step backwards through array
        For InnerLoop = LB + 1 To OuterLoop     'Step forwards through array
            'Compare Time Arrivals and force higher times to the OuterLoop array index
            If ERPats(tempIndex(InnerLoop - 1)).TA > ERPats(tempIndex(OuterLoop)).TA Then
                Swap tempIndex(InnerLoop - 1), tempIndex(OuterLoop)
            End If
        Next InnerLoop
    Next OuterLoop

    ArriveTimeSort = tempIndex()    'Return Sorted index of pointers
    
End Function

Private Sub Swap(ByRef firstValue As Long, ByRef secondValue As Long)
    'Swaps the First value with the Second Value using an intermediate variable
    Dim tmpValue As Variant
    tmpValue = firstValue
    firstValue = secondValue
    secondValue = tmpValue
End Sub

Private Sub lstPList_DblClick()
    Dim response As Integer     'Return value indicating which button the user clicked.
    Dim tempIndex As Integer    'Record the index item of the Patient ListBox
    'Ask User for confirmation of a delete operation
    response = MsgBox("Delete Patient?", vbInformation + vbYesNo, "Delete")
    If response = 6 Then
        tempIndex = lstPList.ListIndex + 1 'Add one because listbox index array is base 0
        'Reposition the array contents from the postion of the deleted item
        For response = tempIndex To UBound(ERPats) - 1
            ERPats(response) = ERPats(response + 1)
        Next response
        'Tidy up the array and release the data held in the deleted index position
        ERPats(response).SN = "": ERPats(response).PS = Empty: ERPats(response).TA = Empty
        ReDim Preserve ERPats(response - 1) 'Decrease the Patient Array
        pAdd = pAdd - 1                     'Decrease the Patient counter
        Call PopPatientList                 'Re-populate the Patient list
        'Highlight the next indexed item in the Patient list
        If UBound(ERPats) > tempIndex - 1 Then lstPList.Selected(tempIndex - 1) = True
        Call PopPriorityList                'Re-populate the Priority list
        cmdSave.Enabled = True              'Allow user to Save changes to the file
        ListChanged = True                  'Set to record that a change has taken place
    End If
End Sub

