VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6945
   ClientLeft      =   1635
   ClientTop       =   2115
   ClientWidth     =   8835
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   8835
   Begin VB.CommandButton Command4 
      Caption         =   "Sume"
      Height          =   255
      Left            =   4800
      TabIndex        =   8
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Get Sub"
      Height          =   375
      Left            =   6600
      TabIndex        =   7
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get Functions"
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   2295
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   4680
      Width           =   8775
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   3855
      Left            =   2880
      TabIndex        =   4
      Top             =   840
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6800
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get All"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1695
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
   Begin VB.FileListBox File1 
      Height          =   2625
      Left            =   0
      Pattern         =   "*.frm"
      TabIndex        =   0
      Top             =   2040
      Width           =   2775
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Put a Err.H in Sub"
      Height          =   375
      Left            =   7320
      TabIndex        =   10
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   6600
      TabIndex        =   9
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AcusticoV() As String
Dim AcusticoF() As String
Function Getting()
ReDim AcusticoV(File1.ListCount)
ReDim AcusticoF(File1.ListCount)
lv1.ColumnHeaders.Clear
lv1.ColumnHeaders.Add , , "File", 2000
lv1.ColumnHeaders.Add , , "Function", 3000
lv1.ColumnHeaders.Add , , "getFI", 0
lv1.ColumnHeaders.Add , , "getFF", 0
lv1.ColumnHeaders.Add , , "getFC", 0
lv1.ColumnHeaders.Add , , "SIZE", 1000
lv1.ListItems.Clear
Dim xFob As String
    For k = 0 To File1.ListCount - 1
        Open File1.Path & "\" & File1.List(k) For Binary As #1
        xFob = String(LOF(1), "x")
        Get #1, , xFob
        AcusticoV(k) = xFob
        AcusticoF(k) = File1.List(k)
    Close #1
    Next k
End Function
Private Sub Command1_Click()
    Getting
    GetAcusticos
End Sub
Function GetAcusticos()
    Debug.Print UBound(AcusticoV)
    For k = 0 To UBound(AcusticoV)
        GetAllF AcusticoV(k), AcusticoF(k)
        GetAllC AcusticoV(k), AcusticoF(k)
    Next k
End Function
Function setSub()
Debug.Print UBound(AcusticoV)
Dim big As String
    For k = 0 To UBound(AcusticoV)
        acus = AcusticoV(k)
        fileN = AcusticoF(k)
        getFF = InStr(1, acus, "Private Sub")
        If getfi > 0 Then
         big = Mid(acus, 1, getFF - 1)
        End If
       For j = 1 To Len(acus)
       getfi = InStr(j, acus, "Private Sub")
       DoEvents
       If getfi > 0 Then
            getFF = InStr(getfi, acus, "End Sub")
            getFFI = InStr(getfi, acus, ")")
           ' lv1.ListItems.Add , , fileN
           ' lv1.ListItems.Item(lv1.ListItems.Count).SubItems(1) = Mid(acus, getfi, getFFI - getfi + 1)
           ' lv1.ListItems.Item(lv1.ListItems.Count).SubItems(2) = getfi
           ' lv1.ListItems.Item(lv1.ListItems.Count).SubItems(3) = getff
           ' lv1.ListItems.Item(lv1.ListItems.Count).SubItems(4) = Mid(acus, getfi, getff - getfi + 7)
           ' lv1.ListItems.Item(lv1.ListItems.Count).SubItems(5) = Len(Mid(acus, getfi, getff - getfi + 7))
            'List1.AddItem Mid(acus, getFI, getFFI - getFI + 1)
            Debug.Print "-->  " & Mid(acus, getfi, getFFI - getfi + 1)
            Debug.Print Mid(big, Len(big) - 25, 25)
            big = big & Mid(acus, getfi, getFFI - getfi + 1)
            big = big & vbCrLf & "On error goto ErrHandler"
            big = big & Mid(acus, getFFI + 1, getFF - getFFI - 1)
            Myob = Mid(acus, getfi + 12, getFFI - getfi - 13)
            Myob = Mid(Myob, 1, InStr(1, Myob, "_") - 1)
            big = big & vbCrLf & "Exit Sub"
            big = big & vbCrLf & "ErrHandler:" & vbCrLf & "frmDebug.printF err," & Myob & vbCrLf
            big = big & Mid(acus, getFF, 10)
            Debug.Print "-->  " & Mid(acus, getFF, 7)
            
            j = getFFI
        Else
            j = Len(acus)
       End If
    Next j
writer:
If Len(fileN) > 0 Then
    Open "c:\" & fileN For Output As #1
                Print #1, big
            Close #1
End If
   Next k

'writer:
'    Open "c:\test.frm" For Output As #1
'                Print #1, big
'            Close #1
End Function
Function GetAllC(acus As String, fileN As String)
        For k = 1 To Len(acus)
       getfi = InStr(k, acus, "Private Sub")
       Debug.Print Mid(acus, k, 50)
       DoEvents
       If getfi > 0 Then
            getFF = InStr(getfi, acus, "End Sub")
            getFFI = InStr(getfi, acus, ")")
            lv1.ListItems.Add , , fileN
            lv1.ListItems.Item(lv1.ListItems.Count).SubItems(1) = Mid(acus, getfi, getFFI - getfi + 1)
            lv1.ListItems.Item(lv1.ListItems.Count).SubItems(2) = getfi
            lv1.ListItems.Item(lv1.ListItems.Count).SubItems(3) = getFF
            lv1.ListItems.Item(lv1.ListItems.Count).SubItems(4) = Mid(acus, getfi, getFF - getfi + 7)
            lv1.ListItems.Item(lv1.ListItems.Count).SubItems(5) = Len(Mid(acus, getfi, getFF - getfi + 7))
            'List1.AddItem Mid(acus, getFI, getFFI - getFI + 1)
            Debug.Print "-->  " & Mid(acus, getfi, getFFI - getfi + 1)
            k = getFFI
        Else
            k = Len(acus)
       End If
    Next k
End Function
Function GetAllF(acus As String, fileN As String)
    For k = 1 To Len(acus)
       getfi = InStr(k, acus, "Function ")
       Debug.Print Mid(acus, k, 50)
       DoEvents
       If getfi > 0 Then
            getFF = InStr(getfi, acus, "End Function")
            If getFF <> 0 Then
            getFFI = InStr(getfi, acus, ")")
            lv1.ListItems.Add , , fileN
            lv1.ListItems.Item(lv1.ListItems.Count).SubItems(1) = Mid(acus, getfi, getFFI - getfi + 1)
            lv1.ListItems.Item(lv1.ListItems.Count).SubItems(2) = getfi
            lv1.ListItems.Item(lv1.ListItems.Count).SubItems(3) = getFF
            lv1.ListItems.Item(lv1.ListItems.Count).SubItems(4) = Mid(acus, getfi, getFF - getfi + 12)
            lv1.ListItems.Item(lv1.ListItems.Count).SubItems(5) = Len(Mid(acus, getfi, getFF - getfi + 7))
            'List1.AddItem Mid(acus, getFI, getFFI - getFI + 1)
            Debug.Print "-->  " & Mid(acus, getfi, getFFI - getfi + 1)
            k = getFFI
            End If
        Else
            k = Len(acus)
       End If
    Next k
End Function

Private Sub Command2_Click()
    Getting
    Debug.Print UBound(AcusticoV)
    For k = 0 To UBound(AcusticoV)
        GetAllF AcusticoV(k), AcusticoF(k)
    Next k
End Sub

Private Sub Command3_Click()
    Getting
    Debug.Print UBound(AcusticoV)
    For k = 0 To UBound(AcusticoV)
        GetAllC AcusticoV(k), AcusticoF(k)
    Next k
End Sub

Private Sub Command4_Click()
suma = 0
    For k = 1 To lv1.ListItems.Count
        If lv1.ListItems(k).Selected = True Then
            suma = suma + lv1.ListItems(k).SubItems(5)
        End If
    Next k
    Label1.Caption = suma
End Sub

Private Sub Command5_Click()
    Getting
    setSub
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    SaveSetting "LoadFC", "PATH", "Local", Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
On Error Resume Next
    Dir1.Path = GetSetting("LoadFC", "PATH", "Local")
End Sub

Private Sub lv1_Click()
    Text1.Text = lv1.SelectedItem.SubItems(4)
End Sub
