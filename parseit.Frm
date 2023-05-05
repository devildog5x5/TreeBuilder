VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ConversionTool 
   Caption         =   "User List LDIF Conversion"
   ClientHeight    =   8085
   ClientLeft      =   11070
   ClientTop       =   945
   ClientWidth     =   6705
   Icon            =   "parseit.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8085
   ScaleWidth      =   6705
   Begin VB.CheckBox CheckNewImportFile 
      Caption         =   "Check here if you want to import your own text file"
      Height          =   255
      Left            =   2640
      TabIndex        =   30
      ToolTipText     =   "Check here if you want to import your own text file"
      Top             =   120
      Width           =   3855
   End
   Begin VB.CommandButton ViewTextFileExp 
      Caption         =   "View LDIF File"
      Height          =   375
      Left            =   5280
      TabIndex        =   29
      ToolTipText     =   "View the file created by the conversion tool"
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton ViewTextFileImp 
      Caption         =   "View User File"
      Height          =   375
      Left            =   5280
      TabIndex        =   28
      ToolTipText     =   "View selected import file"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox TextImportBox 
      Height          =   315
      Left            =   1440
      TabIndex        =   26
      Text            =   "C:\Temp\Import.txt"
      ToolTipText     =   "Import file to be converted"
      Top             =   1560
      Width           =   3075
   End
   Begin VB.CommandButton BrowseImportTFile 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   25
      ToolTipText     =   "Browse to select file to import"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox GivenName 
      Height          =   315
      Left            =   1800
      TabIndex        =   24
      Text            =   "John"
      ToolTipText     =   "Enter the user GivenName"
      Top             =   2760
      Width           =   2355
   End
   Begin VB.TextBox Password 
      Height          =   315
      Left            =   1800
      TabIndex        =   23
      Text            =   "test"
      ToolTipText     =   "Enter the user password"
      Top             =   4560
      Width           =   2355
   End
   Begin VB.TextBox Location 
      Height          =   315
      Left            =   1800
      TabIndex        =   22
      Text            =   "Provo"
      ToolTipText     =   "Enter the user Location"
      Top             =   4200
      Width           =   2355
   End
   Begin VB.TextBox SurName 
      Height          =   315
      Left            =   1800
      TabIndex        =   21
      Text            =   "Doe"
      ToolTipText     =   "Enter the user SurName"
      Top             =   3840
      Width           =   2355
   End
   Begin VB.TextBox Telephone 
      Height          =   315
      Left            =   1800
      TabIndex        =   20
      Text            =   "1-(801)555-1212"
      ToolTipText     =   "Enter the user Telephone Number"
      Top             =   3480
      Width           =   2355
   End
   Begin VB.TextBox Title 
      Height          =   315
      Left            =   1800
      TabIndex        =   19
      Text            =   "Mananger"
      ToolTipText     =   "Enter the user Title"
      Top             =   3120
      Width           =   2355
   End
   Begin VB.TextBox UserContext 
      Height          =   315
      Left            =   1800
      TabIndex        =   10
      Text            =   ",ou=Provo,o=Novell"
      ToolTipText     =   "Enter the user Context"
      Top             =   2040
      Width           =   4755
   End
   Begin VB.TextBox Domain 
      Height          =   315
      Left            =   1800
      TabIndex        =   9
      Text            =   "novell.com"
      ToolTipText     =   "Enter the user Domain Name"
      Top             =   2400
      Width           =   2355
   End
   Begin VB.CommandButton BrowseExportTFile 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      ToolTipText     =   "Browse to the file created by the conversion tool"
      Top             =   7680
      Width           =   615
   End
   Begin VB.TextBox TextExportBox 
      Height          =   315
      Left            =   1440
      TabIndex        =   6
      Text            =   "C:\Temp\Export.txt"
      Top             =   7680
      Width           =   3075
   End
   Begin VB.ListBox UserList 
      Height          =   2205
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "List of Users created"
      Top             =   5400
      Width           =   6495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Convert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3960
      TabIndex        =   2
      Text            =   ","
      Top             =   4980
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   915
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "parseit.frx":08CA
      ToolTipText     =   "This is the list of users to be created, must have a leading and trailing comma."
      Top             =   480
      Width           =   6435
   End
   Begin MSComDlg.CommonDialog cmnDialog 
      Left            =   6120
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label12 
      Caption         =   "Import File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   27
      ToolTipText     =   "Select the file to import"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Given Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "Location"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "Surname"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Telephone #"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Domain Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "User Context"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Save to File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Separator(s) Comma is Default"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Select the delimiter"
      Top             =   5040
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "User List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "ConversionTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Const STILL_ACTIVE = &H103
Const PROCESS_QUERY_INFORMATION = &H400

Private Sub Command1_Click()

    Dim A$, i As Integer, P$, k As Integer, B$
    Dim TextImportFile As String
    Dim TextExportFile As String
    Dim domainname As String
    Dim context As String
    Dim givename1 As String
    Dim Title1 As String
    Dim Telephone1 As String
    Dim surname1 As String
    Dim Location1 As String
    Dim password1 As String
    Dim UserName As String
    Dim UserFirstName As String
    Dim UserLastName As String
    Dim User1 As String
    Dim First1 As String
    Dim Last1 As String
    Dim User As String
    Dim First As String
    Dim Last As String

    domainname = Domain.Text
    context = UserContext.Text
    givename1 = GivenName.Text
    Title1 = Title.Text
    Telephone1 = Telephone.Text
    surname1 = SurName.Text
    Location1 = Location.Text
    password1 = Password.Text
    TextExportFile = TextExportBox.Text
    TextImportFile = TextImportBox.Text
    
' This shows how to parse strings (separate words)
'On Error GoTo ErrorHandler
                   If CheckNewImportFile = 1 Then GoTo FParser
            Open TextExportFile For Output As #1
'            On Error GoTo ErrorHandler
                Print #1, "#This LDIF file was generated by Robert Foster's User File Conversion Tool part of Tree Builder version 3.4."
                Print #1, "version: 1"
            Close #1
                    P$ = Text2.Text
                    'A$ = Text1.Text
                    'If CheckNewImportFile = 1 Then GoTo FParser
                    A$ = Text1.Text
                        ' First, make all the separators the same
                        For i = 2 To Len(P$)
                        k = InStr(A$, Mid$(P$, i, 1))
                    Do While k
                        Mid$(A$, k, 1) = Left$(P$, 1)
                        k = InStr(A$, Mid$(P$, i, 1))
                            Loop
                                Next
                                    ' Now parse it!
                                        Do While Len(A$)
                                        k = InStr(A$, Left$(P$, 1))
                                        If k = 1 Then
                                        A$ = Mid$(A$, 2)
                                        ElseIf k Then
                                        B$ = Left$(A$, k - 1)
            
             Open TextExportFile For Append As #1
                'On Error GoTo ErrorHandler
                Print #1,
                Print #1, "dn: cn=" + B$ + context
                Print #1, "changetype: " + "add"
                Print #1, "mail: " + B$ + "@" + domainname
                Print #1, "uid: " + B$
                Print #1, "givenName: " + givename1
                Print #1, "title: " + Title1
                Print #1, "telephoneNumber: " + Telephone1
                Print #1, "sn: " + surname1
                Print #1, "objectclass: inetOrgPerson"
                Print #1, "objectclass: organizationalPerson"
                Print #1, "objectclass: person"
                Print #1, "objectclass: ndsLoginProperties"
                Print #1, "objectclass: top"
                Print #1, "l: " + Location1
                Print #1, "userpassword: " + password1
             Close #1
            
            UserList.AddItem B$
            A$ = Mid$(A$, k + 1)
        Else
            '***UserList.AddItem A$
            A$ = ""
        End If
    Loop
    
FParser: 'If a file is used it is processed here.
    
    If CheckNewImportFile = 0 Then GoTo FParserEnd
            Open TextExportFile For Output As #1
            On Error GoTo ErrorHandler
                Print #1, "#This LDIF file was generated by Robert Foster's User File Conversion Tool part of Tree Builder version 3.4."
                Print #1, "version: 1"
            Close #1
                    P$ = Text2.Text
                    'A$ = Text1.Text
'***                    Open TextImportFile For Output As #3
                        '***Do While FileExists(TextImportFile) = "False" 'Make sure the file exists
                        
                        If FileExists(TextImportFile) = "False" Then cmnDialog.ShowOpen 'Make sure the file exists
                        TextImportBox.Text = cmnDialog.FileName
                        TextImportFile = cmnDialog.FileName
                        If TextImportBox.Text = "" Then TextImportBox.Text = "C:\Temp\Import.txt"
                        On Error GoTo ErrorHandler
    Sleep 100
    Do While FileExists(TextImportFile) = "False" 'Make sure the file exists
    DoEvents: Sleep 100
    Loop
Open TextImportFile For Input As #3
         While Not EOF(3)
      Line Input #3, A$
                        '***DoEvents: Sleep 100
                            '***Loop
                                'Open TextImportFile For Input As #3
                                'On Error GoTo ErrorHandler
                                    'ParsedText = "'"
                        'While Not EOF(TextImportFile)
                        'Line Input #3, A$
'***
                    'A$ = Text1.Text
                        ' First, make all the separators the same
                        For i = 2 To Len(P$)
                        k = InStr(A$, Mid$(P$, i, 1))
                    Do While k
                        Mid$(A$, k, 1) = Left$(P$, 1)
                        k = InStr(A$, Mid$(P$, i, 1))
                            Loop
                                Next
                                    ' Now parse it!
                                        Do While Len(A$)
                                        k = InStr(A$, Left$(P$, 1))
                                        If k = 1 Then
                                        A$ = Mid$(A$, 2)
                                        ElseIf k Then
                                        B$ = Left$(A$, k - 1)

                                        If B$ Like "USER0*" = True Then User1 = B$
                                        If B$ Like "FIRST0*" = True Then First1 = B$
                                        If B$ Like "LAST0*" = True Then Last1 = B$

            UserList.AddItem B$
            A$ = Mid$(A$, k + 1)
    If B$ Like "USER0*" = True And B$ Like "FIRST0*" = True And B$ Like "LAST0*" = True Then GoTo JmpUF
    If B$ Like "USER0*" = True Or B$ Like "FIRST0*" = True Or B$ Like "LAST0*" = True Then GoTo JmpLp
               '***
               If B$ = "" Then GoTo JPrnt
             Open TextExportFile For Append As #1
                Print #1,
                Print #1, "dn: cn=" + B$ + context
                Print #1, "changetype: " + "add"
                Print #1, "mail: " + B$ + "@" + domainname
                Print #1, "uid: " + B$
                Print #1, "givenName: " + B$
                Print #1, "title: " + Title1
                Print #1, "telephoneNumber: " + Telephone1
                Print #1, "sn: " + B$
                Print #1, "objectclass: inetOrgPerson"
                Print #1, "objectclass: organizationalPerson"
                Print #1, "objectclass: person"
                Print #1, "objectclass: ndsLoginProperties"
                Print #1, "objectclass: top"
                Print #1, "l: " + Location1
                Print #1, "userpassword: " + password1
             Close #1
JPrnt:
JmpLp:
        '***
        Else
        
            '***UserList.AddItem A$
            A$ = ""

        End If '***


    Loop
JmpUF:
'***
          If User1 = "" Then GoTo JPrnt1
             Open TextExportFile For Append As #1
                On Error GoTo ErrorHandler
                Print #1,
                Print #1, "dn: cn=" + User1 + context
                Print #1, "changetype: " + "add"
                Print #1, "mail: " + User1 + "@" + domainname
                Print #1, "uid: " + User1
                Print #1, "givenName: " + First1
                Print #1, "title: " + Title1
                Print #1, "telephoneNumber: " + Telephone1
                Print #1, "sn: " + Last1
                Print #1, "objectclass: inetOrgPerson"
                Print #1, "objectclass: organizationalPerson"
                Print #1, "objectclass: person"
                Print #1, "objectclass: ndsLoginProperties"
                Print #1, "objectclass: top"
                Print #1, "l: " + Location1
                Print #1, "userpassword: " + password1
             Close #1
JPrnt1:
        '***
   Wend
Close #3   ' Close file.
FParserEnd:
On Error GoTo ErrorResume   'Resume next in case file is locked
ErrorResume:
Resume Next   'in case file is locked
ErrorHandler:
If Err.Description = "" Then Resume Next Else MsgBox Err.Description
 End Sub
Private Sub BrowseExportTFile_Click()
cmnDialog.FileName = TextExportBox.Text
cmnDialog.ShowOpen
TextExportBox.Text = cmnDialog.FileName
If TextExportBox.Text = "" Then TextExportBox.Text = "C:\Temp\TExport.txt"
End Sub

Private Sub ViewTextFileExp_Click()
      Dim RetVal As Long
      Dim filein As String
      filein = TextExportBox.Text
      RetVal = ShellExecute(0, "open", "notepad", filein, "", SW_SHOW)
End Sub

Private Sub BrowseImportTFile_Click()
cmnDialog.FileName = TextImportBox.Text
cmnDialog.ShowOpen
TextImportBox.Text = cmnDialog.FileName
If TextImportBox.Text = "" Then TextImportBox.Text = "C:\Temp\Import.txt"
End Sub

Private Sub ViewTextFileImp_Click()
      Dim RetVal As Long
      Dim filein As String
      filein = TextImportBox.Text
      RetVal = ShellExecute(0, "open", "notepad", filein, "", SW_SHOW)
End Sub
