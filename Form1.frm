VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Weather"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3000
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   3000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Feels like"
      Height          =   240
      Left            =   1155
      TabIndex        =   12
      Top             =   3120
      Width           =   1230
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Refresh"
      Height          =   330
      Left            =   750
      TabIndex        =   11
      Top             =   1110
      Width           =   1425
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Humidity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1125
      TabIndex        =   4
      Top             =   1995
      Width           =   1230
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Condition"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1200
      TabIndex        =   3
      Top             =   2370
      Width           =   1140
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2790
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2640
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Temp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1215
      TabIndex        =   1
      Top             =   2760
      Width           =   1140
   End
   Begin VB.TextBox Text 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3435
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   2820
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label lblFeel 
      Height          =   225
      Left            =   1020
      TabIndex        =   14
      Top             =   300
      Width           =   1845
   End
   Begin VB.Label Label4 
      Caption         =   "Feels like:"
      Height          =   195
      Left            =   90
      TabIndex        =   13
      Top             =   300
      Width           =   1650
   End
   Begin VB.Label lblHumid 
      Height          =   270
      Left            =   1005
      TabIndex        =   10
      Top             =   735
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Humidity:"
      Height          =   225
      Left            =   75
      TabIndex        =   9
      Top             =   735
      Width           =   1845
   End
   Begin VB.Label lblCond 
      Height          =   240
      Left            =   1800
      TabIndex        =   8
      Top             =   510
      Width           =   2280
   End
   Begin VB.Label Label2 
      Caption         =   "Current Condition:"
      Height          =   240
      Left            =   75
      TabIndex        =   7
      Top             =   510
      Width           =   1950
   End
   Begin VB.Label lblTemp 
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1425
      TabIndex        =   6
      Top             =   75
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Temperature:"
      Height          =   240
      Left            =   90
      TabIndex        =   5
      Top             =   75
      Width           =   1545
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Command1_Click()
On Error Resume Next
Dim search
Dim spot As Integer
Dim spot2 As Integer
Dim done
search = "insert current temp"
spot = InStr(Text, search) + 23
done = Mid(Text, spot, 20)
spot2 = InStr(done, "</B>")
done = Mid(done, 1, spot2 - 1)
done = pReplace(Trim(done), "&deg;", "° ")

lblTemp = done
If pReplace(Trim(done), "° F", "") > 75 Then
lblTemp.ForeColor = &HFF&
Else
lblTemp.ForeColor = &HFF0000
End If
End Sub

Public Sub Command2_Click()
On Error Resume Next
Dim search
Dim spot As Integer
Dim spot2 As Integer
Dim done
search = "insert forecast text"
spot = InStr(Text, search) + 24
done = Mid(Text, spot, 100)
spot2 = InStr(done, "</td>")
done = Mid(done, 1, spot2 - 1)
done = pReplace(Trim(done), "&deg;", "° ")

lblCond = done
End Sub

Public Sub Command3_Click()
On Error Resume Next
Dim search
Dim spot As Integer
Dim spot2 As Integer
Dim done
search = "insert humidity"
spot = InStr(Text, "insert humidity") + 21
done = Mid(Text, spot - 2, 100)
spot2 = InStr(done, "</td>")
done = Mid(done, 1, spot2 - 1)
Dim Zip
done = Trim(done)

lblHumid = done
End Sub

Private Sub Command4_Click()
lblTemp = ""
lblHumid = ""
lblCond = ""

Command4.Enabled = False
Form_Load
Command4.Enabled = True
End Sub

Public Sub Command5_Click()
On Error Resume Next
Dim search
Dim spot As Integer
Dim spot2 As Integer
Dim done
search = "insert feels like temp"
spot = InStr(Text, search) + 26
done = Mid(Text, spot, 20)
spot2 = InStr(done, "</font>")
done = Mid(done, 1, spot2 - 1)
done = pReplace(Trim(done), "&deg;", "° ")

done = pReplace(Trim(done), "Feels Like: ", "")

lblFeel = done
If pReplace(Trim(done), "° F", "") > 75 Then
lblFeel.ForeColor = &HFF&
Else
lblFeel.ForeColor = &HFF0000
End If
End Sub

Public Sub Form_Load()
Zip = "11756"
Text = GetUrlSource("http://www.weather.com/weather/local/11756") ' where 11756 is, insert zip code
Command1_Click ' check out the command button's code for parsing
Command2_Click
Command3_Click
Command5_Click
End Sub

