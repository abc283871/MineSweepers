VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   5220
   StartUpPosition =   3  '系統預設值
   Begin VB.Label score 
      BackColor       =   &H00C0C0FF&
      Height          =   615
      Left            =   1680
      TabIndex        =   10
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Height          =   735
      Index           =   8
      Left            =   3480
      TabIndex        =   9
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Height          =   735
      Index           =   7
      Left            =   2160
      TabIndex        =   8
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Height          =   735
      Index           =   6
      Left            =   840
      TabIndex        =   7
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Height          =   735
      Index           =   5
      Left            =   3480
      TabIndex        =   6
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Height          =   735
      Index           =   4
      Left            =   2160
      TabIndex        =   5
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Height          =   735
      Index           =   3
      Left            =   840
      TabIndex        =   4
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Height          =   735
      Index           =   2
      Left            =   3480
      TabIndex        =   3
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Height          =   735
      Index           =   1
      Left            =   2160
      TabIndex        =   2
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Height          =   735
      Index           =   0
      Left            =   840
      TabIndex        =   1
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label tghj 
      BackColor       =   &H00C0C0FF&
      Caption         =   "score"
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim 已抽地雷 As Integer
Dim 被抽過的格子(0 To 8) As Integer
Dim 抽籤 As Integer
Dim i As Integer
Dim p As Integer


Private Sub Form_Load()
已抽地雷 = 0
Randomize

Do
    抽籤 = Int(Rnd * 9)
    Print 抽籤
        If 被抽過的格子(抽籤) = 0 Then
            Label4(抽籤).Tag = "*"
            被抽過的格子(抽籤) = 1
            已抽地雷 = 已抽地雷 + 1
        End If
Loop Until 已抽地雷 = 2

For i = 0 To 8
    p = 0
    If Label4(i).Tag <> "*" Then
        If (i - 4) >= 0 And (i - 4) Mod 3 <> 2 Then
            If Label4(i - 4).Tag = "*" Then
                p = p + 1
            End If
        End If
        
         If (i - 3) >= 0 Then
            If Label4(i - 3).Tag = "*" Then
                p = p + 1
            End If
        End If
        
        If (i - 2) > 0 And (i - 2) Mod 3 <> 0 Then
            If Label4(i - 2).Tag = "*" Then
                p = p + 1
            End If
        End If
        
        If (i - 1) > 0 And (i - 1) Mod 3 <> 2 Then
            If Label4(i - 1).Tag = "*" Then
                p = p + 1
            End If
        End If
        
        If (i + 1) > 0 And (i + 1) < 9 And (i + 1) Mod 3 <> 0 Then
            If Label4(i + 1).Tag = "*" Then
                p = p + 1
            End If
        End If
        
        If (i + 2) > 0 And (i + 2) < 9 And (i + 2) Mod 3 <> 2 Then
            If Label4(i + 2).Tag = "*" Then
                p = p + 1
            End If
        End If
        
        If (i + 3) > 0 And (i + 3) < 9 Then
            If Label4(i + 3).Tag = "*" Then
                p = p + 1
            End If
        End If
        
        If (i + 4) > 0 And (i + 4) < 9 And (i + 4) Mod 3 <> 0 Then
            If Label4(i + 4).Tag = "*" Then
                p = p + 1
            End If
        End If
        Label4(i).Tag = p
    End If
Next i
    
End Sub

Private Sub Label4_Click(Index As Integer)

'如果這個標籤已經踩過，就不做任何處理
If Label4(Index).BorderStyle = 1 Then
    Exit Sub
End If

'未踩-踩過
Label4(Index).Caption = Label4(Index).Tag
Label4(Index).BorderStyle = 1

'中地雷的話
If Label4(Index).Caption = "*" Then
    Label4(Index).BackColor = RGB(255, 0, 0)
    MsgBox ("呵呵~~笑你~~")
    MsgBox ("踩到地雷，程式結束")
    End
End If

'分數
score.Caption = Val(score.Caption) + 1

'踩完了
If score.Caption = 9 - 2 Then
    MsgBox ("恭喜你過關，可惜沒獎品XD")
    MsgBox ("小朋友等級已通關")
    MsgBox ("程式結束，請多修行!!!")
    End
End If


End Sub

