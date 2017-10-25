VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "绘制滞回曲线动画－清华大学陆新征"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8535
   DrawMode        =   15  'Merge Pen Not
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   8535
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   7200
      TabIndex        =   28
      Text            =   "255"
      Top             =   5880
      Width           =   855
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   5280
      TabIndex        =   26
      Text            =   "255"
      Top             =   5880
      Width           =   855
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   3240
      TabIndex        =   24
      Text            =   "0"
      Top             =   5880
      Width           =   855
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   1080
      TabIndex        =   22
      Text            =   "1"
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "播放动画"
      Height          =   495
      Left            =   7320
      TabIndex        =   21
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   7200
      TabIndex        =   20
      Text            =   "Text8"
      Top             =   5280
      Width           =   855
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   7200
      TabIndex        =   18
      Text            =   "Text7"
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "设定"
      Height          =   495
      Left            =   7200
      TabIndex        =   16
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   5280
      TabIndex        =   15
      Text            =   "Text6"
      Top             =   5280
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   5280
      TabIndex        =   13
      Text            =   "Text5"
      Top             =   4800
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3240
      TabIndex        =   11
      Text            =   "Text4"
      Top             =   5280
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   4800
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "序列存盘"
      Height          =   495
      Left            =   7320
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7680
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "读入数据"
      Height          =   495
      Left            =   7320
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   240
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   213
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "图像存盘"
      Height          =   495
      Left            =   7320
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "颜色－蓝"
      Height          =   375
      Left            =   6360
      TabIndex        =   29
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label Label11 
      Caption         =   "颜色－绿"
      Height          =   375
      Left            =   4440
      TabIndex        =   27
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "颜色－红"
      Height          =   375
      Left            =   2400
      TabIndex        =   25
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "线宽"
      Height          =   375
      Left            =   240
      TabIndex        =   23
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "输出步数"
      Height          =   375
      Left            =   6360
      TabIndex        =   19
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "标题"
      Height          =   375
      Left            =   6360
      TabIndex        =   17
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Y最大值"
      Height          =   255
      Left            =   4440
      TabIndex        =   14
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Y最小值"
      Height          =   255
      Left            =   4440
      TabIndex        =   12
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "X最大值"
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "X最小值"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "图像高度"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "图像宽度"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   4800
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MaxX, MaxY, MinX, MinY As Double
'Dim PicWidth, PicHeight As Integer
Dim Scale_X, Scale_Y As Double
Dim X(1 To 10000), Y(1 To 10000) As Double
Dim PicX(1 To 10000), PicY(1 To 10000) As Integer
Dim NPoint As Integer
Dim I, J As Integer
Dim Original_X, Original_Y As Integer
Dim LineWidth As Integer
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Sub Command1_Click()
    SavePicture Picture1.Image, App.Path & "\1.bmp"
End Sub

Private Sub Command2_Click()
    Dim X1, X2, Y1, Y2 As Double
    CommonDialog1.InitDir = App.Path
    CommonDialog1.Filter = "文本文件|*.txt"
    CommonDialog1.Action = 1
    CommonDialog1.DialogTitle = "打开文件"
    If CommonDialog1.FileName = "" Then Exit Sub  '加上个判定，免得取消打开时下面的OPEN出错
    Open CommonDialog1.FileName For Input As #1
    
    X1 = 0
    Y1 = 0
    X2 = 0
    Y2 = 0
    I = 1
    Do While Not EOF(1)
        Input #1, J, X(I), Y(I)
        If X(I) < X1 Then
            X1 = X(I)
        End If
        If X(I) > X2 Then
            X2 = X(I)
        End If
        If Y(I) < Y1 Then
            Y1 = Y(I)
        End If
        If Y(I) > Y2 Then
            Y2 = Y(I)
        End If
        NPoint = I
        I = I + 1
    Loop
    Close #1
    
    If MinX = 0 Then '赋值
        MinX = X1
        MinY = Y1
        MaxX = X2
        MaxY = Y2
        Scale_X = Picture1.ScaleWidth / (MaxX - MinX) * 0.9
        Scale_Y = Picture1.ScaleHeight / (MaxY - MinY) * 0.9
        Text3.Text = Str(MinX)
        Text4.Text = Str(MaxX)
        Text5.Text = Str(MinY)
        Text6.Text = Str(MaxY)
        Text8.Text = Str(NPoint)
    End If

    
    Call GetPicCoord
    Call DrawLines
    
End Sub

Private Sub Command3_Click()
    Dim TextString As String
    For J = 1 To NPoint
        Picture1.Cls
        Call DrawLines
        Call DrawDot(J)
        TextString = Format(Str$(J), "000000")
        SavePicture Picture1.Image, App.Path & "\" & TextString & ".bmp"
    Next J
    
    If Int(Val(Text8.Text)) > NPoint Then
    For J = NPoint + 1 To Val(Text8.Text)
        Picture1.Cls
        Call DrawLines
        Call DrawDot(NPoint)
        TextString = Format(Str$(J), "000000")
        SavePicture Picture1.Image, App.Path & "\" & TextString & ".bmp"
    Next J
    End If
    
End Sub

Private Sub Command4_Click()
    
    Picture1.AutoSize = True
    Picture1.Width = Picture1.Width / Picture1.ScaleWidth * Val(Text1.Text)
    Picture1.Height = Picture1.Height / Picture1.ScaleHeight * Val(Text2.Text)
    Picture1.ScaleWidth = Val(Text1.Text)
    Picture1.ScaleHeight = Val(Text2.Text)
    
    
    MinX = Val(Text3.Text)
    MaxX = Val(Text4.Text)
    MinY = Val(Text5.Text)
    MaxY = Val(Text6.Text)
    LineWidth = Val(Text9.Text)
    Scale_X = Picture1.ScaleWidth / (MaxX - MinX) * 0.9
    Scale_Y = Picture1.ScaleHeight / (MaxY - MinY) * 0.9
    

    Call GetPicCoord
    
    Call DrawLines

End Sub

Private Sub Command5_Click()
    For J = 1 To NPoint
        Picture1.Cls
        Call DrawLines
        Call DrawDot(J)
        Savetime = timeGetTime '记下开始时的时间
        While timeGetTime < Savetime + 25 '循环等待
            DoEvents '转让控制权，以便让操作系统处理其它的事件。
        Wend
    Next J
End Sub

Private Sub Form_Load()
    Text1.Text = Str(Picture1.ScaleWidth)
    Text2.Text = Str(Picture1.ScaleHeight)
    LineWidth = 1
    
'    MaxX = 0.01
'    MinX = -0.01
'    MaxY = 400
'    MinY = -400
'    PicWidth = Picture1.ScaleWidth
'    PicHeight = Picture1.ScaleHeight


End Sub


Private Sub DrawLines()
    Picture1.Cls
    Original_X = (0 - MinX) * Scale_X + 0.05 * Picture1.ScaleWidth
    Original_Y = Picture1.ScaleHeight + ((MinY) * Scale_Y - 0.05 * Picture1.ScaleHeight)
    
    Picture1.Line (Original_X, Original_Y)-(0, Original_Y), RGB(255, 255, 255)
    Picture1.Line (Original_X, Original_Y)-(Picture1.ScaleWidth, Original_Y), RGB(255, 255, 255)
    Picture1.Line (Original_X, Original_Y)-(Original_X, 0), RGB(255, 255, 255)
    Picture1.Line (Original_X, Original_Y)-(Original_X, Picture1.ScaleHeight), RGB(255, 255, 255)
    
    Picture1.DrawWidth = LineWidth
    For I = 1 To NPoint - 1
        X1 = PicX(I)
        Y1 = PicY(I)
        Picture1.Line (PicX(I), PicY(I))-(PicX(I + 1), PicY(I + 1)), RGB(Val(Text10.Text), Val(Text11.Text), Val(Text12.Text))
    Next I
    Picture1.DrawWidth = 1
    
    x0 = Picture1.TextWidth(Text7.Text)
    
    Picture1.PSet (Picture1.ScaleWidth / 2 - x0 / 2, 0), RGB(0, 0, 0)
    Picture1.Print Text7.Text


End Sub

Private Sub DrawDot(NP)
    Picture1.FillStyle = 0
    Picture1.FillColor = RGB(255, 0, 0)
    Picture1.Circle (PicX(NP), PicY(NP)), Picture1.ScaleWidth / 50, RGB(255, 0, 0)

End Sub

Private Sub GetPicCoord()
    For I = 1 To NPoint
        PicX(I) = X(I) * Scale_X
        PicX(I) = PicX(I) + (0 - MinX) * Scale_X + 0.05 * Picture1.ScaleWidth
        PicY(I) = Y(I) * Scale_Y
        PicY(I) = Picture1.ScaleHeight + (-PicY(I) + (MinY) * Scale_Y - 0.05 * Picture1.ScaleHeight)
    Next I

End Sub

