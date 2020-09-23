VERSION 5.00
Object = "*\AjcFrames.vbp"
Begin VB.Form Demo 
   BackColor       =   &H00C9ECEF&
   Caption         =   "jcFrames (version 2.0.1)"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9435
   Icon            =   "Demo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   462
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   629
   StartUpPosition =   2  'CenterScreen
   Begin jcFramesOCX.jcFrames jcFrames 
      Height          =   1905
      Left            =   5460
      Top             =   150
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   3360
      BackColor       =   14933984
      FillColor       =   0
      TextBoxColor    =   0
      Caption         =   "jcFrames2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      HeaderStyle     =   1
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "When ""enabled"" property is false usercontrol automatically disables all the contained controls"
         ForeColor       =   &H00C00000&
         Height          =   465
         Left            =   120
         TabIndex        =   48
         Top             =   840
         Width           =   3645
      End
      Begin VB.Label lblBtn 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Exit Demo"
         Height          =   285
         Left            =   1380
         MouseIcon       =   "Demo.frx":08CA
         MousePointer    =   99  'Custom
         TabIndex        =   45
         ToolTipText     =   "Exit this Demo project"
         Top             =   1440
         Width           =   1125
      End
      Begin VB.Shape ShpBtn 
         BorderColor     =   &H00808080&
         FillColor       =   &H00B0EFF0&
         Height          =   345
         Left            =   1380
         Shape           =   4  'Rounded Rectangle
         Top             =   1380
         Width           =   1125
      End
   End
   Begin jcFramesOCX.jcFrames jcFrames7 
      Height          =   4725
      Left            =   120
      Top             =   2190
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   8334
      BackColor       =   -2147483633
      FillColor       =   14745599
      TextBoxColor    =   0
      Style           =   3
      Caption         =   "jcFrames 2.0.1 - main features"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Demo.frx":0A1C
      IconSize        =   32
      ThemeColor      =   4
      ColorFrom       =   49344
      ColorTo         =   14745599
      HeaderStyle     =   1
      Begin jcFramesOCX.jcFrames jcFrames3 
         Height          =   3885
         Left            =   2490
         Top             =   690
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   6853
         FrameColor      =   6974058
         BackColor       =   14745599
         FillColor       =   14745599
         TextBoxColor    =   0
         Style           =   0
         RoundedCornerTxtBox=   -1  'True
         Caption         =   "General features"
         Alignment       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColorFrom       =   14745599
         ColorTo         =   8421631
         HeaderStyle     =   1
         Begin VB.CheckBox Check1 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Usercontrol enabled"
            Height          =   255
            Left            =   4410
            TabIndex        =   49
            Top             =   0
            Value           =   1  'Checked
            Width           =   1785
         End
         Begin VB.ComboBox CboTxtBoxShadow 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Demo.frx":12F6
            Left            =   1590
            List            =   "Demo.frx":1300
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   1845
            Width           =   1455
         End
         Begin VB.ComboBox CboIconAlign 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Demo.frx":1317
            Left            =   1590
            List            =   "Demo.frx":1321
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   3360
            Width           =   1455
         End
         Begin VB.ComboBox cboGradientHeaderStyle 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Demo.frx":1344
            Left            =   4920
            List            =   "Demo.frx":1354
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   2970
            Width           =   1455
         End
         Begin VB.ComboBox CboHeaderStyle 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Demo.frx":1384
            Left            =   4920
            List            =   "Demo.frx":138E
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   2595
            Width           =   1455
         End
         Begin VB.ComboBox CboFrameColor 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Demo.frx":13A7
            Left            =   4920
            List            =   "Demo.frx":13DC
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   2220
            Width           =   1455
         End
         Begin VB.ComboBox CboThemeColor 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Demo.frx":1413
            Left            =   4920
            List            =   "Demo.frx":1429
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   1080
            Width           =   1455
         End
         Begin VB.ComboBox CboColorFrom 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Demo.frx":1462
            Left            =   4920
            List            =   "Demo.frx":149B
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   1455
            Width           =   1455
         End
         Begin VB.ComboBox CboColorTo 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Demo.frx":14E3
            Left            =   4920
            List            =   "Demo.frx":151C
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   1845
            Width           =   1455
         End
         Begin VB.ComboBox CboRoundTxtBox 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Demo.frx":1569
            Left            =   4920
            List            =   "Demo.frx":1573
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   705
            Width           =   1455
         End
         Begin VB.ComboBox CboTextBoxColor 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Demo.frx":1584
            Left            =   1590
            List            =   "Demo.frx":15DF
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1455
            Width           =   1455
         End
         Begin VB.ComboBox CboFillColor 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Demo.frx":1630
            Left            =   4920
            List            =   "Demo.frx":167D
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   330
            Width           =   1455
         End
         Begin VB.ComboBox CboTextColor 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Demo.frx":16CD
            Left            =   1590
            List            =   "Demo.frx":1700
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   1080
            Width           =   1455
         End
         Begin VB.ComboBox CboCaptionAlig 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Demo.frx":172F
            Left            =   1590
            List            =   "Demo.frx":173C
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   330
            Width           =   1455
         End
         Begin VB.ComboBox CboRoundCorner 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Demo.frx":1769
            Left            =   1590
            List            =   "Demo.frx":1773
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   705
            Width           =   1455
         End
         Begin VB.ComboBox CboPicture 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Demo.frx":1784
            Left            =   1590
            List            =   "Demo.frx":178E
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   2595
            Width           =   1455
         End
         Begin VB.ComboBox CboIconSize 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Demo.frx":179B
            Left            =   1590
            List            =   "Demo.frx":17A8
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   2970
            Width           =   1455
         End
         Begin VB.ComboBox CboTextBoxHeight 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Demo.frx":17B8
            Left            =   1590
            List            =   "Demo.frx":17D7
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   2220
            Width           =   1455
         End
         Begin VB.TextBox txtCaption 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4140
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   3360
            Width           =   2205
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TextBoxShadow:"
            Height          =   195
            Left            =   240
            TabIndex        =   47
            Top             =   1905
            Width           =   1215
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Icon Alignment:"
            Height          =   195
            Left            =   240
            TabIndex        =   44
            Top             =   3420
            Width           =   1095
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gradient header style:"
            Height          =   195
            Left            =   3300
            TabIndex        =   42
            Top             =   3030
            Width           =   1545
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Header style:"
            Height          =   195
            Left            =   3300
            TabIndex        =   40
            Top             =   2655
            Width           =   930
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FrameColor:"
            Height          =   195
            Left            =   3300
            TabIndex        =   38
            Top             =   2280
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ThemeColor:"
            Height          =   195
            Left            =   3300
            TabIndex        =   36
            Top             =   1140
            Width           =   900
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ColorFrom:"
            Height          =   195
            Left            =   3300
            TabIndex        =   35
            Top             =   1515
            Width           =   750
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ColorTo:"
            Height          =   195
            Left            =   3300
            TabIndex        =   34
            Top             =   1905
            Width           =   600
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TextBoxColor:"
            Height          =   195
            Left            =   240
            TabIndex        =   30
            Top             =   1515
            Width           =   990
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RoundCornerTxtBox:"
            Height          =   195
            Left            =   3300
            TabIndex        =   29
            Top             =   765
            Width           =   1485
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FillColor:"
            Height          =   195
            Left            =   3300
            TabIndex        =   28
            Top             =   390
            Width           =   585
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TextColor:"
            Height          =   195
            Left            =   240
            TabIndex        =   24
            Top             =   1140
            Width           =   720
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Caption Alignment:"
            Height          =   195
            Left            =   240
            TabIndex        =   23
            Top             =   390
            Width           =   1320
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RoundedCorner:"
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   765
            Width           =   1170
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Picture:"
            Height          =   195
            Left            =   240
            TabIndex        =   21
            Top             =   2655
            Width           =   540
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IconSize:"
            Height          =   195
            Left            =   240
            TabIndex        =   20
            Top             =   3030
            Width           =   660
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TextBoxHeight:"
            Height          =   195
            Left            =   240
            TabIndex        =   19
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Caption:"
            Height          =   195
            Left            =   3300
            TabIndex        =   18
            Top             =   3390
            Width           =   585
         End
      End
      Begin jcFramesOCX.jcFrames jcFrames1 
         Height          =   3960
         Left            =   120
         Top             =   615
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   6985
         FrameColor      =   6974058
         BackColor       =   14745599
         FillColor       =   14745599
         TextBoxColor    =   11595760
         Style           =   2
         RoundedCornerTxtBox=   -1  'True
         Caption         =   "Frame styles"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColorFrom       =   4304047
         ColorTo         =   14745599
         GradientHeaderStyle=   1
         Begin VB.OptionButton OptCaptionStyle 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0FFFF&
            Caption         =   "Panel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   8
            Left            =   120
            TabIndex        =   8
            Tag             =   "Panel style"
            Top             =   3540
            Width           =   1965
         End
         Begin VB.OptionButton OptCaptionStyle 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0FFFF&
            Caption         =   "Header"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   7
            Left            =   120
            TabIndex        =   7
            Tag             =   "Header style"
            Top             =   3165
            Width           =   1965
         End
         Begin VB.OptionButton OptCaptionStyle 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0FFFF&
            Caption         =   "Outer Wedge"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   6
            Left            =   120
            TabIndex        =   6
            Tag             =   "Outer wedge style"
            Top             =   2790
            Width           =   1965
         End
         Begin VB.OptionButton OptCaptionStyle 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0FFFF&
            Caption         =   "Inner Wedge"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   5
            Left            =   120
            TabIndex        =   5
            Tag             =   "Inner wedge style"
            Top             =   2415
            Width           =   1965
         End
         Begin VB.OptionButton OptCaptionStyle 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0FFFF&
            Caption         =   "Messenger"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   4
            Left            =   120
            TabIndex        =   4
            Tag             =   "Messenger style"
            Top             =   2040
            Width           =   2085
         End
         Begin VB.OptionButton OptCaptionStyle 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0FFFF&
            Caption         =   "Windows"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   3
            Left            =   120
            TabIndex        =   3
            Tag             =   "Windows style"
            Top             =   1665
            Width           =   2085
         End
         Begin VB.OptionButton OptCaptionStyle 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0FFFF&
            Caption         =   "TextBox (from EZFrame)"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   2
            Left            =   120
            TabIndex        =   2
            Tag             =   "Text box style"
            ToolTipText     =   "Thanks to ElectroZ for his frame style"
            Top             =   1290
            Width           =   2085
         End
         Begin VB.OptionButton OptCaptionStyle 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0FFFF&
            Caption         =   "jcGradient"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   1
            Tag             =   "jcGradient style"
            Top             =   915
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton OptCaptionStyle 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0FFFF&
            Caption         =   "XpDefault"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   0
            Tag             =   "Xp default style"
            Top             =   540
            Width           =   1815
         End
      End
   End
   Begin jcFramesOCX.jcFrames jcFrames4 
      Height          =   1035
      Left            =   2400
      Top             =   630
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1826
      BackColor       =   -2147483633
      FillColor       =   0
      TextBoxColor    =   0
      Style           =   8
      Caption         =   ""
      TextColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   5
      ColorFrom       =   14215660
      ColorTo         =   12632256
      HeaderStyle     =   1
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "jcFrames 2.0.1 with 9 styles"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H006A6A6A&
         Height          =   825
         Left            =   150
         TabIndex        =   9
         Top             =   120
         Width           =   2025
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   2520
         Picture         =   "Demo.frx":17FF
         Top             =   390
         Width           =   240
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "jcFrames 2.0.1 with 9 styles"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   825
         Left            =   165
         TabIndex        =   10
         Top             =   135
         Width           =   2025
      End
   End
   Begin VB.Image Image2 
      Height          =   1800
      Left            =   240
      Picture         =   "Demo.frx":1D89
      Top             =   210
      Width           =   1800
   End
End
Attribute VB_Name = "Demo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim blnFormLoaded As Boolean

Private Sub CboCaptionAlig_Click()
    If blnFormLoaded = True Then
        If CboCaptionAlig.ListIndex = -1 Then Exit Sub
        jcFrames.Alignment = CboCaptionAlig.ListIndex
    End If
End Sub

Private Sub CboColorFrom_Click()
    If CboColorFrom.ListIndex = -1 Then Exit Sub
    If CboColorFrom.ListIndex = 5 Then
        jcFrames.ColorFrom = jcFrames.BackColor ' &H8000000F
    Else
        jcFrames.ColorFrom = CboColorFrom.ItemData(CboColorFrom.ListIndex)
    End If
End Sub

Private Sub CboColorTo_Click()
    If CboColorTo.ListIndex = -1 Then Exit Sub
    If CboColorTo.ListIndex = 5 Then
        jcFrames.ColorTO = jcFrames.BackColor ' &H8000000F
    Else
        jcFrames.ColorTO = CboColorTo.ItemData(CboColorTo.ListIndex)
    End If
End Sub

Private Sub CboFillColor_Click()
    If blnFormLoaded = True Then
        If CboFillColor.ListIndex = -1 Then Exit Sub
        If CboFillColor.ListIndex = 4 Then
            jcFrames.FillColor = jcFrames.BackColor ' &H8000000F
        Else
            jcFrames.FillColor = CboFillColor.ItemData(CboFillColor.ListIndex)
        End If
    End If
End Sub

Private Sub CboFrameColor_Click()
    If blnFormLoaded = True Then
        If CboFrameColor.ListIndex = -1 Then Exit Sub
        If CboFrameColor.ListIndex = 6 Then
            jcFrames.FrameColor = jcFrames.BackColor ' &H8000000F
        Else
            jcFrames.FrameColor = CboFrameColor.ItemData(CboFrameColor.ListIndex)
        End If
    End If
End Sub

Private Sub cboGradientHeaderStyle_Click()
    If cboGradientHeaderStyle.ListIndex = -1 Then Exit Sub
    If blnFormLoaded = True Then
        jcFrames.GradientHeaderStyle = cboGradientHeaderStyle.ListIndex
    End If
End Sub

Private Sub CboHeaderStyle_Click()
    If blnFormLoaded = True Then
        jcFrames.HeaderStyle = CboHeaderStyle.ListIndex
        If CboHeaderStyle.ListIndex = 1 Then
            cboGradientHeaderStyle.Enabled = True
            Label19.Enabled = True
            CboThemeColor.Enabled = True
            Label1.Enabled = True
            If jcFrames.Style = Panel Then
                CboThemeColor.ListIndex = 1
                cboGradientHeaderStyle.ListIndex = 2
                CboFillColor.Enabled = False
                CboFillColor.ListIndex = -1
                Label9.Enabled = False
            End If
            If jcFrames.Style = jcGradient Then
                CboThemeColor.ListIndex = 2
                cboGradientHeaderStyle.ListIndex = 0
                CboFillColor.Enabled = False
                CboFillColor.ListIndex = -1
                Label9.Enabled = False
                CboFrameColor.Enabled = False
                CboFrameColor.ListIndex = -1
                Label2.Enabled = False
            End If
            If jcFrames.Style = Windows Or jcFrames.Style = TextBox Then
                CboTextBoxColor.ListIndex = -1
                Label3.Enabled = False
                CboTextBoxColor.Enabled = False
                CboThemeColor.ListIndex = 4
                If jcFrames.Style = TextBox Then
                    cboGradientHeaderStyle.ListIndex = 2
                Else
                    cboGradientHeaderStyle.ListIndex = 1
                End If
                If CboRoundTxtBox.ListIndex = 1 Then CboRoundTxtBox.ListIndex = 0
            End If
            If jcFrames.Style = Messenger Then
                cboGradientHeaderStyle.ListIndex = 0
            End If
        Else
            cboGradientHeaderStyle.Enabled = False
            cboGradientHeaderStyle.ListIndex = -1
            Label19.Enabled = False
            If jcFrames.Style = jcGradient Then
                CboFillColor.Enabled = True
                CboFillColor.ListIndex = 7
                Label9.Enabled = True
                CboFrameColor.Enabled = True
                CboFrameColor.ListIndex = 0
                Label2.Enabled = True
            End If
            If jcFrames.Style = Panel Then
                CboFillColor.Enabled = True
                CboFillColor.ListIndex = 5
                Label9.Enabled = True
                CboThemeColor.Enabled = False
                CboThemeColor.ListIndex = -1
                Label1.Enabled = False
                CboFrameColor.Enabled = True
                CboFrameColor.ListIndex = 3
                Label2.Enabled = True
            End If
            If jcFrames.Style = Windows Or jcFrames.Style = TextBox Then
                CboTextBoxColor.Enabled = True
                CboTextBoxColor.ListIndex = 4
                Label3.Enabled = True
                CboThemeColor.Enabled = False
                CboThemeColor.ListIndex = -1
                Label1.Enabled = False
                If jcFrames.Style = TextBox Then CboRoundTxtBox.ListIndex = 1
            End If
            If jcFrames.Style = Messenger Then
                CboThemeColor.Enabled = True
                CboThemeColor.ListIndex = 0
                Label1.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub CboIconAlign_Click()
    If CboIconAlign.ListIndex = -1 Then Exit Sub
    If blnFormLoaded = True Then jcFrames.IconAlignment = CboIconAlign.ListIndex
End Sub

Private Sub CboIconSize_Click()
    If CboIconSize.ListIndex = -1 Then Exit Sub
    jcFrames.IconSize = Val(CboIconSize.Text)
End Sub

Private Sub CboPicture_Click()
    Select Case CboPicture.Text
        Case "Yes"
            Set jcFrames.Picture = LoadPicture(App.Path & "\103_56.ico")
            CboIconSize.Enabled = True
            CboIconAlign.Enabled = True
            Label16.Enabled = True
            Label13.Enabled = True
            CboIconSize.ListIndex = 0
            CboIconAlign.ListIndex = 0
        Case "No"
            Set jcFrames.Picture = Nothing
            CboIconSize.Enabled = False
            CboIconAlign.Enabled = False
            Label16.Enabled = False
            Label13.Enabled = False
            CboIconSize.ListIndex = -1
            CboIconAlign.ListIndex = -1
    End Select
End Sub

Private Sub CboRoundCorner_Click()
    If blnFormLoaded = True Then
        If CboRoundCorner.ListIndex = -1 Then Exit Sub
        jcFrames.RoundedCorner = CboRoundCorner.ListIndex
    End If
End Sub

Private Sub CboRoundTxtBox_Click()
    If blnFormLoaded = True Then
        If CboRoundTxtBox.ListIndex = -1 Then Exit Sub
        If jcFrames.HeaderStyle = Gradient And (jcFrames.Style = Windows Or jcFrames.Style = TextBox) And CboRoundTxtBox.ListIndex = 1 Then
            CboRoundTxtBox.ListIndex = 0
            Exit Sub
        End If
        jcFrames.RoundedCornerTxtBox = CboRoundTxtBox.ListIndex
    End If
End Sub

Private Sub CboTextBoxColor_Click()
    If blnFormLoaded = True Then
        If CboTextBoxColor.ListIndex = -1 Then Exit Sub
        jcFrames.TextboxColor = CboTextBoxColor.ItemData(CboTextBoxColor.ListIndex)
    End If
End Sub

Private Sub CboTextBoxHeight_Click()
    If blnFormLoaded = True Then
        If CboTextBoxHeight.ListIndex = -1 Then Exit Sub
        jcFrames.TextBoxHeight = Val(CboTextBoxHeight.Text)
    End If
End Sub

Private Sub CboTextColor_Click()
    If blnFormLoaded = True Then
        If CboTextColor.ListIndex = -1 Then Exit Sub
        jcFrames.TextColor = CboTextColor.ItemData(CboTextColor.ListIndex)
    End If
End Sub

Private Sub CboThemeColor_Click()
    If blnFormLoaded = True Then
        If CboThemeColor.ListIndex = -1 Then
            If Not ((jcFrames.Style = jcGradient Or jcFrames.Style = Messenger) And jcFrames.HeaderStyle = TxtBoxColor) Then
                CboThemeColor.ListIndex = 0
                Exit Sub
            End If
        End If
        If CboThemeColor.ListIndex = 5 Then
            CboColorFrom.Enabled = True
            CboColorTo.Enabled = True
            CboColorFrom.ListIndex = 4
            CboColorTo.ListIndex = 4
            Label14.Enabled = True
            Label15.Enabled = True
            If Not ((jcFrames.Style = jcGradient Or jcFrames.Style = Messenger) And jcFrames.HeaderStyle = TxtBoxColor) Then
                Label2.Enabled = True
                CboFrameColor.Enabled = True
                CboFrameColor.ListIndex = 0
            End If
        Else
            CboColorFrom.Enabled = False
            CboColorTo.Enabled = False
            CboColorFrom.ListIndex = -1
            CboColorTo.ListIndex = -1
            Label14.Enabled = False
            Label15.Enabled = False
            If Not ((jcFrames.Style = jcGradient Or jcFrames.Style = Messenger) And jcFrames.HeaderStyle = TxtBoxColor) Then
                Label2.Enabled = False
                CboFrameColor.Enabled = False
                CboFrameColor.ListIndex = -1
            End If
        End If
        If jcFrames.Style = Windows Or jcFrames.Style = TextBox Then
            Label2.Enabled = True
            CboFrameColor.Enabled = True
            CboFrameColor.ListIndex = 5
        End If
        jcFrames.ThemeColor = CboThemeColor.ListIndex
    End If
End Sub

Private Sub CboTxtBoxShadow_Click()
    If blnFormLoaded = True Then
        If CboTxtBoxShadow.ListIndex = -1 Then Exit Sub
        jcFrames.TxtBoxShadow = CboTxtBoxShadow.ListIndex
    End If
End Sub

Private Sub Check1_Click()
    Me.jcFrames.Enabled = Check1.Value
End Sub

Private Sub Form_Load()
    blnFormLoaded = True
    jcFrames.BackColor = Me.BackColor
    jcFrames4.ColorFrom = Me.BackColor
    jcFrames4.ColorTO = BlendColors(Me.BackColor, &HABABAB)
    jcFrames4.FrameColor = Me.BackColor
    OptCaptionStyle_Click 1
    'jcFrames1.Enabled = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Unselect_Control
End Sub

Private Sub jcFrames_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Unselect_Control
End Sub

Private Sub jcFrames7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Unselect_Control
End Sub

Private Sub lblBtn_Click()
    Unload Me
End Sub

Private Sub lblBtn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShpFocus ShpBtn, &H4080&, 0
End Sub

Private Sub OptCaptionStyle_Click(Index As Integer)
    Dim i As Integer
    
    If blnFormLoaded = False Then Exit Sub
    CboThemeColor.ListIndex = -1
    jcFrames.Style = Index
    jcFrames.Caption = OptCaptionStyle(Index).Tag
    jcFrames3.Caption = "Main enabled properties for " & OptCaptionStyle(Index).Tag
    txtCaption.Text = jcFrames.Caption
    Label21.Visible = True
    
    Select Case Index
        Case 0  'xp default
            CboCaptionAlig.Enabled = True:     CboCaptionAlig.ListIndex = 0:     Label5.Enabled = True
            CboRoundCorner.Enabled = True:     CboRoundCorner.ListIndex = 1:     Label6.Enabled = True
            CboTextColor.Enabled = True:       CboTextColor.ListIndex = 1:       Label7.Enabled = True
            CboTextBoxColor.Enabled = False:   CboTextBoxColor.ListIndex = -1:   Label3.Enabled = False
            CboTxtBoxShadow.Enabled = False:   CboTxtBoxShadow.ListIndex = -1:   Label3.Enabled = False
            CboTextBoxHeight.Enabled = False:  CboTextBoxHeight.ListIndex = -1:  Label4.Enabled = False
            CboPicture.Enabled = False:        CboPicture.ListIndex = 1:         Label12.Enabled = False
            CboIconSize.Enabled = False:       CboIconSize.ListIndex = -1:       Label13.Enabled = False
            CboIconAlign.Enabled = False:      CboIconAlign.ListIndex = -1:      Label16.Enabled = False
            CboFillColor.Enabled = False:      CboFillColor.ListIndex = -1:      Label9.Enabled = False
            CboRoundTxtBox.Enabled = False:    CboRoundTxtBox.ListIndex = -1:    Label8.Enabled = False
            CboThemeColor.Enabled = False:     CboThemeColor.ListIndex = -1:     Label1.Enabled = False
            CboColorFrom.Enabled = False:      CboColorFrom.ListIndex = -1:      Label14.Enabled = False
            CboColorTo.Enabled = False:        CboColorTo.ListIndex = -1:        Label15.Enabled = False
            CboFrameColor.Enabled = True:      CboFrameColor.ListIndex = 3:      Label2.Enabled = True
            CboHeaderStyle.Enabled = False:    CboHeaderStyle.ListIndex = -1:    Label18.Enabled = False
            cboGradientHeaderStyle.Enabled = False:  cboGradientHeaderStyle.ListIndex = -1:  Label19.Enabled = False
        Case 1  'jcGradient
            CboCaptionAlig.Enabled = True:     CboCaptionAlig.ListIndex = 2:     Label5.Enabled = True
            CboRoundCorner.Enabled = True:     CboRoundCorner.ListIndex = 1:     Label6.Enabled = True
            CboTextColor.Enabled = True:       CboTextColor.ListIndex = 0:       Label7.Enabled = True
            CboTextBoxColor.Enabled = False:   CboTextBoxColor.ListIndex = -1:   Label3.Enabled = False
            CboTxtBoxShadow.Enabled = False:   CboTxtBoxShadow.ListIndex = -1:   Label20.Enabled = False
            CboTextBoxHeight.Enabled = True:   CboTextBoxHeight.ListIndex = 2:   Label4.Enabled = True
            CboPicture.Enabled = True:         CboPicture.ListIndex = 1:         Label12.Enabled = True
            CboIconSize.Enabled = False:       CboIconSize.ListIndex = -1:       Label13.Enabled = False
            CboIconAlign.Enabled = False:      CboIconAlign.ListIndex = -1:      Label16.Enabled = False
            CboFillColor.Enabled = False:      CboFillColor.ListIndex = -1:      Label9.Enabled = False
            CboRoundTxtBox.Enabled = False:    CboRoundTxtBox.ListIndex = -1:    Label8.Enabled = False
            CboThemeColor.Enabled = True:      CboThemeColor.ListIndex = 2:      Label1.Enabled = True
            CboColorFrom.Enabled = False:      CboColorFrom.ListIndex = -1:      Label14.Enabled = False
            CboColorTo.Enabled = False:        CboColorTo.ListIndex = -1:        Label15.Enabled = False
            CboFrameColor.Enabled = False:     CboFrameColor.ListIndex = -1:     Label2.Enabled = False
            CboHeaderStyle.Enabled = True:     CboHeaderStyle.ListIndex = 1:     Label18.Enabled = True
            cboGradientHeaderStyle.Enabled = True:   cboGradientHeaderStyle.ListIndex = 0:   Label19.Enabled = True
        Case 2  'textbox
            CboCaptionAlig.Enabled = True:     CboCaptionAlig.ListIndex = 2:     Label5.Enabled = True
            CboRoundCorner.Enabled = True:     CboRoundCorner.ListIndex = 1:     Label6.Enabled = True
            CboTextColor.Enabled = True:       CboTextColor.ListIndex = 0:       Label7.Enabled = True
            CboTextBoxColor.Enabled = True:    CboTextBoxColor.ListIndex = 4:    Label3.Enabled = True
            CboTxtBoxShadow.Enabled = True:    CboTxtBoxShadow.ListIndex = 0:    Label20.Enabled = True
            CboTextBoxHeight.Enabled = True:   CboTextBoxHeight.ListIndex = 2:   Label4.Enabled = True
            CboPicture.Enabled = True:         CboPicture.ListIndex = 1:         Label12.Enabled = True
            CboIconSize.Enabled = False:       CboIconSize.ListIndex = -1:       Label13.Enabled = False
            CboIconAlign.Enabled = False:      CboIconAlign.ListIndex = -1:      Label16.Enabled = False
            CboFillColor.Enabled = True:       CboFillColor.ListIndex = 4:       Label9.Enabled = True
            CboRoundTxtBox.Enabled = True:     CboRoundTxtBox.ListIndex = 1:     Label8.Enabled = True
            CboThemeColor.Enabled = False:     CboThemeColor.ListIndex = -1:     Label1.Enabled = False
            CboColorFrom.Enabled = False:      CboColorFrom.ListIndex = -1:      Label14.Enabled = False
            CboColorTo.Enabled = False:        CboColorTo.ListIndex = -1:        Label15.Enabled = False
            CboFrameColor.Enabled = True:      CboFrameColor.ListIndex = 5:      Label2.Enabled = True
            CboHeaderStyle.Enabled = True:     CboHeaderStyle.ListIndex = 0:     Label18.Enabled = True
            cboGradientHeaderStyle.Enabled = False:  cboGradientHeaderStyle.ListIndex = -1:  Label19.Enabled = False
        Case 3  'windows
            CboCaptionAlig.Enabled = True:     CboCaptionAlig.ListIndex = 2:     Label5.Enabled = True
            CboRoundCorner.Enabled = True:     CboRoundCorner.ListIndex = 1:     Label6.Enabled = True
            CboTextColor.Enabled = True:       CboTextColor.ListIndex = 0:       Label7.Enabled = True
            CboTextBoxColor.Enabled = True:    CboTextBoxColor.ListIndex = 4:    Label3.Enabled = True
            CboTxtBoxShadow.Enabled = False:   CboTxtBoxShadow.ListIndex = -1:   Label20.Enabled = False
            CboTextBoxHeight.Enabled = True:   CboTextBoxHeight.ListIndex = 2:   Label4.Enabled = True
            CboPicture.Enabled = True:         CboPicture.ListIndex = 1:         Label12.Enabled = True
            CboIconSize.Enabled = False:       CboIconSize.ListIndex = -1:       Label13.Enabled = False
            CboIconAlign.Enabled = False:      CboIconAlign.ListIndex = -1:      Label16.Enabled = False
            CboFillColor.Enabled = True:       CboFillColor.ListIndex = 5:       Label9.Enabled = True
            CboRoundTxtBox.Enabled = True:     CboRoundTxtBox.ListIndex = 0:     Label8.Enabled = True
            CboThemeColor.Enabled = False:     CboThemeColor.ListIndex = -1:     Label1.Enabled = False
            CboColorFrom.Enabled = False:      CboColorFrom.ListIndex = -1:      Label14.Enabled = False
            CboColorTo.Enabled = False:        CboColorTo.ListIndex = -1:        Label15.Enabled = False
            CboFrameColor.Enabled = True:      CboFrameColor.ListIndex = 0:      Label2.Enabled = True
            CboHeaderStyle.Enabled = True:     CboHeaderStyle.ListIndex = 0:     Label18.Enabled = True
            cboGradientHeaderStyle.Enabled = False:  cboGradientHeaderStyle.ListIndex = -1:  Label19.Enabled = False
        Case 4  'messenger
            CboCaptionAlig.Enabled = True:     CboCaptionAlig.ListIndex = 2:     Label5.Enabled = True
            CboRoundCorner.Enabled = True:     CboRoundCorner.ListIndex = 0:     Label6.Enabled = True
            CboTextColor.Enabled = True:       CboTextColor.ListIndex = 0:       Label7.Enabled = True
            CboTextBoxColor.Enabled = False:   CboTextBoxColor.ListIndex = -1:   Label3.Enabled = False
            CboTxtBoxShadow.Enabled = False:   CboTxtBoxShadow.ListIndex = -1:   Label20.Enabled = False
            CboTextBoxHeight.Enabled = True:   CboTextBoxHeight.ListIndex = 2:   Label4.Enabled = True
            CboPicture.Enabled = True:         CboPicture.ListIndex = 1:         Label12.Enabled = True
            CboIconSize.Enabled = False:       CboIconSize.ListIndex = -1:       Label13.Enabled = False
            CboIconAlign.Enabled = False:      CboIconAlign.ListIndex = -1:      Label16.Enabled = False
            CboFillColor.Enabled = False:      CboFillColor.ListIndex = -1:      Label9.Enabled = False
            CboRoundTxtBox.Enabled = False:    CboRoundTxtBox.ListIndex = -1:    Label8.Enabled = False
            CboColorFrom.Enabled = False:      CboColorFrom.ListIndex = -1:      Label14.Enabled = False
            CboColorTo.Enabled = False:        CboColorTo.ListIndex = -1:        Label15.Enabled = False
            CboFrameColor.Enabled = False:     CboFrameColor.ListIndex = -1:     Label2.Enabled = False
            CboThemeColor.Enabled = True:      CboThemeColor.ListIndex = 0:      Label1.Enabled = True
            CboHeaderStyle.Enabled = True:     CboHeaderStyle.ListIndex = 0:     Label18.Enabled = True
            cboGradientHeaderStyle.Enabled = False:  cboGradientHeaderStyle.ListIndex = -1:  Label19.Enabled = False
        Case 5 'inner wedge
            CboCaptionAlig.Enabled = True:     CboCaptionAlig.ListIndex = 0:     Label5.Enabled = True
            CboRoundCorner.Enabled = True:     CboRoundCorner.ListIndex = 1:     Label6.Enabled = True
            CboTextColor.Enabled = True:       CboTextColor.ListIndex = 4:       Label7.Enabled = True
            CboTextBoxColor.Enabled = True:    CboTextBoxColor.ListIndex = 7:    Label3.Enabled = True
            CboTxtBoxShadow.Enabled = False:   CboTxtBoxShadow.ListIndex = -1:   Label20.Enabled = False
            CboTextBoxHeight.Enabled = False:  CboTextBoxHeight.ListIndex = -1:  Label4.Enabled = False
            CboPicture.Enabled = False:        CboPicture.ListIndex = 1:         Label12.Enabled = False
            CboIconSize.Enabled = False:       CboIconSize.ListIndex = -1:       Label13.Enabled = False
            CboIconAlign.Enabled = False:      CboIconAlign.ListIndex = -1:      Label16.Enabled = False
            CboFillColor.Enabled = True:       CboFillColor.ListIndex = 4:       Label9.Enabled = True
            CboRoundTxtBox.Enabled = False:    CboRoundTxtBox.ListIndex = -1:    Label8.Enabled = False
            CboThemeColor.Enabled = False:     CboThemeColor.ListIndex = -1:     Label1.Enabled = False
            CboColorFrom.Enabled = False:      CboColorFrom.ListIndex = -1:      Label14.Enabled = False
            CboColorTo.Enabled = False:        CboColorTo.ListIndex = -1:        Label15.Enabled = False
            CboFrameColor.Enabled = True:      CboFrameColor.ListIndex = 4:      Label2.Enabled = True
            CboHeaderStyle.Enabled = False:    CboHeaderStyle.ListIndex = -1:    Label18.Enabled = False
            cboGradientHeaderStyle.Enabled = False:  cboGradientHeaderStyle.ListIndex = -1:  Label19.Enabled = False
        Case 6  'outer wedge
            CboCaptionAlig.Enabled = True:     CboCaptionAlig.ListIndex = 0:     Label5.Enabled = True
            CboRoundCorner.Enabled = True:     CboRoundCorner.ListIndex = 1:     Label6.Enabled = True
            CboTextColor.Enabled = True:       CboTextColor.ListIndex = 4:       Label7.Enabled = True
            CboTextBoxColor.Enabled = True:    CboTextBoxColor.ListIndex = 6:    Label3.Enabled = True
            CboTxtBoxShadow.Enabled = False:   CboTxtBoxShadow.ListIndex = -1:   Label20.Enabled = False
            CboTextBoxHeight.Enabled = False:  CboTextBoxHeight.ListIndex = -1:  Label4.Enabled = False
            CboPicture.Enabled = False:        CboPicture.ListIndex = 1:         Label12.Enabled = False
            CboIconSize.Enabled = False:       CboIconSize.ListIndex = -1:       Label13.Enabled = False
            CboIconAlign.Enabled = False:      CboIconAlign.ListIndex = -1:      Label16.Enabled = False
            CboFillColor.Enabled = True:       CboFillColor.ListIndex = 4:       Label9.Enabled = True
            CboRoundTxtBox.Enabled = False:    CboRoundTxtBox.ListIndex = -1:    Label8.Enabled = False
            CboThemeColor.Enabled = False:     CboThemeColor.ListIndex = -1:     Label1.Enabled = False
            CboColorFrom.Enabled = False:      CboColorFrom.ListIndex = -1:      Label14.Enabled = False
            CboColorTo.Enabled = False:        CboColorTo.ListIndex = -1:        Label15.Enabled = False
            CboFrameColor.Enabled = True:      CboFrameColor.ListIndex = 1:      Label2.Enabled = True
            CboHeaderStyle.Enabled = False:    CboHeaderStyle.ListIndex = -1:    Label18.Enabled = False
            cboGradientHeaderStyle.Enabled = False:  cboGradientHeaderStyle.ListIndex = -1:  Label19.Enabled = False
        Case 7  'header
            CboCaptionAlig.Enabled = True:     CboCaptionAlig.ListIndex = 0:     Label5.Enabled = True
            CboRoundCorner.Enabled = False:    CboRoundCorner.ListIndex = -1:    Label6.Enabled = False
            CboTextColor.Enabled = True:       CboTextColor.ListIndex = 1:       Label7.Enabled = True
            CboTextBoxColor.Enabled = False:   CboTextBoxColor.ListIndex = -1:   Label3.Enabled = False
            CboTxtBoxShadow.Enabled = False:   CboTxtBoxShadow.ListIndex = -1:   Label20.Enabled = False
            CboTextBoxHeight.Enabled = False:  CboTextBoxHeight.ListIndex = -1:  Label4.Enabled = False
            CboPicture.Enabled = False:        CboPicture.ListIndex = 1:         Label12.Enabled = False
            CboIconSize.Enabled = False:       CboIconSize.ListIndex = -1:       Label13.Enabled = False
            CboIconAlign.Enabled = False:      CboIconAlign.ListIndex = -1:      Label16.Enabled = False
            CboFillColor.Enabled = False:      CboFillColor.ListIndex = -1:      Label9.Enabled = False
            CboRoundTxtBox.Enabled = False:    CboRoundTxtBox.ListIndex = -1:    Label8.Enabled = False
            CboThemeColor.Enabled = False:     CboThemeColor.ListIndex = -1:     Label1.Enabled = False
            CboColorFrom.Enabled = False:      CboColorFrom.ListIndex = -1:      Label14.Enabled = False
            CboColorTo.Enabled = False:        CboColorTo.ListIndex = -1:        Label15.Enabled = False
            CboFrameColor.Enabled = False:     CboFrameColor.ListIndex = -1:     Label2.Enabled = False
            CboHeaderStyle.Enabled = False:    CboHeaderStyle.ListIndex = -1:    Label18.Enabled = False
            cboGradientHeaderStyle.Enabled = False:  cboGradientHeaderStyle.ListIndex = -1:  Label19.Enabled = False
        Case 8 'panel
            Label21.Visible = False
            CboCaptionAlig.Enabled = True:     CboCaptionAlig.ListIndex = 2:     Label5.Enabled = True
            CboRoundCorner.Enabled = True:     CboRoundCorner.ListIndex = 1:     Label6.Enabled = True
            CboTextColor.Enabled = False:      CboTextColor.ListIndex = -1:      Label7.Enabled = False
            CboTextBoxColor.Enabled = False:   CboTextBoxColor.ListIndex = -1:   Label3.Enabled = False
            CboTxtBoxShadow.Enabled = False:   CboTxtBoxShadow.ListIndex = -1:   Label20.Enabled = False
            CboTextBoxHeight.Enabled = False:  CboTextBoxHeight.ListIndex = -1:  Label4.Enabled = False
            CboPicture.Enabled = True:         CboPicture.ListIndex = 1:         Label12.Enabled = True
            CboIconSize.Enabled = False:       CboIconSize.ListIndex = -1:       Label13.Enabled = False
            CboIconAlign.Enabled = False:      CboIconAlign.ListIndex = -1:      Label16.Enabled = False
            CboFillColor.Enabled = False:      CboFillColor.ListIndex = -1:      Label9.Enabled = False
            CboRoundTxtBox.Enabled = False:    CboRoundTxtBox.ListIndex = -1:    Label8.Enabled = False
            CboThemeColor.Enabled = True:      CboThemeColor.ListIndex = 1:      Label1.Enabled = True
            CboColorFrom.Enabled = False:      CboColorFrom.ListIndex = -1:      Label14.Enabled = False
            CboColorTo.Enabled = False:        CboColorTo.ListIndex = -1:        Label15.Enabled = False
            CboFrameColor.Enabled = True:      CboFrameColor.ListIndex = 0:      Label2.Enabled = True
            CboHeaderStyle.Enabled = True:     CboHeaderStyle.ListIndex = 1:     Label18.Enabled = True
            cboGradientHeaderStyle.Enabled = True:   cboGradientHeaderStyle.ListIndex = 3:   Label19.Enabled = True
    End Select
End Sub

Private Sub txtCaption_Change()
    jcFrames.Caption = txtCaption.Text
End Sub

Private Sub ShpFocus(shp As Shape, lngColor As Long, lngStyle As Long)
    shp.FillStyle = lngStyle
    shp.BorderColor = lngColor
End Sub

Private Sub Unselect_Control()
    ShpFocus ShpBtn, &H808080, 1
End Sub

'Blend two colors
Private Function BlendColors(ByVal lcolor1 As Long, ByVal lcolor2 As Long)
    BlendColors = RGB(((lcolor1 And &HFF) + (lcolor2 And &HFF)) / 2, (((lcolor1 \ &H100) And &HFF) + ((lcolor2 \ &H100) And &HFF)) / 2, (((lcolor1 \ &H10000) And &HFF) + ((lcolor2 \ &H10000) And &HFF)) / 2)
End Function

