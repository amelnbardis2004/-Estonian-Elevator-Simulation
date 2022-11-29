VERSION 5.00
Begin VB.Form keypad_floor 
   BackColor       =   &H8000000C&
   Caption         =   "Form1"
   ClientHeight    =   7800
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13515
   LinkTopic       =   "Form1"
   ScaleHeight     =   520
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   901
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox log 
      BackColor       =   &H8000000F&
      Height          =   7080
      Left            =   9840
      TabIndex        =   51
      Top             =   360
      Width           =   3135
   End
   Begin VB.CommandButton up 
      BackColor       =   &H8000000C&
      Caption         =   "UP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   10
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton down 
      BackColor       =   &H8000000C&
      Caption         =   "DOWN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   6840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton keypad_floor 
      BackColor       =   &H80000011&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   10
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton keypad_floor 
      BackColor       =   &H80000011&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   9
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton keypad_floor 
      BackColor       =   &H80000011&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   8
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton keypad_floor 
      BackColor       =   &H80000011&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   7
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton keypad_floor 
      BackColor       =   &H80000011&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton keypad_floor 
      BackColor       =   &H80000011&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton keypad_floor 
      BackColor       =   &H80000011&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton keypad_floor 
      BackColor       =   &H80000011&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton keypad_floor 
      BackColor       =   &H80000011&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton down 
      BackColor       =   &H8000000C&
      Caption         =   "DOWN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   10
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton elevator_door_right 
      Enabled         =   0   'False
      Height          =   615
      Left            =   8520
      TabIndex        =   22
      Top             =   6840
      Width           =   375
   End
   Begin VB.CommandButton elevator_door_left 
      Enabled         =   0   'False
      Height          =   615
      Left            =   8160
      TabIndex        =   21
      Top             =   6840
      Width           =   375
   End
   Begin VB.Timer timer 
      Interval        =   1000
      Left            =   840
      Top             =   6600
   End
   Begin VB.CommandButton down 
      BackColor       =   &H8000000C&
      Caption         =   "DOWN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton down 
      BackColor       =   &H8000000C&
      Caption         =   "DOWN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton up 
      BackColor       =   &H8000000C&
      Caption         =   "UP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton down 
      BackColor       =   &H8000000C&
      Caption         =   "DOWN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton up 
      BackColor       =   &H8000000C&
      Caption         =   "UP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton down 
      BackColor       =   &H8000000C&
      Caption         =   "DOWN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton up 
      BackColor       =   &H8000000C&
      Caption         =   "UP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton down 
      BackColor       =   &H8000000C&
      Caption         =   "DOWN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton up 
      BackColor       =   &H8000000C&
      Caption         =   "UP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton down 
      BackColor       =   &H8000000C&
      Caption         =   "DOWN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton up 
      BackColor       =   &H8000000C&
      Caption         =   "UP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton down 
      BackColor       =   &H8000000C&
      Caption         =   "DOWN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton up 
      BackColor       =   &H8000000C&
      Caption         =   "UP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton down 
      BackColor       =   &H8000000C&
      Caption         =   "DOWN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton up 
      BackColor       =   &H8000000C&
      Caption         =   "UP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton up 
      BackColor       =   &H8000000C&
      Caption         =   "UP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton up 
      BackColor       =   &H8000000C&
      Caption         =   "UP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton keypad_stop 
      BackColor       =   &H80000010&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton keypad_call 
      BackColor       =   &H80000010&
      Caption         =   "("
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   24
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5520
      Width           =   735
   End
   Begin VB.CommandButton keypad_door_hold 
      BackColor       =   &H80000010&
      Caption         =   "<|>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5520
      Width           =   735
   End
   Begin VB.CommandButton keypad_floor 
      BackColor       =   &H80000011&
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5520
      Width           =   735
   End
   Begin VB.Label elevator_back 
      BackColor       =   &H80000011&
      Height          =   615
      Left            =   8160
      TabIndex        =   50
      Top             =   6840
      Width           =   735
   End
   Begin VB.Line Line21 
      X1              =   872
      X2              =   872
      Y1              =   456
      Y2              =   504
   End
   Begin VB.Line Line20 
      X1              =   648
      X2              =   648
      Y1              =   456
      Y2              =   504
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000C&
      Caption         =   "* * * *"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   37
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000C&
      Caption         =   "* * * * *"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   36
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000C&
      Caption         =   "* * * *"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   35
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label floor_indicator 
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   1320
      TabIndex        =   33
      Top             =   720
      Width           =   495
   End
   Begin VB.Label floor_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " 9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   4560
      TabIndex        =   32
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label floor_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " 8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   4560
      TabIndex        =   31
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label floor_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " 7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   4560
      TabIndex        =   30
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label floor_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " 6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   4560
      TabIndex        =   29
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label floor_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " 5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   4560
      TabIndex        =   28
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label floor_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   4560
      TabIndex        =   27
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label floor_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   4560
      TabIndex        =   26
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label floor_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   4560
      TabIndex        =   25
      Top             =   6120
      Width           =   615
   End
   Begin VB.Label floor_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " L"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   4560
      TabIndex        =   24
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label floor_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   10
      Left            =   4560
      TabIndex        =   23
      Top             =   360
      Width           =   615
   End
   Begin VB.Line Line17 
      X1              =   648
      X2              =   872
      Y1              =   504
      Y2              =   504
   End
   Begin VB.Line Line16 
      X1              =   872
      X2              =   872
      Y1              =   456
      Y2              =   16
   End
   Begin VB.Line Line15 
      X1              =   648
      X2              =   648
      Y1              =   456
      Y2              =   16
   End
   Begin VB.Line Line14 
      X1              =   648
      X2              =   872
      Y1              =   16
      Y2              =   16
   End
   Begin VB.Line Line13 
      X1              =   536
      X2              =   600
      Y1              =   16
      Y2              =   16
   End
   Begin VB.Line Line12 
      X1              =   600
      X2              =   600
      Y1              =   504
      Y2              =   16
   End
   Begin VB.Line Line11 
      X1              =   536
      X2              =   600
      Y1              =   504
      Y2              =   504
   End
   Begin VB.Line Line10 
      X1              =   536
      X2              =   536
      Y1              =   16
      Y2              =   504
   End
   Begin VB.Line Line9 
      X1              =   296
      X2              =   296
      Y1              =   504
      Y2              =   16
   End
   Begin VB.Line Line8 
      X1              =   296
      X2              =   352
      Y1              =   504
      Y2              =   504
   End
   Begin VB.Line Line7 
      X1              =   352
      X2              =   352
      Y1              =   504
      Y2              =   16
   End
   Begin VB.Line Line6 
      X1              =   296
      X2              =   352
      Y1              =   16
      Y2              =   16
   End
   Begin VB.Line Line5 
      X1              =   528
      X2              =   528
      Y1              =   504
      Y2              =   16
   End
   Begin VB.Line Line4 
      X1              =   528
      X2              =   360
      Y1              =   504
      Y2              =   504
   End
   Begin VB.Line Line3 
      X1              =   360
      X2              =   360
      Y1              =   504
      Y2              =   16
   End
   Begin VB.Line Line2 
      X1              =   360
      X2              =   528
      Y1              =   16
      Y2              =   16
   End
   Begin VB.Line Line1 
      Index           =   11
      X1              =   48
      X2              =   160
      Y1              =   40
      Y2              =   40
   End
   Begin VB.Line Line1 
      Index           =   10
      X1              =   224
      X2              =   176
      Y1              =   88
      Y2              =   88
   End
   Begin VB.Line Line1 
      Index           =   9
      X1              =   176
      X2              =   176
      Y1              =   88
      Y2              =   40
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   224
      X2              =   224
      Y1              =   40
      Y2              =   88
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   24
      X2              =   248
      Y1              =   16
      Y2              =   16
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   24
      X2              =   248
      Y1              =   504
      Y2              =   504
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   248
      X2              =   248
      Y1              =   504
      Y2              =   16
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   24
      X2              =   24
      Y1              =   504
      Y2              =   16
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   160
      X2              =   160
      Y1              =   40
      Y2              =   88
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   176
      X2              =   224
      Y1              =   40
      Y2              =   40
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   48
      X2              =   48
      Y1              =   88
      Y2              =   40
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   160
      X2              =   48
      Y1              =   88
      Y2              =   88
   End
   Begin VB.Label keypad_direction 
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   27.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2775
      TabIndex        =   34
      Top             =   675
      Width           =   480
   End
End
Attribute VB_Name = "keypad_floor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Dim idle
Dim door

Private Sub Form_Load()

Randomize timer

End Sub


Private Sub timer_Timer()

For X = 1 To 10

    If elevator_door_left.Top = floor_label(X).Top Then
        floor_indicator.Caption = floor_label(X).Caption
    End If
    
    If State = 0 Then
        If up(X).BackColor <> &HC0FFFF Or down(X).BackColor <> &HC0FFFF Then
        idle = idle + 1
        End If
        If idle = 6000 Then
            idle = 0
            If elevator_door_left.Top <> up(1).Top Then
                elevator_door_left.Top = elevator_door_left.Top + 12
                elevator_door_right.Top = elevator_door_right.Top + 12
                elevator_back.Top = elevator_back.Top + 12
            End If
        End If
       
        
        
        If up(X).BackColor = &HC0FFFF Or down(X).BackColor = &HC0FFFF Then
            State = 2
            
            If up(X).Top < elevator_door_left.Top Then 'Is the requested floor above the elevator?
                keypad_direction.Caption = "ñ"
                keypad_direction.ForeColor = &HC000&
                
                elevator_door_left.Top = elevator_door_left.Top - 12
                elevator_door_right.Top = elevator_door_right.Top - 12
                elevator_back.Top = elevator_back.Top - 12
                
            ElseIf up(X).Top > elevator_door_left.Top Then 'Is the requested floor below the elevator?
                keypad_direction.Caption = "ò"
                keypad_direction.ForeColor = &HC0&
                
                elevator_door_left.Top = elevator_door_left.Top + 12
                elevator_door_right.Top = elevator_door_right.Top + 12
                elevator_back.Top = elevator_back.Top + 12
            
            Else:
                elevator_door_left.Top = up(X).Top
                
                up(X).BackColor = &H8000000C
                down(X).BackColor = &H8000000C
                keypad_floor(X).BackColor = &H80000011
                keypad_direction.Caption = ""
                
                log.AddItem (Time$ + ": Elevator arrived on floor " + CStr(X) + ".")
                
                door = 12
                     
            End If
        End If
        
        
        Select Case door
    
            Case 12 To 9:
                elevator_door_left.Width = elevator_door_left.Width - 4
                elevator_door_right.Left = elevator_door_right.Left + 4
                elevator_door_right.Width = elevator_door_right.Width - 4
                
            Case 8 To 5:
            
            Case 4 To 1:
                elevator_door_left.Width = elevator_door_left.Width + 4
                elevator_door_right.Left = elevator_door_right.Left - 4
                elevator_door_right.Width = elevator_door_right.Width + 4
        
        End Select
        End If
    
Next

End Sub

Private Sub keypad_floor_Click(Index As Integer)

keypad_floor(Index).BackColor = &HC0FFFF

End Sub

Private Sub keypad_door_hold_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

keypad_door_hold.BackColor = &HC0FFFF

End Sub

Private Sub up_Click(Index As Integer)

up(Index).BackColor = &HC0FFFF
keypad_floor(Index).BackColor = &HC0FFFF

log.AddItem (Time$ + ": Floor " + CStr(Index) + " UP button pressed.")

End Sub

Private Sub down_Click(Index As Integer)

down(Index).BackColor = &HC0FFFF
keypad_floor(Index).BackColor = &HC0FFFF

log.AddItem (Time$ + ": Floor " + CStr(Index) + " DOWN button pressed.")

End Sub

Private Sub keypad_call_Click()

log.AddItem (Time$ + ": Emergency call placed.")

End Sub

Private Sub keypad_stop_Click()

timer.Enabled = False
log.AddItem (Time$ + ": Elevator shut down.")

End Sub
