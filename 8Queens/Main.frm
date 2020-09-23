VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "8Queens (By Salan Sinan)"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7260
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   402
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   484
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNew1 
      Caption         =   "&New"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.PictureBox Queen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      DragIcon        =   "Main.frx":08CA
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   240
      Picture         =   "Main.frx":1194
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5160
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   0
      Left            =   1320
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   120
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   1
      Left            =   2040
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   120
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   2
      Left            =   2760
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   120
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   3
      Left            =   3480
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   120
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   4
      Left            =   4200
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   120
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   5
      Left            =   4920
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   120
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   6
      Left            =   5640
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   120
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   7
      Left            =   6360
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   120
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   8
      Left            =   1320
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   840
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   9
      Left            =   2040
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   840
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   10
      Left            =   2760
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   840
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   11
      Left            =   3480
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   840
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   12
      Left            =   4200
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   840
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   13
      Left            =   4920
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   840
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   14
      Left            =   5640
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   840
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   15
      Left            =   6360
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   840
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   16
      Left            =   1320
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   1560
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   17
      Left            =   2040
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   1560
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   18
      Left            =   2760
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   1560
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   19
      Left            =   3480
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   1560
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   20
      Left            =   4200
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   1560
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   21
      Left            =   4920
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   1560
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   22
      Left            =   5640
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   1560
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   23
      Left            =   6360
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   1560
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   24
      Left            =   1320
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   2280
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   25
      Left            =   2040
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   2280
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   26
      Left            =   2760
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   2280
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   27
      Left            =   3480
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   2280
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   28
      Left            =   4200
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   2280
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   29
      Left            =   4920
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   2280
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   30
      Left            =   5640
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   2280
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   31
      Left            =   6360
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   2280
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   32
      Left            =   1320
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   3000
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   33
      Left            =   2040
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   3000
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   34
      Left            =   2760
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   3000
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   35
      Left            =   3480
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   3000
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   36
      Left            =   4200
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   3000
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   37
      Left            =   4920
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   3000
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   38
      Left            =   5640
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   3000
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   39
      Left            =   6360
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   3000
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   40
      Left            =   1320
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   3720
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   41
      Left            =   2040
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   3720
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   42
      Left            =   2760
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   3720
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   43
      Left            =   3480
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3720
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   44
      Left            =   4200
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3720
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   45
      Left            =   4920
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3720
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   46
      Left            =   5640
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3720
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   47
      Left            =   6360
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3720
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   48
      Left            =   1320
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4440
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   49
      Left            =   2040
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4440
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   50
      Left            =   2760
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4440
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   51
      Left            =   3480
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4440
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   52
      Left            =   4200
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4440
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   53
      Left            =   4920
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4440
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   54
      Left            =   5640
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4440
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   55
      Left            =   6360
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4440
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   56
      Left            =   1320
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5160
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   57
      Left            =   2040
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5160
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   58
      Left            =   2760
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5160
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   59
      Left            =   3480
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5160
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   60
      Left            =   4200
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5160
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   61
      Left            =   4920
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5160
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   62
      Left            =   5640
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5160
      Width           =   750
   End
   Begin VB.PictureBox Puz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   63
      Left            =   6360
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5160
      Width           =   750
   End
   Begin VB.Label lblQN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Queen Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Q(64) As Integer
Dim QN As Integer

Private Sub cmdHelp_Click()
    Dim str As String
    
    str = "The job is to place an 8 queens to a white boxes." & vbCrLf & _
    "Movements:" & vbCrLf & _
    "   You can drag (move) the queen to a white box," & vbCrLf & _
    "   or you can remove the queen by only double click on it." & vbCrLf & _
    "Constrains:" & vbCrLf & _
    "   Each queen you place, according to its position," & vbCrLf & _
    "   a horizontal, vertical, and two diagonals lines will be in red (can't be used)." & vbCrLf & vbCrLf & _
    "If you solve the problem, just send me a picture of your solution at:" & vbCrLf & _
    "salan@uruklink.net" & vbCrLf & vbCrLf & _
    "Thank you for playing 8Queens" & vbCrLf & _
    "Salan Sinan"
    
    MsgBox str, vbOKOnly
End Sub

Private Sub cmdNew1_Click()
    Dim i As Integer
    
    QN = 8
    
    For i = 0 To 63
        Q(i) = 0
    Next i
    
    Help
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    QN = 8
    
    For i = 0 To 63
        Q(i) = 0
    Next i
    
    Help
End Sub

Private Sub Puz_dblClick(Index As Integer)
    Dim i As Integer
    Dim row As Integer
    Dim col As Integer
    Dim irow As Integer
    Dim icol As Integer
    
    If Q(Index) <> 2 Then
        MsgBox "There's no queen here to remove", vbCritical
        Exit Sub
    End If
    
    row = Int(Index / 8)
    col = Index Mod 8
    
    For i = 0 To 63
        irow = Int(i / 8)
        icol = i Mod 8
        If irow = row Or icol = col Or irow - icol = row - col Or irow + icol = row + col Then
            Q(i) = 0
        End If
    Next i
    
    Check
    QN = QN + 1
    Help
End Sub

Private Sub Puz_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    If Q(Index) <> 0 Then
        MsgBox "Sorry, can't place the queen here", vbCritical
        Exit Sub
    End If
    
    Q(Index) = 2
    Check
    QN = QN - 1
    
    If QN = 0 Then
        Help
        MsgBox "You have done", vbExclamation
        Exit Sub
    End If

    Help
End Sub

Function Help()
    Dim i As Integer
    
    lblQN.Caption = CStr(QN)
    
    For i = 0 To 63
        Puz(i).Picture = LoadPicture()
        If Q(i) = 1 Then
            Puz(i).BackColor = QBColor(12)
        ElseIf Q(i) = 2 Then
            Puz(i).Picture = Queen.Picture
        Else
            Puz(i).BackColor = QBColor(15)
        End If
    Next i
End Function

Function Check()
    Dim i As Integer
    Dim j As Integer
    
    Dim row As Integer
    Dim col As Integer
    Dim irow As Integer
    Dim icol As Integer
    
    For j = 0 To 63
        If Q(j) = 2 Then
            row = Int(j / 8)
            col = j Mod 8
    
            For i = 0 To 63
                If i <> j Then
                    irow = Int(i / 8)
                    icol = i Mod 8
                    If irow = row Or icol = col Or irow - icol = row - col Or irow + icol = row + col Then
                        Q(i) = 1
                    End If
                End If
            Next i
        End If
    Next j
End Function
