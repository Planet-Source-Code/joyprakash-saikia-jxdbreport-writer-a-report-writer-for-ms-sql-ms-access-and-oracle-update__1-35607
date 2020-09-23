VERSION 5.00
Object = "{A97B8938-0414-11D5-83E3-008048D61E92}#2.0#0"; "SoftBttn.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   5685
   StartUpPosition =   3  'Windows Default
   Begin SoftButton.SoftBttn SoftBttn3 
      Height          =   345
      Left            =   3360
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   6660
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   609
      PictureAlignment=   4
      BackColor       =   -2147483633
      Enabled         =   -1  'True
      MaskColor       =   16777215
      UseMaskColor    =   0   'False
      PlaySounds      =   0   'False
      Object.ToolTipText     =   "Click To Exit Demo"
      Caption         =   "Exit"
      ShrinkIcon      =   0   'False
      IconHeight      =   16
      IconWidth       =   16
      TextAlignment   =   4
      ForeColor       =   0
      PopUpButton     =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      ForeColor       =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "Picture Alignments"
      Height          =   2970
      Index           =   1
      Left            =   195
      TabIndex        =   28
      Top             =   3450
      Width           =   5340
      Begin SoftButton.SoftBttn SoftBttn2 
         Height          =   885
         Index           =   9
         Left            =   150
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   210
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   1561
         Picture         =   "Form1.frx":0000
         PictureAlignment=   0
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   0   'False
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   ""
         ShrinkIcon      =   0   'False
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   0
         ForeColor       =   0
         PopUpButton     =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn SoftBttn2 
         Height          =   885
         Index           =   10
         Left            =   1830
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   210
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   1561
         Picture         =   "Form1.frx":27B2
         PictureAlignment=   1
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   0   'False
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   ""
         ShrinkIcon      =   0   'False
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   1
         ForeColor       =   0
         PopUpButton     =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn SoftBttn2 
         Height          =   885
         Index           =   11
         Left            =   3510
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   210
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   1561
         Picture         =   "Form1.frx":4F64
         PictureAlignment=   2
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   0   'False
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   ""
         ShrinkIcon      =   0   'False
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   2
         ForeColor       =   0
         PopUpButton     =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn SoftBttn2 
         Height          =   885
         Index           =   12
         Left            =   150
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1095
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   1561
         Picture         =   "Form1.frx":7716
         PictureAlignment=   3
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   0   'False
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   ""
         ShrinkIcon      =   0   'False
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   3
         ForeColor       =   0
         PopUpButton     =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn SoftBttn2 
         Height          =   885
         Index           =   13
         Left            =   1830
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1095
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   1561
         Picture         =   "Form1.frx":9EC8
         PictureAlignment=   4
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   0   'False
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   ""
         ShrinkIcon      =   0   'False
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   4
         ForeColor       =   0
         PopUpButton     =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn SoftBttn2 
         Height          =   885
         Index           =   14
         Left            =   3510
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1095
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   1561
         Picture         =   "Form1.frx":C67A
         PictureAlignment=   5
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   0   'False
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   ""
         ShrinkIcon      =   0   'False
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   5
         ForeColor       =   0
         PopUpButton     =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn SoftBttn2 
         Height          =   885
         Index           =   15
         Left            =   150
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1980
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   1561
         Picture         =   "Form1.frx":EE2C
         PictureAlignment=   6
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   0   'False
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   ""
         ShrinkIcon      =   0   'False
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   6
         ForeColor       =   0
         PopUpButton     =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn SoftBttn2 
         Height          =   885
         Index           =   16
         Left            =   1830
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1980
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   1561
         Picture         =   "Form1.frx":F27E
         PictureAlignment=   7
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   0   'False
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   ""
         ShrinkIcon      =   0   'False
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   7
         ForeColor       =   0
         PopUpButton     =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn SoftBttn2 
         Height          =   885
         Index           =   17
         Left            =   3510
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1980
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   1561
         Picture         =   "Form1.frx":11A30
         PictureAlignment=   8
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   0   'False
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   ""
         ShrinkIcon      =   0   'False
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   8
         ForeColor       =   0
         PopUpButton     =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Text Alignments"
      Height          =   2100
      Index           =   0
      Left            =   90
      TabIndex        =   18
      Top             =   1305
      Width           =   5550
      Begin SoftButton.SoftBttn SoftBttn2 
         Height          =   570
         Index           =   0
         Left            =   150
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   210
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   1005
         PictureAlignment=   4
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   0   'False
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   "Top Left"
         ShrinkIcon      =   0   'False
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   0
         ForeColor       =   0
         PopUpButton     =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn SoftBttn2 
         Height          =   570
         Index           =   1
         Left            =   1905
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   210
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   1005
         PictureAlignment=   4
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   0   'False
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   "Top Center"
         ShrinkIcon      =   0   'False
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   1
         ForeColor       =   0
         PopUpButton     =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn SoftBttn2 
         Height          =   570
         Index           =   2
         Left            =   3660
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   210
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   1005
         PictureAlignment=   4
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   0   'False
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   "Top Right"
         ShrinkIcon      =   0   'False
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   2
         ForeColor       =   0
         PopUpButton     =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn SoftBttn2 
         Height          =   570
         Index           =   3
         Left            =   150
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   810
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   1005
         PictureAlignment=   4
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   0   'False
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   "Left Center"
         ShrinkIcon      =   0   'False
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   3
         ForeColor       =   0
         PopUpButton     =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn SoftBttn2 
         Height          =   570
         Index           =   4
         Left            =   1905
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   810
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   1005
         PictureAlignment=   4
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   0   'False
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   "Center"
         ShrinkIcon      =   0   'False
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   4
         ForeColor       =   0
         PopUpButton     =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn SoftBttn2 
         Height          =   570
         Index           =   5
         Left            =   3660
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   810
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   1005
         PictureAlignment=   4
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   0   'False
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   "Right Center"
         ShrinkIcon      =   0   'False
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   5
         ForeColor       =   0
         PopUpButton     =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn SoftBttn2 
         Height          =   570
         Index           =   6
         Left            =   150
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1410
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   1005
         PictureAlignment=   4
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   0   'False
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   "Bottom Left"
         ShrinkIcon      =   0   'False
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   6
         ForeColor       =   0
         PopUpButton     =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn SoftBttn2 
         Height          =   570
         Index           =   7
         Left            =   1905
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1410
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   1005
         PictureAlignment=   4
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   0   'False
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   "Bottom Center"
         ShrinkIcon      =   0   'False
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   7
         ForeColor       =   0
         PopUpButton     =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn SoftBttn2 
         Height          =   570
         Index           =   8
         Left            =   3660
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1410
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   1005
         PictureAlignment=   4
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   0   'False
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   "Right Center"
         ShrinkIcon      =   0   'False
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   8
         ForeColor       =   0
         PopUpButton     =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   630
      Index           =   1
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   5685
      TabIndex        =   9
      Top             =   570
      Width           =   5685
      Begin SoftButton.SoftBttn SoftBttn1 
         Height          =   555
         Index           =   8
         Left            =   30
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   30
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Picture         =   "Form1.frx":141E2
         PictureAlignment=   4
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   -1  'True
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   ""
         ShrinkIcon      =   0   'False
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   7
         ForeColor       =   0
         PopUpButton     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn SoftBttn1 
         Height          =   555
         Index           =   9
         Left            =   675
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   30
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Picture         =   "Form1.frx":14634
         PictureAlignment=   4
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   -1  'True
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   ""
         ShrinkIcon      =   0   'False
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   7
         ForeColor       =   0
         PopUpButton     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn SoftBttn1 
         Height          =   555
         Index           =   10
         Left            =   1320
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   30
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Picture         =   "Form1.frx":1478E
         PictureAlignment=   4
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   -1  'True
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   ""
         ShrinkIcon      =   0   'False
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   7
         ForeColor       =   0
         PopUpButton     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn SoftBttn1 
         Height          =   555
         Index           =   11
         Left            =   1965
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   30
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Picture         =   "Form1.frx":16F40
         PictureAlignment=   4
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   -1  'True
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   ""
         ShrinkIcon      =   0   'False
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   7
         ForeColor       =   0
         PopUpButton     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn SoftBttn1 
         Height          =   555
         Index           =   12
         Left            =   2610
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   30
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Picture         =   "Form1.frx":1725A
         PictureAlignment=   4
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   -1  'True
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   ""
         ShrinkIcon      =   0   'False
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   7
         ForeColor       =   0
         PopUpButton     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn SoftBttn1 
         Height          =   555
         Index           =   13
         Left            =   3255
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   30
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Picture         =   "Form1.frx":19A0C
         PictureAlignment=   4
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   -1  'True
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   ""
         ShrinkIcon      =   0   'False
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   7
         ForeColor       =   0
         PopUpButton     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn SoftBttn1 
         Height          =   555
         Index           =   14
         Left            =   3900
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   30
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Picture         =   "Form1.frx":1C1BE
         PictureAlignment=   4
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   -1  'True
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   ""
         ShrinkIcon      =   0   'False
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   7
         ForeColor       =   0
         PopUpButton     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn SoftBttn1 
         Height          =   555
         Index           =   15
         Left            =   4545
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   30
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   979
         Picture         =   "Form1.frx":1E970
         PictureAlignment=   4
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   -1  'True
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   ""
         ShrinkIcon      =   0   'False
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   7
         ForeColor       =   0
         PopUpButton     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   0
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   5685
      TabIndex        =   0
      Top             =   0
      Width           =   5685
      Begin SoftButton.SoftBttn SoftBttn1 
         Height          =   480
         Index           =   0
         Left            =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   30
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   847
         Picture         =   "Form1.frx":21122
         PictureAlignment=   4
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   -1  'True
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   ""
         ShrinkIcon      =   -1  'True
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   7
         ForeColor       =   0
         PopUpButton     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn SoftBttn1 
         Height          =   480
         Index           =   1
         Left            =   585
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   30
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   847
         Picture         =   "Form1.frx":21574
         PictureAlignment=   4
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   -1  'True
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   ""
         ShrinkIcon      =   -1  'True
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   7
         ForeColor       =   0
         PopUpButton     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn SoftBttn1 
         Height          =   480
         Index           =   2
         Left            =   1140
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   30
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   847
         Picture         =   "Form1.frx":216CE
         PictureAlignment=   4
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   -1  'True
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   ""
         ShrinkIcon      =   -1  'True
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   7
         ForeColor       =   0
         PopUpButton     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn SoftBttn1 
         Height          =   480
         Index           =   3
         Left            =   1695
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   30
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   847
         Picture         =   "Form1.frx":23E80
         PictureAlignment=   4
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   -1  'True
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   ""
         ShrinkIcon      =   -1  'True
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   7
         ForeColor       =   0
         PopUpButton     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn SoftBttn1 
         Height          =   480
         Index           =   4
         Left            =   2250
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   30
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   847
         Picture         =   "Form1.frx":2419A
         PictureAlignment=   4
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   -1  'True
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   ""
         ShrinkIcon      =   -1  'True
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   7
         ForeColor       =   0
         PopUpButton     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn SoftBttn1 
         Height          =   480
         Index           =   5
         Left            =   2805
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   30
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   847
         Picture         =   "Form1.frx":2694C
         PictureAlignment=   4
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   -1  'True
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   ""
         ShrinkIcon      =   -1  'True
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   7
         ForeColor       =   0
         PopUpButton     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn SoftBttn1 
         Height          =   480
         Index           =   6
         Left            =   3360
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   30
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   847
         Picture         =   "Form1.frx":290FE
         PictureAlignment=   4
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   -1  'True
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   ""
         ShrinkIcon      =   -1  'True
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   7
         ForeColor       =   0
         PopUpButton     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn SoftBttn1 
         Height          =   480
         Index           =   7
         Left            =   3915
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   30
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   847
         Picture         =   "Form1.frx":2B8B0
         PictureAlignment=   4
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   -1  'True
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   ""
         ShrinkIcon      =   -1  'True
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   7
         ForeColor       =   0
         PopUpButton     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



'************************************************************************
'Author            :   Vijay Phulwadhawa     Date    : 23/02/2001 12:59:50 PM
'Project Name      :   Project1
'Form/Class Name   :   Form1 (Code)
'Version           :   6.00
'Description       :   <Purpose>
'Links             :   <Links With Any Other Form Modules>
'Change History    :
'Date      Author      Description Of Changes          Reason Of Change
'************************************************************************



Private Sub SoftBttn3_Click()
Unload Me
End Sub

Private Sub SoftBttn3_MouseEnter()
SoftBttn3.ForeColor = vbBlue
SoftBttn3.FontBold = True
End Sub

Private Sub SoftBttn3_MouseExit()
    SoftBttn3.ForeColor = vbBlack
    SoftBttn3.FontBold = False

End Sub
