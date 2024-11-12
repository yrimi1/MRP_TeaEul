VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRecipeCalcView 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   9255
   ClientLeft      =   360
   ClientTop       =   330
   ClientWidth     =   15270
   Icon            =   "frmRecipeCalcView.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   15270
   Begin VB.TextBox txtRPCalcRemark 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   7950
      MultiLine       =   -1  'True
      TabIndex        =   44
      Top             =   8610
      Visible         =   0   'False
      Width           =   855
   End
   Begin VSFlex7LCtl.VSFlexGrid grdCard 
      Height          =   585
      Left            =   6690
      TabIndex        =   43
      Top             =   8730
      Visible         =   0   'False
      Width           =   1125
      _cx             =   1984
      _cy             =   1032
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  '없음
      Height          =   630
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   18
      Text            =   "frmRecipeCalcView.frx":000C
      Top             =   8550
      Width           =   4785
   End
   Begin Crystal.CrystalReport cryReport 
      Left            =   3660
      Top             =   8865
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   3630
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   8493
      _Version        =   196609
      BackColor       =   12632256
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel pnlCalc 
         Height          =   4440
         Index           =   1
         Left            =   7755
         TabIndex        =   1
         Top             =   360
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   7832
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin TabDlg.SSTab tabDye 
            Height          =   2475
            Left            =   15
            TabIndex        =   2
            Top             =   30
            Width           =   7410
            _ExtentX        =   13070
            _ExtentY        =   4366
            _Version        =   393216
            TabHeight       =   520
            TabCaption(0)   =   "추가 1"
            TabPicture(0)   =   "frmRecipeCalcView.frx":0073
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "grdDye(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "추가 2"
            TabPicture(1)   =   "frmRecipeCalcView.frx":008F
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grdDye(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "추가 3"
            TabPicture(2)   =   "frmRecipeCalcView.frx":00AB
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grdDye(2)"
            Tab(2).ControlCount=   1
            Begin VSFlex7LCtl.VSFlexGrid grdDye 
               Height          =   2085
               Index           =   0
               Left            =   45
               TabIndex        =   3
               Top             =   360
               Width           =   7290
               _cx             =   12859
               _cy             =   3678
               _ConvInfo       =   1
               Appearance      =   1
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MousePointer    =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483632
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483642
               FocusRect       =   1
               HighLight       =   1
               AllowSelection  =   -1  'True
               AllowBigSelection=   -1  'True
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   10
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   ""
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   -1  'True
               AutoSizeMode    =   0
               AutoSearch      =   0
               AutoSearchDelay =   2
               MultiTotals     =   -1  'True
               SubtotalPosition=   1
               OutlineBar      =   0
               OutlineCol      =   0
               Ellipsis        =   0
               ExplorerBar     =   0
               PicturesOver    =   0   'False
               FillStyle       =   0
               RightToLeft     =   0   'False
               PictureType     =   0
               TabBehavior     =   0
               OwnerDraw       =   0
               Editable        =   0
               ShowComboButton =   -1  'True
               WordWrap        =   0   'False
               TextStyle       =   0
               TextStyleFixed  =   0
               OleDragMode     =   0
               OleDropMode     =   0
               ComboSearch     =   3
               AutoSizeMouse   =   -1  'True
               FrozenRows      =   0
               FrozenCols      =   0
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
            End
            Begin VSFlex7LCtl.VSFlexGrid grdDye 
               Height          =   2055
               Index           =   1
               Left            =   -74955
               TabIndex        =   4
               Top             =   360
               Width           =   7290
               _cx             =   12859
               _cy             =   3625
               _ConvInfo       =   1
               Appearance      =   1
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MousePointer    =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483632
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483642
               FocusRect       =   1
               HighLight       =   1
               AllowSelection  =   -1  'True
               AllowBigSelection=   -1  'True
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   10
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   ""
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   -1  'True
               AutoSizeMode    =   0
               AutoSearch      =   0
               AutoSearchDelay =   2
               MultiTotals     =   -1  'True
               SubtotalPosition=   1
               OutlineBar      =   0
               OutlineCol      =   0
               Ellipsis        =   0
               ExplorerBar     =   0
               PicturesOver    =   0   'False
               FillStyle       =   0
               RightToLeft     =   0   'False
               PictureType     =   0
               TabBehavior     =   0
               OwnerDraw       =   0
               Editable        =   0
               ShowComboButton =   -1  'True
               WordWrap        =   0   'False
               TextStyle       =   0
               TextStyleFixed  =   0
               OleDragMode     =   0
               OleDropMode     =   0
               ComboSearch     =   3
               AutoSizeMouse   =   -1  'True
               FrozenRows      =   0
               FrozenCols      =   0
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
            End
            Begin VSFlex7LCtl.VSFlexGrid grdDye 
               Height          =   2055
               Index           =   2
               Left            =   -74955
               TabIndex        =   5
               Top             =   360
               Width           =   7290
               _cx             =   12859
               _cy             =   3625
               _ConvInfo       =   1
               Appearance      =   1
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MousePointer    =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483632
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483642
               FocusRect       =   1
               HighLight       =   1
               AllowSelection  =   -1  'True
               AllowBigSelection=   -1  'True
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   10
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   ""
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   -1  'True
               AutoSizeMode    =   0
               AutoSearch      =   0
               AutoSearchDelay =   2
               MultiTotals     =   -1  'True
               SubtotalPosition=   1
               OutlineBar      =   0
               OutlineCol      =   0
               Ellipsis        =   0
               ExplorerBar     =   0
               PicturesOver    =   0   'False
               FillStyle       =   0
               RightToLeft     =   0   'False
               PictureType     =   0
               TabBehavior     =   0
               OwnerDraw       =   0
               Editable        =   0
               ShowComboButton =   -1  'True
               WordWrap        =   0   'False
               TextStyle       =   0
               TextStyleFixed  =   0
               OleDragMode     =   0
               OleDropMode     =   0
               ComboSearch     =   3
               AutoSizeMouse   =   -1  'True
               FrozenRows      =   0
               FrozenCols      =   0
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
            End
         End
         Begin TabDlg.SSTab tabAux 
            Height          =   1890
            Left            =   30
            TabIndex        =   6
            Top             =   2505
            Width           =   7410
            _ExtentX        =   13070
            _ExtentY        =   3334
            _Version        =   393216
            TabHeight       =   520
            TabCaption(0)   =   "추가 1"
            TabPicture(0)   =   "frmRecipeCalcView.frx":00C7
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "grdAux(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "추가 2"
            TabPicture(1)   =   "frmRecipeCalcView.frx":00E3
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grdAux(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "추가 3"
            TabPicture(2)   =   "frmRecipeCalcView.frx":00FF
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grdAux(2)"
            Tab(2).ControlCount=   1
            Begin VSFlex7LCtl.VSFlexGrid grdAux 
               Height          =   1455
               Index           =   0
               Left            =   45
               TabIndex        =   7
               Top             =   360
               Width           =   7290
               _cx             =   12859
               _cy             =   2566
               _ConvInfo       =   1
               Appearance      =   1
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MousePointer    =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483632
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483642
               FocusRect       =   1
               HighLight       =   1
               AllowSelection  =   -1  'True
               AllowBigSelection=   -1  'True
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   10
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   ""
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   -1  'True
               AutoSizeMode    =   0
               AutoSearch      =   0
               AutoSearchDelay =   2
               MultiTotals     =   -1  'True
               SubtotalPosition=   1
               OutlineBar      =   0
               OutlineCol      =   0
               Ellipsis        =   0
               ExplorerBar     =   0
               PicturesOver    =   0   'False
               FillStyle       =   0
               RightToLeft     =   0   'False
               PictureType     =   0
               TabBehavior     =   0
               OwnerDraw       =   0
               Editable        =   0
               ShowComboButton =   -1  'True
               WordWrap        =   0   'False
               TextStyle       =   0
               TextStyleFixed  =   0
               OleDragMode     =   0
               OleDropMode     =   0
               ComboSearch     =   3
               AutoSizeMouse   =   -1  'True
               FrozenRows      =   0
               FrozenCols      =   0
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
            End
            Begin VSFlex7LCtl.VSFlexGrid grdAux 
               Height          =   1470
               Index           =   1
               Left            =   -74955
               TabIndex        =   8
               Top             =   360
               Width           =   7290
               _cx             =   12859
               _cy             =   2593
               _ConvInfo       =   1
               Appearance      =   1
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MousePointer    =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483632
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483642
               FocusRect       =   1
               HighLight       =   1
               AllowSelection  =   -1  'True
               AllowBigSelection=   -1  'True
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   10
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   ""
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   -1  'True
               AutoSizeMode    =   0
               AutoSearch      =   0
               AutoSearchDelay =   2
               MultiTotals     =   -1  'True
               SubtotalPosition=   1
               OutlineBar      =   0
               OutlineCol      =   0
               Ellipsis        =   0
               ExplorerBar     =   0
               PicturesOver    =   0   'False
               FillStyle       =   0
               RightToLeft     =   0   'False
               PictureType     =   0
               TabBehavior     =   0
               OwnerDraw       =   0
               Editable        =   0
               ShowComboButton =   -1  'True
               WordWrap        =   0   'False
               TextStyle       =   0
               TextStyleFixed  =   0
               OleDragMode     =   0
               OleDropMode     =   0
               ComboSearch     =   3
               AutoSizeMouse   =   -1  'True
               FrozenRows      =   0
               FrozenCols      =   0
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
            End
            Begin VSFlex7LCtl.VSFlexGrid grdAux 
               Height          =   1470
               Index           =   2
               Left            =   -74955
               TabIndex        =   9
               Top             =   360
               Width           =   7290
               _cx             =   12859
               _cy             =   2593
               _ConvInfo       =   1
               Appearance      =   1
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MousePointer    =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483632
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483642
               FocusRect       =   1
               HighLight       =   1
               AllowSelection  =   -1  'True
               AllowBigSelection=   -1  'True
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   10
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   ""
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   -1  'True
               AutoSizeMode    =   0
               AutoSearch      =   0
               AutoSearchDelay =   2
               MultiTotals     =   -1  'True
               SubtotalPosition=   1
               OutlineBar      =   0
               OutlineCol      =   0
               Ellipsis        =   0
               ExplorerBar     =   0
               PicturesOver    =   0   'False
               FillStyle       =   0
               RightToLeft     =   0   'False
               PictureType     =   0
               TabBehavior     =   0
               OwnerDraw       =   0
               Editable        =   0
               ShowComboButton =   -1  'True
               WordWrap        =   0   'False
               TextStyle       =   0
               TextStyleFixed  =   0
               OleDragMode     =   0
               OleDropMode     =   0
               ComboSearch     =   3
               AutoSizeMouse   =   -1  'True
               FrozenRows      =   0
               FrozenCols      =   0
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
            End
         End
      End
      Begin Threed.SSPanel pnlTitle 
         Height          =   360
         Left            =   15
         TabIndex        =   10
         Top             =   0
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   635
         _Version        =   196609
         ForeColor       =   16777215
         BackColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "   염조제 투입량"
         Alignment       =   1
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCalc 
         Height          =   4425
         Index           =   0
         Left            =   30
         TabIndex        =   11
         Top             =   360
         Width           =   7740
         _ExtentX        =   13653
         _ExtentY        =   7805
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSPanel SSPanel6 
            Height          =   360
            Left            =   45
            TabIndex        =   12
            Top             =   30
            Width           =   7680
            _ExtentX        =   13547
            _ExtentY        =   635
            _Version        =   196609
            BackColor       =   -2147483638
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "염료 투입량"
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin VSFlex7LCtl.VSFlexGrid grdDyeAux 
            Height          =   2085
            Index           =   0
            Left            =   45
            TabIndex        =   13
            Top             =   390
            Width           =   7680
            _cx             =   13547
            _cy             =   3678
            _ConvInfo       =   1
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmRecipeCalcView.frx":011B
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   -1  'True
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
         End
         Begin VSFlex7LCtl.VSFlexGrid grdDyeAux 
            Height          =   1515
            Index           =   1
            Left            =   45
            TabIndex        =   14
            Top             =   2865
            Width           =   7680
            _cx             =   13547
            _cy             =   2672
            _ConvInfo       =   1
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   -1  'True
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
         End
         Begin Threed.SSPanel SSPanel5 
            Height          =   360
            Left            =   45
            TabIndex        =   15
            Top             =   2490
            Width           =   7680
            _ExtentX        =   13547
            _ExtentY        =   635
            _Version        =   196609
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "조제 투입량"
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
      End
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Left            =   11880
      TabIndex        =   16
      Top             =   8490
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      인쇄(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   13575
      TabIndex        =   17
      Top             =   8490
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&C)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSFrame fraSearch 
      Height          =   915
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   1614
      _Version        =   196609
      Begin VB.TextBox txtSearch 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   5
         Left            =   7830
         MaxLength       =   4
         TabIndex        =   25
         Top             =   495
         Width           =   675
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   4
         Left            =   6600
         MaxLength       =   8
         TabIndex        =   24
         Top             =   495
         Width           =   1185
      End
      Begin VB.TextBox txtSearch 
         Height          =   300
         Index           =   1
         Left            =   2850
         TabIndex        =   23
         Top             =   75
         Width           =   1905
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   2820
         TabIndex        =   22
         Top             =   495
         Width           =   1905
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         Left            =   6600
         TabIndex        =   21
         Top             =   75
         Width           =   1905
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "검색(&F)"
         Height          =   780
         Left            =   14370
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   20
         ToolTipText     =   "자료 저장"
         Top             =   60
         Width           =   780
      End
      Begin Threed.SSPanel pnlOrder 
         Height          =   795
         Left            =   60
         TabIndex        =   26
         Top             =   60
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   1402
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.OptionButton optOrder 
            Caption         =   "관리 번호"
            Height          =   180
            Index           =   1
            Left            =   60
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   480
            Value           =   -1  'True
            Width           =   1200
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "Order No."
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   120
            Width           =   1200
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   1
         Left            =   1440
         TabIndex        =   29
         Top             =   75
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "거 래 처"
            Height          =   240
            Index           =   1
            Left            =   60
            TabIndex        =   30
            Top             =   45
            Width           =   975
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   1
         Left            =   4785
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   75
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         _Version        =   196609
         ButtonStyle     =   3
         Outline         =   0   'False
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   3
         Left            =   1440
         TabIndex        =   32
         Top             =   495
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   196609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "품     명"
            Height          =   180
            Index           =   2
            Left            =   60
            TabIndex        =   33
            Top             =   60
            Width           =   975
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   2
         Left            =   4770
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   495
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         _Version        =   196609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   3
         Outline         =   0   'False
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   0
         Left            =   5220
         TabIndex        =   35
         Top             =   75
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   196609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "관리번호"
            Height          =   180
            Index           =   3
            Left            =   60
            TabIndex        =   36
            Top             =   60
            Width           =   1185
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   4
         Left            =   5220
         TabIndex        =   37
         Top             =   495
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   196609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "카드번호"
            Height          =   180
            Index           =   4
            Left            =   60
            TabIndex        =   38
            Top             =   60
            Width           =   1185
         End
      End
   End
   Begin Threed.SSPanel pnlProgress 
      Height          =   870
      Left            =   2130
      TabIndex        =   39
      Top             =   2190
      Visible         =   0   'False
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   1535
      _Version        =   196609
      Alignment       =   0
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      FloodShowPct    =   -1  'True
      Begin MSComctlLib.ProgressBar proProgress 
         Height          =   390
         Left            =   90
         TabIndex        =   40
         Top             =   375
         Width           =   10890
         _ExtentX        =   19209
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lblCount 
         AutoSize        =   -1  'True
         Caption         =   "180"
         Height          =   180
         Left            =   195
         TabIndex        =   41
         Top             =   120
         Width           =   270
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   2685
      Left            =   0
      TabIndex        =   42
      Top             =   930
      Width           =   15225
      _cx             =   26855
      _cy             =   4736
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "frmRecipeCalcView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const REPORTFILE = "\Report\RecipeCalc.rpt"

Private Const LIMIT_ROW1 = 6
Private Const LIMIT_ROW2 = 6
Private Const LIMIT_ROW3 = 23
Private Const LIMIT_WIDTH = 2710
Private Const LIMIT_WIDTH3 = 1250

Private m_bloading  As Boolean

Private m_nDyeID    As Long   ' 스케쥴 번호
Private m_nDyeSeq   As Integer  ' 염색 순위

Private Sub chkSearch_Click(Index As Integer)
    If Index >= 1 And Index <= 3 Then
        If chkSearch(Index).Value = vbChecked Then
            txtSearch(Index).Enabled = True
            txtSearch(Index).SetFocus
            If Index = 1 Or Index = 2 Then
                cmdFind(Index).Enabled = True
            End If
        Else
            txtSearch(Index).Enabled = False
            If Index = 1 Or Index = 2 Then
                cmdFind(Index).Enabled = False
            End If
        End If
    ElseIf Index = 4 Then
        If chkSearch(Index).Value = vbChecked Then
            txtSearch(4).Enabled = True
            txtSearch(5).Enabled = True
            txtSearch(4).SetFocus
        Else
            txtSearch(4).Enabled = False
            txtSearch(5).Enabled = False
        End If
    End If
End Sub


Private Sub cmdExit_Click()
    Unload Me
End Sub


Private Sub cmdFind_Click(Index As Integer)
    If Index = 1 Then
        Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(1))
    ElseIf Index = 2 Then
        Call ReturnCode(LG_ARTICLE, , False, txtSearch(2))
    End If
End Sub

Private Sub cmdPrint_Click()
    Dim oRecipe As PlusLib2.CRecipe
    Dim rs As ADODB.Recordset
    Dim sParam() As String
    Dim i%, nPos%, j%, k%
    Dim sDye() As String
    Dim sAux() As String
    Dim nDyeCnt%, nAuxCnt%
    Dim bFind As Boolean
    Dim sCard$
    
    On Error GoTo ErrHandler
    
    If grdCard.Rows = grdCard.FixedRows Then
        MessageBox "공정카드가 설정되지 않았습니다"
        Exit Sub
    End If
    
    Me.PopupMenu PlusMDI.mnuPopup
    
    ' Printing
    Screen.MousePointer = vbHourglass
    
    Set oRecipe = New PlusLib2.CRecipe
    oRecipe.Connection = g_adoCon
    Set rs = oRecipe.GetDyeCommandOne(m_nDyeID, m_nDyeSeq)

    Set oRecipe = Nothing
    
    ReDim sParam(130)
    ReDim sDye(10)
    ReDim sAux(10)
    
    ' Parameters
    ' 0~9 : 염료명,             10~19: 염료비율,            20~29: 염료사용량,
    ' 30~39: 추가1회 사용량,    40~49: 추가2회 사용량,      50~59: 추가3회 사용량
    ' 60~69: 조제명,            70~79: 조제비율,            80~89: 조제사용량
    ' 90~99: 추가1회 사용량,    100~109: 추가2회 사용량,    110~119: 추가3회 사용량
    ' 120~129: 카드내역
    
    With grdDyeAux(0)
        For i = 0 To .Rows - 2
            sParam(i) = .TextMatrix(i + 1, 1)   ' 염료명
            sDye(i) = .TextMatrix(i + 1, 5)     ' 염료코드 배열(염료, 추가제거분에 대해 위치 찾기위함)
            nDyeCnt = nDyeCnt + 1
            
            sParam(i + 10) = Format(.TextMatrix(i + 1, 2), "#0.000000")  ' 투입비율
            sParam(i + 20) = Format(.TextMatrix(i + 1, 4), "#####0.00")  ' 염료 투입량
        Next i
    End With
    
    With grdDyeAux(1)
        For i = 0 To .Rows - 2
            sParam(i + 60) = .TextMatrix(i + 1, 1)  ' 조제명
            sAux(i) = .TextMatrix(i + 1, 5)         ' 조제코드 배열(염료, 추가제거분에 대해 위치 찾기위함)
            nAuxCnt = nAuxCnt + 1
            
            sParam(i + 70) = Format(.TextMatrix(i + 1, 2), "#0.000000")  ' 투입비율
            sParam(i + 80) = Format(.TextMatrix(i + 1, 4), "#####0.00")  ' 조제 투입량
        Next i
    End With
    
    If m_nDyeSeq > 1 Then
        For i = 2 To m_nDyeSeq
            ' 추가작업 염료 투입량
            With grdDye(i - 2)
                ' 추가작업 염료 그리드 항목 Loop
                For j = 1 To .Rows - 1
                    ' 동일 염료의 출력위치 확인
                    For k = 0 To 9
                        ' 기존 염료 내역중에서 현재 염료 항목의 위치를 찾음 - 출력위치 지정
                        ' 해당 출력 위치에 염료 투입량 입력
                        If sDye(k) = .TextMatrix(j, 4) Then
                            sParam(i * 10 + 10 + k) = Format(.TextMatrix(j, 3), "#####0.00")
                            bFind = True    ' 기존 염료내역 중에서 현재 염료항목을 찾았음을 의미
                        End If
                    Next k
                    
                    ' 기존 염료 내역에서 현재 염료 항목을 찾지 못함
                    ' 기존 염료 내역에 없는 염료항목이 추가되어있을 경우
                    If bFind = False Then
                        ' 새로운 염료 항목을 출력할 염료 항목에 추가
                        sParam(nDyeCnt) = .TextMatrix(j, 1)     ' 염료명
                        sParam(i * 10 + 10 + nDyeCnt) = Format(.TextMatrix(j, 3), "#####0.00")  ' 투입량
                        nDyeCnt = nDyeCnt + 1
                    Else
                        bFind = False
                    End If
                    
                Next j ' 추가작업 염료 그리드 항목 Loop
            
            End With
            
            
            ' 추가작업 조제량 투입량
            With grdAux(i - 2)
                ' 추가작업 조제 그리드 항목 Loop
                For j = 1 To .Rows - 1
                    ' 동일 조제의 출력위치 확인
                    For k = 0 To 9
                        ' 기존 조제 내역중에서 현재 조제 항목의 위치를 찾음 - 출력위치 지정
                        ' 해당 출력 위치에 염료 투입량 입력
                        If sAux(k) = .TextMatrix(j, 4) Then
                            sParam(i * 10 + 70 + k) = Format(.TextMatrix(j, 3), "#####0.00")
                            bFind = True    ' 기존 조제내역 중에서 현재 조제항목을 찾았음을 의미
                        End If
                    Next k
                    
                    ' 기존 조제 내역에서 현재 염료 항목을 찾지 못함
                    ' 기존 조제 내역에 없는 염료항목이 추가되어있을 경우
                    If bFind = False Then
                        ' 새로운 조제 항목을 출력할 조제 항목에 추가
                        sParam(nAuxCnt + 60) = .TextMatrix(j, 1)    ' 조제명
                        sParam(i * 10 + 70 + nAuxCnt) = Format(.TextMatrix(j, 3), "#####0.00")  ' 투입량
                        nAuxCnt = nAuxCnt + 1
                    Else
                        bFind = False
                    End If
                    
                Next j ' 추가작업 조제 그리드 항목 Loop
            
            End With
        
        Next i  ' 염색 차수 m_nDyeSeq
        
    End If
    
    With grdCard
        For i = 1 To .Rows - 1
            sCard = .TextMatrix(i, 1)
            sCard = sCard & IIf(Len(Trim(.TextMatrix(i, 2))) = 0, "", "(" & Trim(.TextMatrix(i, 2)) & ")")
            sParam(119 + i) = sCard
        Next i
    
    End With
    
    sParam(130) = txtRPCalcRemark.Text
    
    Call PrintReport(REPORTFILE, rs, sParam, PlusMDI.PrintPreview)
        
    Screen.MousePointer = vbDefault
    
    Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault
    
    Call ErrorBox(Err.Number, "cmdPrint_Click", Err.Description)
End Sub

Private Sub cmdSearch_Click()
    Call FillGridData
End Sub

Private Sub Form_Activate()
    PlusMDI.pnlMenu.Visible = False
End Sub

Private Sub Form_Deactivate()
    PlusMDI.pnlMenu.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True
    PlusMDI.tbrMain.Buttons("Menu").Value = tbrPressed
    Unload Me
End Sub

Private Sub grdData_RowColChange()
    If m_bloading Then Exit Sub
    Call ShowData
End Sub

Private Sub optOrder_Click(Index As Integer)
    With grdData
        If optOrder(0).Value Then
            .ColWidth(5) = 1350
            .ColWidth(4) = 0
            chkSearch(3).Caption = "Order No."
        Else
            .ColWidth(5) = 0
            .ColWidth(4) = 1350
            chkSearch(3).Caption = "관리번호"
        End If
    End With
End Sub

Private Sub Form_Load()
    Dim i%
    
    Me.Move 0, 0, 15390, 9740
    
    Call InitGrid
    Call SetOperate(Me)

    For i = 1 To 2
        cmdFind(i).Picture = LoadResPicture("FIND", vbResIcon)
        cmdFind(i).Enabled = False
    Next i
    
    For i = 1 To 5
        txtSearch(i).Enabled = False
    Next i
                
    pnlProgress.Visible = False
    
    Call ClearData
    
    Me.Show
End Sub


Public Sub ShowData()
    Dim nRecipeCnt%
    Dim sTitle$
    
    m_nDyeID = grdData.TextMatrix(grdData.Row, 21)
    m_nDyeSeq = grdData.TextMatrix(grdData.Row, 22)
 
    If m_nDyeSeq <= 1 Then
        pnlCalc(0).Enabled = True
        pnlCalc(1).Enabled = False
        tabDye.Tab = 0
        tabAux.Tab = 0
    Else
        pnlCalc(0).Enabled = False
        pnlCalc(1).Enabled = True
        
    End If
    
    ' 염색 지시내역
'    If ShowDyeCommand(m_nDyeID, m_nDyeSeq) = True Then
        ' 염색지시 카드내역
        Call ShowCardList(m_nDyeID, m_nDyeSeq)
        Call ShowMatchData(m_nDyeID, m_nDyeSeq)
                    
        If m_nDyeSeq = 1 Then
            sTitle = "염조제 투입량  - 본작업 평량지시"
            tabDye.Tab = 0
            tabAux.Tab = 0
        ElseIf m_nDyeSeq = 2 Then
            sTitle = "염조제 투입량  - 추가 1회 평량지시"
            tabDye.Tab = m_nDyeSeq - 2
            tabAux.Tab = m_nDyeSeq - 2
        ElseIf m_nDyeSeq = 3 Then
            sTitle = "염조제 투입량  - 추가 2회 평량지시"
            tabDye.Tab = m_nDyeSeq - 2
            tabAux.Tab = m_nDyeSeq - 2
        ElseIf m_nDyeSeq = 4 Then
            sTitle = "염조제 투입량  - 추가 3회 평량지시"
            tabDye.Tab = m_nDyeSeq - 2
            tabAux.Tab = m_nDyeSeq - 2
        End If
'    End If
End Sub



Private Sub ClearData()
    Dim i%
    
    grdCard.Rows = grdCard.FixedRows
    grdDyeAux(0).Rows = grdDyeAux(0).FixedRows
    grdDyeAux(1).Rows = grdDyeAux(1).FixedRows
    
    For i = 0 To 2
        grdDye(i).Rows = grdDye(i).FixedRows
        grdAux(i).Rows = grdAux(i).FixedRows
    Next i
    
    txtRPCalcRemark = ""
End Sub


Private Sub InitGrid()
    Dim i%

    With grdData
        .Redraw = flexRDNone
        .Cols = 23
        
        Call SetVSFlexGrid(grdData)
        .Rows = 1
        .RowHeightMin = 390
        
        .TextArray(0) = " ":
        .TextArray(1) = " ":            .ColWidth(1) = 250
        .TextArray(2) = "밧자":         .ColWidth(2) = 500:             .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "순위":         .ColWidth(3) = 600:             .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "거래처":       .ColWidth(4) = 1800:            .ColAlignment(4) = flexAlignLeftCenter
        .TextArray(5) = "품명":         .ColWidth(5) = 2000:            .ColAlignment(5) = flexAlignLeftCenter
        .TextArray(6) = "관리번호":     .ColWidth(6) = 1350:            .ColAlignment(6) = flexAlignCenterCenter
        .TextArray(7) = "OrderNo":      .ColWidth(7) = 0:               .ColAlignment(7) = flexAlignLeftCenter
        .TextArray(8) = "카드번호":     .ColWidth(8) = 1000:               .ColAlignment(8) = flexAlignCenterCenter
        .TextArray(9) = "분할" & vbCrLf & "번호":     .ColWidth(9) = 500:            .ColAlignment(9) = flexAlignCenterCenter
        .TextArray(10) = "색상명":         .ColWidth(10) = 1500:            .ColAlignment(10) = flexAlignLeftCenter
        .TextArray(11) = "절수":         .ColWidth(11) = 800:            .ColAlignment(11) = flexAlignRightCenter
        .TextArray(12) = "수량":         .ColWidth(12) = 800:            .ColAlignment(12) = flexAlignRightCenter
        
        .TextArray(21) = "스케쥴번호":  .ColWidth(21) = 0
        .TextArray(22) = "차수":        .ColWidth(22) = 0
        
        .ColFormat(11) = "#,##0"
        .ColFormat(12) = "#,##0"
        
        For i = 13 To 20
            .ColWidth(i) = 0
        Next i
        
        .OutlineBar = flexOutlineBarSimple
        .OutlineCol = 1
        
        .ExtendLastCol = False
        .Redraw = flexRDDirect
    End With
    
    With grdCard
        .Cols = 10

        Call SetVSFlexGrid(grdCard)

        .Redraw = flexRDNone

        .TextArray(1) = "카드번호":                 .ColWidth(1) = 1500:      .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "분할" & vbCrLf & "번호":   .ColWidth(2) = 600:       .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "관리번호":                 .ColWidth(3) = 1300:      .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "Order NO":                 .ColWidth(4) = 0
        .TextArray(5) = "투입" & vbCrLf & "절수":   .ColWidth(5) = 600
        .TextArray(6) = "투입" & vbCrLf & "수량":   .ColWidth(6) = 800
        .TextArray(7) = "단위":                     .ColWidth(7) = 600
        .TextArray(8) = "UnitClss":                 .ColWidth(8) = 0
        .TextArray(9) = "색상명":                   .ColWidth(9) = 1000

        .ColFormat(5) = "#,###"
        .ColFormat(6) = "#,###"

        .ExtendLastCol = True
        .WordWrap = False
        
        .Redraw = flexRDDirect
    End With
    
    With grdDyeAux(0)
        .Cols = 6
        Call SetVSFlexGrid(grdDyeAux(0))

        .Redraw = flexRDNone

        .TextArray(1) = "염료명":                           .ColWidth(1) = 3000:    .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "실험실" & vbCrLf & "투입비율":     .ColWidth(2) = 1400:     .ColAlignment(2) = flexAlignRightCenter
        .TextArray(3) = "투입비율" & vbCrLf & "(%)":        .ColWidth(3) = 1400:     .ColAlignment(3) = flexAlignRightCenter
        .TextArray(4) = "염료투입량":                       .ColWidth(4) = 1010:    .ColAlignment(4) = flexAlignRightCenter
        .TextArray(5) = "염료코드":                         .ColWidth(5) = 0

        .ColFormat(2) = "#,##0.000000"
        .ColFormat(3) = "#,##0.000000"
        .ColFormat(4) = "#,##0.00"
                
        .Editable = flexEDKbdMouse
        .HighLight = flexHighlightWithFocus
        .FocusRect = flexFocusHeavy
        
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False

        .WordWrap = False

        .Redraw = flexRDDirect
    End With

    With grdDyeAux(1)
        .Cols = 6
        Call SetVSFlexGrid(grdDyeAux(1))

        .Redraw = flexRDNone

        .TextArray(1) = "조제명":                           .ColWidth(1) = 3000:        .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "실험실" & vbCrLf & "투입비율":     .ColWidth(2) = 1400:         .ColAlignment(2) = flexAlignRightCenter
        .TextArray(3) = "투입비율" & vbCrLf & "(g/ℓ)":     .ColWidth(3) = 1400:         .ColAlignment(3) = flexAlignRightCenter
        .TextArray(4) = "조제투입량":                       .ColWidth(4) = 1010:        .ColAlignment(4) = flexAlignRightCenter
        .TextArray(5) = "조제코드":                             .ColWidth(5) = 0

        .ColFormat(2) = "#,##0.000000"
        .ColFormat(3) = "#,##0.000000"
        .ColFormat(4) = "#,##0.00"
        
        .Editable = flexEDKbdMouse
        .HighLight = flexHighlightWithFocus
        .FocusRect = flexFocusHeavy
        
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False


        .WordWrap = False

        .Redraw = flexRDDirect
    End With

    For i = 0 To 2
        
        With grdDye(i)
            .Cols = 5
            Call SetVSFlexGrid(grdDye(i))
    
            .Redraw = flexRDNone
    
            .TextArray(1) = "염료명":                           .ColWidth(1) = 3000:        .ColAlignment(1) = flexAlignLeftCenter
            .TextArray(2) = "투입비율" & vbCrLf & "(%)":        .ColWidth(2) = 1500:         .ColAlignment(2) = flexAlignRightCenter
            .TextArray(3) = "염료투입량":                       .ColWidth(3) = 1010:        .ColAlignment(3) = flexAlignRightCenter
            .TextArray(4) = "염료코드":                         .ColWidth(4) = 0
    
            .ColFormat(2) = "#,##0.000000"
            .ColFormat(3) = "#,##0.00"
                        
            .ExtendLastCol = True
            .Editable = flexEDKbdMouse
            .HighLight = flexHighlightWithFocus
            .FocusRect = flexFocusHeavy
            
            .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
    
    
            .WordWrap = False
    
            .Redraw = flexRDDirect
        End With
        
        With grdAux(i)
            .Cols = 5
            Call SetVSFlexGrid(grdAux(i))
    
            .Redraw = flexRDNone
    
            .TextArray(1) = "조제명":                           .ColWidth(1) = 3000:        .ColAlignment(1) = flexAlignLeftCenter
            .TextArray(2) = "투입비율" & vbCrLf & "(%)":        .ColWidth(2) = 1500:         .ColAlignment(2) = flexAlignRightCenter
            .TextArray(3) = "조제투입량":                       .ColWidth(3) = 1010:        .ColAlignment(3) = flexAlignRightCenter
            .TextArray(4) = "조제코드":                         .ColWidth(4) = 0
    
            .ColFormat(2) = "#,##0.000000"
            .ColFormat(3) = "#,##0.00"
                        
            .ExtendLastCol = True
            .Editable = flexEDKbdMouse
            .HighLight = flexHighlightWithFocus
            .FocusRect = flexFocusHeavy
            
            .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
        
            .WordWrap = False
    
            .Redraw = flexRDDirect
        End With
        
    Next i


End Sub



Private Sub ClearGrid()
    grdDyeAux(0).Rows = grdDyeAux(0).FixedRows
    grdDyeAux(1).Rows = grdDyeAux(1).FixedRows

End Sub


' 이전 처방내역 출력
Private Sub ShowMatchData(nDyeID As Long, nDyeSeq As Integer)
    Dim oRecipe As PlusLib2.CRecipe
    Dim rs      As Recordset
    Dim nRow%, i%, nSeq%
    Dim sOrderID$, nOrderSeq%, nRecipeSeq%, nModifySeq%
    
    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    Set oRecipe = New PlusLib2.CRecipe
    oRecipe.Connection = g_adoCon
    oRecipe.UserName = g_sUserName
    
    
    Set rs = oRecipe.GetMatch(nDyeID, 1)
    
    sOrderID = rs!OrderID           ' 처방전 관리번호
    nOrderSeq = rs!OrderSeq         '색상순위
    nRecipeSeq = rs!RecipeSeq       '처방순위
    nModifySeq = rs!ModifySeq       '변경순위
    txtRPCalcRemark = CheckNull(rs!RPRateRemark)
    rs.Close
    
    grdDyeAux(0).Rows = grdDyeAux(0).FixedRows
    grdDyeAux(1).Rows = grdDyeAux(1).FixedRows

    ' 실험실에서 처방된 투입비율을 표기
    Set rs = oRecipe.GetRecipeSub(sOrderID, nOrderSeq, 1, nRecipeSeq, 0, "0", nModifySeq)

    Do Until rs.EOF
        If Left(rs!DyeAuxID, 1) = "1" Then
            With grdDyeAux(0)
                .AddItem CStr(.Rows) & vbTab & Trim(rs!DyeAux) & vbTab & rs!DyeAuxRate
            End With
        Else
            With grdDyeAux(1)
                .AddItem CStr(.Rows) & vbTab & Trim(rs!DyeAux) & vbTab & rs!DyeAuxRate
            End With
        End If

        rs.MoveNext
    Loop

    rs.Close
    
    
    '**************************************************************************
    '*
    '* 평량지시 세부내역 확인 (2003-12-02)
    '*
    '*  - 본작업 평량지시 내역
    '* Author : 최승백
    '**************************************************************************
    
    Set rs = oRecipe.GetMatchSub(nDyeID, 1, "1")

    
    For i = 1 To rs.RecordCount

        With grdDyeAux(0)
            .TextMatrix(i, 3) = rs!DyeAuxRate
            .TextMatrix(i, 4) = rs!DyeAuxQty
            .TextMatrix(i, 5) = rs!DyeAuxID
            
            .Col = 3
            .Row = i
            .CellBackColor = IIf(CSng(.TextMatrix(i, 2)) <> CSng(.TextMatrix(i, 3)), vbRed, vbWhite)

        End With
            
        rs.MoveNext
    Next i

    Set rs = oRecipe.GetMatchSub(nDyeID, 1, "0")

    ' 본작업시 처방된 투입 비율및 투입수량 출력
    For i = 1 To rs.RecordCount
    
        With grdDyeAux(1)
            .TextMatrix(i, 3) = rs!DyeAuxRate
            .TextMatrix(i, 4) = rs!DyeAuxQty
            .TextMatrix(i, 5) = rs!DyeAuxID
            
            .Col = 3
            .Row = i
            .CellBackColor = IIf(CSng(.TextMatrix(i, 2)) <> CSng(.TextMatrix(i, 3)), vbRed, vbWhite)

        End With

        rs.MoveNext
    Next i
    '''''
    ' 본작업 평량 지시내역 출력 끝

    
    '**************************************************************************
    '*  - 재작업 평량지시 내역
    '*
    '* Author : 최승백
    '**************************************************************************
    If nDyeSeq > 1 Then
        For nSeq = 2 To nDyeSeq
            ' 평량 새로 작성시에는 바로 이전 단계의 재작업 내역 출력
            ' 변경시에는 재작업 평량 지시내역 모두 출력
            Set rs = oRecipe.GetMatchSub(nDyeID, nSeq, "1")
             
            ' 염료 - 추가작업 투입량
            For i = 1 To rs.RecordCount
                With grdDye(nSeq - 2)
                
                    .AddItem CStr(i) & vbTab & rs!DyeAux & vbTab & rs!DyeAuxRate & vbTab & rs!DyeAuxQty & vbTab & rs!DyeAuxID
                                        
                End With
                    
                rs.MoveNext
            Next i
        
            Set rs = oRecipe.GetMatchSub(nDyeID, nSeq, "0")
        
            ' 조제 - 추가작업 투입량
            For i = 1 To rs.RecordCount
                With grdAux(nSeq - 2)
                    .AddItem CStr(i) & vbTab & rs!DyeAux & vbTab & rs!DyeAuxRate & vbTab & rs!DyeAuxQty & vbTab & rs!DyeAuxID

                End With
        
                rs.MoveNext
            Next i
        Next nSeq
    
    End If
    

    Set rs = Nothing
    Set oRecipe = Nothing
    Screen.MousePointer = vbArrow

    Exit Sub

ErrHandler:
    Screen.MousePointer = vbArrow
    Set rs = Nothing
    Set oRecipe = Nothing
    Call ErrorBox(Err.Number, "frmRecipeCalcView.ShowMatchData", Err.Description)
End Sub


Private Sub txtSearch_GotFocus(Index As Integer)
    Call GotFocusText(txtSearch(Index))
End Sub

Private Sub txtSearch_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Index = 1 Then
            Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(Index))
        ElseIf Index = 2 Then
            Call ReturnCode(LG_ARTICLE, , False, txtSearch(Index))
        End If
        Call NextFocus
    ElseIf KeyAscii = vbKeyReturn And Index >= 3 Then
        Call NextFocus
    End If
End Sub

Private Sub FillGridData()
    Dim oRecipe As PlusLib2.CRecipe
    Dim rs As ADODB.Recordset
    Dim i%, nTop%, nDyeSchID&, nDyeSeq%
    Dim nTRoll%, nTQty%
    
    On Error GoTo ErrHandler
    
    m_bloading = True
    
    proProgress.Value = 0
    lblCount = LoadResString(160)
    pnlProgress.Visible = True
        
    Set oRecipe = New PlusLib2.CRecipe
    oRecipe.Connection = g_adoCon
    
    Set rs = oRecipe.GetRecipeCalcList(IIf(chkSearch(1) = vbChecked, 1, 0), txtSearch(1).Tag, IIf(chkSearch(2) = vbChecked, 1, 0), txtSearch(2).Tag, _
                                 IIf(chkSearch(3) = vbChecked, IIf(optOrder(1).Value = True, 1, 2), 0), txtSearch(3), _
                                 IIf(chkSearch(4) = vbChecked, 1, 0), txtSearch(4), txtSearch(5))
    Set oRecipe = Nothing
        
    With grdData
        .Redraw = flexRDDirect
        .Rows = .FixedRows
    
        For i = 0 To rs.RecordCount - 1
            If nDyeSchID <> rs!DyeSchID Then
                .AddItem ""
                .TextMatrix(.Rows - 1, 2) = rs!DyeSchID
                .TextMatrix(.Rows - 1, 3) = rs!DyeSeq
                .TextMatrix(.Rows - 1, 4) = rs!kCustom
                .TextMatrix(.Rows - 1, 5) = rs!Article
                .TextMatrix(.Rows - 1, 6) = MakeOrderID(rs!OrderID, OM_EXPAND)
                .TextMatrix(.Rows - 1, 7) = rs!OrderNo
                .TextMatrix(.Rows - 1, 8) = MakeCardID(rs!CardID, OM_EXPAND)
                .TextMatrix(.Rows - 1, 9) = rs!SplitID
                .TextMatrix(.Rows - 1, 10) = rs!Color
                .TextMatrix(.Rows - 1, 11) = rs!wiRoll
                .TextMatrix(.Rows - 1, 12) = rs!wiQty
                .TextMatrix(.Rows - 1, 21) = rs!DyeSchID
                .TextMatrix(.Rows - 1, 22) = rs!DyeSeq
                            
                Call DoFlexGridGroup(grdData, .Rows - 1, 1)
                Call GridCollapse(grdData, nTop)
                nTop = .Rows - 1
            End If
                        
            If rs!MaxCardSeq > 1 Then
                .AddItem ""
                .TextMatrix(.Rows - 1, 2) = ""
                .TextMatrix(.Rows - 1, 3) = rs!CardSeq
                .TextMatrix(.Rows - 1, 4) = rs!kCustom
                .TextMatrix(.Rows - 1, 5) = rs!Article
                .TextMatrix(.Rows - 1, 6) = MakeOrderID(rs!OrderID, OM_EXPAND)
                .TextMatrix(.Rows - 1, 7) = rs!OrderNo
                .TextMatrix(.Rows - 1, 8) = MakeCardID(rs!CardID, OM_EXPAND)
                .TextMatrix(.Rows - 1, 9) = rs!SplitID
                .TextMatrix(.Rows - 1, 10) = rs!Color
                .TextMatrix(.Rows - 1, 11) = rs!Roll
                .TextMatrix(.Rows - 1, 12) = rs!Qty
                .TextMatrix(.Rows - 1, 21) = rs!DyeSchID
                .TextMatrix(.Rows - 1, 22) = rs!DyeSeq
            End If
            nDyeSchID = rs!DyeSchID

            lblCount = CStr(i) & " / " & CStr(rs.RecordCount) & "  (" & Format((i / rs.RecordCount) * 100, "00.0") & " %)"
            proProgress.Value = CInt((i / rs.RecordCount) * 100)
            
            rs.MoveNext
        Next i
        rs.Close
        Set rs = Nothing
        Call GridCollapse(grdData, nTop)
        
        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways
            .Row = IIf(.Rows > .FixedRows, .FixedRows, .Rows - 1)
            .TopRow = .Row
            .Col = .FixedCols
            .ColSel = .Cols - 1
        Else
            .HighLight = flexHighlightNever
            
            Call ClearData
            MsgBox LoadResString(203), vbInformation
        End If
        
        .Redraw = flexRDDirect
        .SetFocus
    End With
    
    pnlProgress.Visible = False
    m_bloading = False
    
    If grdData.Rows > grdData.FixedRows Then
        Call ShowData
    End If
    Exit Sub

ErrHandler:
    Set oRecipe = Nothing
    Set rs = Nothing
    pnlProgress.Visible = False
    m_bloading = False
    Call ErrorBox(Err.Number, "frmRecipeCalcView.FillGridData", Err.Description)
End Sub

Private Sub ShowCardList(nDyeID As Long, nSeq As Integer)
    Dim oRecipe As PlusLib2.CRecipe
    Dim rs      As Recordset

    On Error GoTo ErrHandler

    Set oRecipe = New PlusLib2.CRecipe
    oRecipe.Connection = g_adoCon
    oRecipe.UserName = g_sUserName

    Set rs = oRecipe.GetRapidCommandSub(nDyeID, nSeq)
    
    Set oRecipe = Nothing

    With grdCard
        .Redraw = flexRDNone
        .Rows = .FixedRows

        Do Until rs.EOF
            
            .AddItem CStr(.Rows) & vbTab & MakeCardID(rs!CardID, OM_EXPAND) & vbTab & rs!SplitID & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & _
                    rs!OrderNo & vbTab & rs!Roll & vbTab & rs!Qty & vbTab & IIf(rs!UnitClss = 0, "Y", "M") & vbTab & _
                    rs!UnitClss & vbTab & CheckNull(rs!Color)
                    
            rs.MoveNext
        Loop

        .Redraw = flexRDDirect

        If .Rows > .FixedRows Then
            .Row = .FixedRows
        End If

    End With
    rs.Close
    Set rs = Nothing
    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oRecipe = Nothing

    Call ErrorBox(Err.Number, "frmRecipeCalcView.ShowCardList", Err.Description)
End Sub


