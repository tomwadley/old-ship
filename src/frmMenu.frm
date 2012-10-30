VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMenu 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ship 0.4"
   ClientHeight    =   7965
   ClientLeft      =   375
   ClientTop       =   645
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   10485
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   9960
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "xml"
      Filter          =   "xml (*.xml)|*.xml|All Files (*.*)|*.*"
   End
   Begin VB.CommandButton cmdMod 
      Caption         =   "Switch Mod"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   1575
   End
   Begin VB.ListBox lstLevels 
      Height          =   645
      Left            =   8040
      TabIndex        =   6
      Top             =   1560
      Width           =   2295
   End
   Begin VB.ComboBox cmbWeapon2 
      Height          =   315
      Left            =   2040
      TabIndex        =   10
      Text            =   "Weapon 2"
      Top             =   6600
      Width           =   1695
   End
   Begin VB.ComboBox cmbWeapon1 
      Height          =   315
      Left            =   2040
      TabIndex        =   9
      Text            =   "Weapon 1"
      Top             =   6000
      Width           =   1695
   End
   Begin VB.ComboBox cmbPlayerShip 
      Height          =   315
      Left            =   2040
      TabIndex        =   8
      Text            =   "Ship"
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "Read This!"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmdBuyShip 
      Caption         =   "Purchase"
      Height          =   375
      Left            =   9120
      TabIndex        =   14
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdBuyWeapon 
      Caption         =   "Purchase"
      Height          =   375
      Left            =   9120
      TabIndex        =   12
      Top             =   3960
      Width           =   1215
   End
   Begin VB.ComboBox cmbShips 
      Height          =   315
      Left            =   4560
      TabIndex        =   13
      Text            =   "Ships"
      Top             =   5640
      Width           =   1695
   End
   Begin VB.ComboBox cmbWeapons 
      Height          =   315
      ItemData        =   "frmMenu.frx":0000
      Left            =   4560
      List            =   "frmMenu.frx":0002
      TabIndex        =   11
      Text            =   "Weapons"
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton cmdBuyHealth 
      Caption         =   "Buy Shield"
      Height          =   615
      Left            =   9120
      TabIndex        =   15
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Game"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Game"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton cmdStartMission 
      Caption         =   "Start Mission!"
      Height          =   615
      Left            =   8040
      TabIndex        =   7
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton cmdNewGame 
      Caption         =   "New Game"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
   Begin MSComctlLib.ProgressBar pbHealth 
      Height          =   4695
      Left            =   3960
      TabIndex        =   23
      Top             =   3120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   8281
      _Version        =   393216
      Appearance      =   0
      Max             =   1
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.PictureBox picPlayerShip 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   2040
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   29
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label18 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "doogle.co.nr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   46
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "Missions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2040
      TabIndex        =   45
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Your Weapons"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2040
      TabIndex        =   44
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label lblModDescription 
      BackColor       =   &H00000000&
      Caption         =   "Mod Description"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2040
      TabIndex        =   43
      Top             =   600
      Width           =   8175
   End
   Begin VB.Label lblModName 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mod Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2040
      TabIndex        =   42
      Top             =   120
      Width           =   7935
   End
   Begin VB.Image imgLoad 
      Height          =   615
      Left            =   120
      Top             =   8160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label17 
      BackColor       =   &H00000000&
      Caption         =   "How to play:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   41
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   2040
      X2              =   10320
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Press Ctrl in flight to switch between these two weapons"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2040
      TabIndex        =   40
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Label lblLevelDescription 
      BackColor       =   &H00000000&
      Caption         =   "The mission breifing goes here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   2040
      TabIndex        =   39
      Top             =   2040
      Width           =   5775
   End
   Begin VB.Label lblLevelTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Level Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3600
      TabIndex        =   38
      Top             =   1560
      Width           =   4335
   End
   Begin VB.Label Label15 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Ships"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   37
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label14 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Weapons"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   36
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   4440
      X2              =   4440
      Y1              =   7800
      Y2              =   3000
   End
   Begin VB.Line shipPicLine 
      BorderColor     =   &H00FFFFFF&
      Visible         =   0   'False
      X1              =   2040
      X2              =   3720
      Y1              =   3960
      Y2              =   5280
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Caption         =   "Your Ships"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2040
      TabIndex        =   35
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackColor       =   &H00000000&
      Caption         =   "cash:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7200
      TabIndex        =   34
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Repair your ship "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4560
      TabIndex        =   33
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      Caption         =   "Price:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   32
      Top             =   7440
      Width           =   735
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   2040
      X2              =   10320
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   "Weapon slot 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   31
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Weapon slot 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   30
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label lblShipName 
      BackColor       =   &H00000000&
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6480
      TabIndex        =   28
      Top             =   5640
      Width           =   2535
   End
   Begin VB.Label lblShipDescription 
      BackColor       =   &H00000000&
      Caption         =   "This is the ship description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4560
      TabIndex        =   27
      Top             =   6000
      Width           =   5775
   End
   Begin VB.Label lblWeaponDescription 
      BackColor       =   &H00000000&
      Caption         =   "This is the weapon description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4560
      TabIndex        =   26
      Top             =   4320
      Width           =   5775
   End
   Begin VB.Label lblWeaponName 
      BackColor       =   &H00000000&
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6480
      TabIndex        =   25
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   4440
      X2              =   10320
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Label lblSheildPrice 
      BackColor       =   &H00000000&
      Caption         =   "$400 for 20 Shield"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6720
      TabIndex        =   24
      Top             =   7440
      Width           =   2295
   End
   Begin VB.Label lblCash 
      BackColor       =   &H00000000&
      Caption         =   "$9999"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8160
      TabIndex        =   22
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label lblQtyHealth 
      BackColor       =   &H00000000&
      Caption         =   "100 / 100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6840
      TabIndex        =   21
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "Sheild:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   20
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Shop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4560
      TabIndex        =   19
      Top             =   3120
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   1800
      X2              =   1800
      Y1              =   120
      Y2              =   7800
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "v 0.4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1200
      TabIndex        =   18
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   $"frmMenu.frx":0004
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   120
      TabIndex        =   17
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "ship"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Update()
    Dim i As Integer
    If gameLoaded = True Then
        
        If cmbWeapons.ListIndex = -1 Then
            cmbWeapons.Clear
            For i = LBound(weapons) To UBound(weapons)
                If i <> 0 Then
                    If weapons(i).cost >= 0 Then
                        cmbWeapons.AddItem "$" & weapons(i).cost & " " & weapons(i).title
                        cmbWeapons.ItemData(cmbWeapons.NewIndex) = i
                    End If
                End If
            Next i
            lblWeaponName.Caption = ""
            lblWeaponDescription.Caption = ""
        Else
            lblWeaponName.Caption = weapons(cmbWeapons.ItemData(cmbWeapons.ListIndex)).title
            lblWeaponDescription.Caption = weapons(cmbWeapons.ItemData(cmbWeapons.ListIndex)).description
        End If
        
        If cmbShips.ListIndex = -1 Then
            cmbShips.Clear
            For i = LBound(ships) To UBound(ships)
                If i <> 0 Then
                    If ships(i).cost >= 0 Then
                        cmbShips.AddItem "$" & ships(i).cost & " " & ships(i).title
                        cmbShips.ItemData(cmbShips.NewIndex) = i
                    End If
                End If
            Next i
            lblShipName.Caption = ""
            lblShipDescription.Caption = ""
        Else
            lblShipName.Caption = ships(cmbShips.ItemData(cmbShips.ListIndex)).title
            lblShipDescription.Caption = ships(cmbShips.ItemData(cmbShips.ListIndex)).description
        End If
        
        If cmbPlayerShip.ListCount = 0 Then
            cmbPlayerShip.Clear
            For i = LBound(player.shipsOwned) To UBound(player.shipsOwned)
                If i <> 0 Then
                    cmbPlayerShip.AddItem ships(player.shipsOwned(i)).title
                    cmbPlayerShip.ItemData(cmbPlayerShip.NewIndex) = i
                    If player.shipSelected = i Then
                        cmbPlayerShip.ListIndex = cmbPlayerShip.NewIndex
                    End If
                End If
            Next i
        End If
        If player.shipSelected <> -1 Then
            picPlayerShip.Visible = False
            picPlayerShip.Picture = LoadPicture(App.Path & "\" & modName & "\" & pics(ships(player.shipsOwned(player.shipSelected)).pics(1)).filePath)
            picPlayerShip.top = shipPicLine.Y1 + (((shipPicLine.Y2 - shipPicLine.Y1) / 2) - (picPlayerShip.height / 2))
            picPlayerShip.left = shipPicLine.X1 + (((shipPicLine.X2 - shipPicLine.X1) / 2) - (picPlayerShip.width / 2))
            picPlayerShip.Visible = True
        Else
            picPlayerShip.Visible = False
        End If
        
        If cmbWeapon1.ListCount = 0 Or cmbWeapon2.ListCount = 0 Then
            cmbWeapon1.Clear
            cmbWeapon2.Clear
            For i = LBound(player.weaponsOwned) To UBound(player.weaponsOwned)
                If i <> 0 Then
                    cmbWeapon1.AddItem weapons(player.weaponsOwned(i)).title
                    cmbWeapon1.ItemData(cmbWeapon1.NewIndex) = i
                    cmbWeapon2.AddItem weapons(player.weaponsOwned(i)).title
                    cmbWeapon2.ItemData(cmbWeapon2.NewIndex) = i
                    If player.weapon1 = i Then
                        cmbWeapon1.ListIndex = cmbWeapon1.NewIndex
                    End If
                    If UBound(player.weaponsOwned) > 1 Then
                        If player.weapon2 = i Then
                            cmbWeapon2.ListIndex = cmbWeapon2.NewIndex
                        End If
                    End If
                End If
            Next i
        End If
        
        
        'calc list of available levels
        If lstLevels.ListCount = 0 Then
            lstLevels.Clear
            For i = LBound(levels) To UBound(levels)
                If i <> 0 Then
                    Dim deps() As Integer
                    deps = levels(i).dependencies
                    Dim passed As Boolean
                    passed = False
                    Dim e As Integer
                    For e = LBound(player.levelsPassed) To UBound(player.levelsPassed)
                        If e <> 0 Then
                            If player.levelsPassed(e) = i Then
                                passed = True
                            End If
                            Dim p As Integer
                            For p = LBound(deps) To UBound(deps)
                                If p <> 0 Then
                                    If deps(p) = player.levelsPassed(e) Then
                                        deps(p) = -1
                                    End If
                                End If
                            Next p
                        End If
                    Next e
                    Dim inlist As Boolean
                    inlist = True
                    If passed = True Then
                        inlist = False
                    Else
                        For e = LBound(deps) To UBound(deps)
                            If e <> 0 Then
                                If deps(e) <> -1 Then
                                    inlist = False
                                End If
                            End If
                        Next e
                    End If
                    If inlist = True Then
                        'put level in list
                        lstLevels.AddItem levels(i).title
                        lstLevels.ItemData(lstLevels.NewIndex) = i
                        If lstLevels.ListCount = 1 Then
                            lstLevels.ListIndex = lstLevels.NewIndex
                        End If
                    End If
                End If
            Next i
        End If
        
        lstLevels.Enabled = True
        lblCash.Caption = "$" & " " & player.cash
        If player.shipSelected <> -1 Then
            lblQtyHealth.Caption = player.shipHealth(player.shipSelected) & " / " & ships(player.shipsOwned(player.shipSelected)).maxHealth
            pbHealth.Max = ships(player.shipsOwned(player.shipSelected)).maxHealth
            pbHealth.Value = player.shipHealth(player.shipSelected)
        Else
            lblQtyHealth.Caption = ""
            pbHealth.Enabled = False
        End If
        If player.level <> -1 Then
            lblLevelTitle.Caption = levels(player.level).title
            lblLevelDescription.Caption = levels(player.level).description
        Else
            lblLevelTitle.Caption = ""
            lblLevelDescription.Caption = ""
        End If
        
        cmbPlayerShip.Enabled = True
        If UBound(player.weaponsOwned) = 0 Then
            cmbWeapon1.Enabled = False
            cmbWeapon2.Enabled = False
        Else
            cmbWeapon1.Enabled = True
            If UBound(player.weaponsOwned) > 1 Then
                cmbWeapon2.Enabled = True
            Else
                cmbWeapon2.Enabled = False
            End If
        End If
        cmbWeapons.Enabled = True
        cmbShips.Enabled = True
        
        cmdBuyHealth.Enabled = True
        cmdStartMission.Enabled = True
        cmdBuyWeapon.Enabled = True
        cmdBuyShip.Enabled = True
        cmdLoad.Enabled = True
        cmdSave.Enabled = True
        
        lblLevelTitle.Visible = True
        lblLevelDescription.Visible = True
        lblSheildPrice.Visible = True
        lblSheildPrice.Caption = "$" & healthPriceCost & " for " & healthPriceQty & " shield"
        lblModDescription.Caption = modDescription
        lblModName.Caption = modName
    Else
        'set everything to disabled and delete captions and all that
        cmbWeapons.ListIndex = -1
        cmbShips.ListIndex = -1
        
        cmbPlayerShip.Clear
        cmbWeapon1.Clear
        cmbWeapon2.Clear
        lstLevels.Clear
    
        lstLevels.Enabled = False
        lblCash.Caption = Empty
        lblQtyHealth.Caption = Empty
        pbHealth.Value = 0
        lblWeaponName.Caption = Empty
        lblWeaponDescription.Caption = Empty
        lblShipName.Caption = Empty
        lblShipDescription.Caption = Empty
        picPlayerShip.Visible = False
        lblLevelTitle.Visible = False
        lblLevelDescription.Visible = False
        lblSheildPrice.Visible = False
        
        cmbPlayerShip.Enabled = False
        cmbWeapon1.Enabled = False
        cmbWeapon2.Enabled = False
        cmbWeapons.Enabled = False
        cmbShips.Enabled = False
        
        cmdBuyHealth.Enabled = False
        cmdStartMission.Enabled = False
        cmdBuyWeapon.Enabled = False
        cmdBuyShip.Enabled = False
        If modLoaded = True Then
            cmdNewGame.Enabled = True
            lblModName.Caption = modName
            lblModDescription.Caption = modDescription
            lblModName.Caption = modName
            cmdLoad.Enabled = True
            cmdSave.Enabled = False
        Else
            cmdNewGame.Enabled = False
            lblModName.Caption = Empty
            lblModDescription.Caption = Empty
            lblModName.Caption = Empty
            cmdLoad.Enabled = False
            cmdSave.Enabled = False
        End If
    End If
End Sub

Sub newgame()
    If gameLoaded = True Then
        If MsgBox("Are you sure u want to start a new game?", vbOKCancel) = vbCancel Then
            Exit Sub
        End If
    End If
    gameLoaded = False
    Update
    
    savedGamePath = Empty
    
    Dim i As Integer
    Dim e As Integer
    
    ReDim player.shipsOwnedNames(UBound(newgameinfo.shipsOwnedNames))
    player.shipsOwnedNames = newgameinfo.shipsOwnedNames
    ReDim player.weaponsOwnedNames(UBound(newgameinfo.weaponsOwnedNames))
    player.weaponsOwnedNames = newgameinfo.weaponsOwnedNames
    
    ReDim player.weaponsOwned(UBound(player.weaponsOwnedNames))
    For i = LBound(player.weaponsOwnedNames) To UBound(player.weaponsOwnedNames)
        If i <> 0 Then
            player.weaponsOwned(i) = -1
            For e = LBound(weapons) To UBound(weapons)
                If e <> 0 Then
                    If player.weaponsOwnedNames(i) = weapons(e).obName Then
                        player.weaponsOwned(i) = e
                    End If
                End If
            Next e
        End If
    Next i
    ReDim player.shipsOwned(UBound(player.shipsOwnedNames))
    For i = LBound(player.shipsOwnedNames) To UBound(player.shipsOwnedNames)
        If i <> 0 Then
            player.shipsOwned(i) = -1
            For e = LBound(ships) To UBound(ships)
                If e <> 0 Then
                    If player.shipsOwnedNames(i) = ships(e).obName Then
                        player.shipsOwned(i) = e
                    End If
                End If
            Next e
        End If
    Next i
    
    
    'setup player
    player.shipSelected = -1
    ReDim player.shipHealth(UBound(player.shipsOwned))
    For i = LBound(player.shipHealth) To UBound(player.shipHealth)
        If i <> 0 Then
            player.shipHealth(i) = ships(player.shipsOwned(i)).maxHealth
            player.shipSelected = i
        End If
    Next i
    player.alive = True
    player.cash = 0
    ReDim player.levelsPassed(0)
    If UBound(player.weaponsOwned) > 0 Then
        player.weapon1 = 1
        If UBound(player.weaponsOwned) > 1 Then
            player.weapon2 = 2
        Else
            player.weapon2 = -1
        End If
    Else
        player.weapon1 = -1
        player.weapon2 = -1
    End If
    player.level = -1
    
    
    backupPlayer = player
    
    gameLoaded = True
    
    Update
End Sub

Sub LoadMod()
    Dim xml_document As DOMDocument
    Dim root As IXMLDOMNode
    Set xml_document = New DOMDocument
    
    xml_document.Load App.Path & "\" & modFileName

    'check if the file was opened correctly
    If xml_document.documentElement Is Nothing Then
        MsgBox modFileName & " could not be loaded. It is not a valid ship mod file."
        Exit Sub
    End If
    Dim dtd As IXMLDOMDocumentType
    Set dtd = xml_document.doctype
    If dtd Is Nothing Then
        MsgBox modFileName & " does not reference any DTD. It should reference shipmod.dtd"
        Exit Sub
    Else
        If dtd.Name <> "shipmod" Then
            MsgBox modFileName & " references the " & dtd.Name & " dtd. It should reference shipmod.dtd"
            Exit Sub
        End If
    End If

    'find the shipmod node
    Set root = xml_document.selectSingleNode("shipmod")
    
    'reset vars
    ReDim enemies(0)
    ReDim ships(0)
    ReDim weapons(0)
    ReDim backgroundobjects(0)
    ReDim levels(0)
    ReDim pics(0)
    ReDim newgameinfo.weaponsOwnedNames(0)
    ReDim newgameinfo.shipsOwnedNames(0)
    
    gameLoaded = False
    modLoaded = False
    
    'load shipmod name
    Dim nameAttribute As IXMLDOMAttribute
    Set nameAttribute = root.Attributes.getNamedItem("name")
    modName = nameAttribute.Text
    Set nameAttribute = Nothing
    
    'load shipmod version
    Dim versionAttribute As IXMLDOMAttribute
    Set versionAttribute = root.Attributes.getNamedItem("shipmodversion")
    If versionAttribute.Text <> fileVersion Then
        If MsgBox(modFileName & " is a version " & versionAttribute.Text & " ship mod. This is ship 0.4 which is only designed to load version " & fileVersion & " ship mod files. Loading this file will most likely fail or have unexpected results. Would you like to continue loading this file anyway?", vbOKCancel) = vbCancel Then
            Exit Sub
        End If
    End If
    Set versionAttribute = Nothing
    
    
    Dim node1 As IXMLDOMNode
    Dim node2 As IXMLDOMNode
    Dim node3 As IXMLDOMNode
    Dim nodelist1 As IXMLDOMNodeList
    Dim nodelist2 As IXMLDOMNodeList
    Dim nodelist3 As IXMLDOMNodeList
    Dim i As Integer
    Dim e As Integer
    Dim p As Integer
    
    'load description
    Set node1 = root.selectSingleNode("description")
    modDescription = node1.Text
    
    'load description
    Set node1 = root.selectSingleNode("picMaskColour")
    picMaskColour = node1.Text
    
    'get newGame 'change this!!!!!!!!!!!!!!!!!!!!
    Set node1 = root.selectSingleNode("newGame")
        Set node2 = node1.selectSingleNode("weaponsOwnedNames")
            Set nodelist1 = node2.selectNodes("weaponName")
                i = 0
                While i < nodelist1.length
                    ReDim Preserve newgameinfo.weaponsOwnedNames(UBound(newgameinfo.weaponsOwnedNames) + 1)
                    newgameinfo.weaponsOwnedNames(UBound(newgameinfo.weaponsOwnedNames)) = nodelist1.Item(i).Text
                    i = i + 1
                Wend
        Set node2 = node1.selectSingleNode("shipsOwnedNames")
            Set nodelist1 = node2.selectNodes("shipName")
                i = 0
                While i < nodelist1.length
                    ReDim Preserve newgameinfo.shipsOwnedNames(UBound(newgameinfo.shipsOwnedNames) + 1)
                    newgameinfo.shipsOwnedNames(UBound(newgameinfo.shipsOwnedNames)) = nodelist1.Item(i).Text
                    i = i + 1
                Wend
    
    'load healthPrice
    Set node1 = root.selectSingleNode("healthPrice")
        Set node2 = node1.selectSingleNode("qty")
            healthPriceQty = node2.Text
        Set node2 = node1.selectSingleNode("cost")
            healthPriceCost = node2.Text
    
    'get enemies
    Set node1 = root.selectSingleNode("enemies")
        Set nodelist1 = node1.selectNodes("enemy")
            i = 0
            While i < nodelist1.length
                ReDim Preserve enemies(UBound(enemies) + 1)
                ReDim enemies(UBound(enemies)).pics(0)
                ReDim enemies(UBound(enemies)).picsDead(0)
                Set node2 = nodelist1.Item(i).selectSingleNode("obName")
                    enemies(UBound(enemies)).obName = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("minMoveTop")
                    enemies(UBound(enemies)).minMoveTop = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("maxMoveTop")
                    enemies(UBound(enemies)).maxMoveTop = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("minMoveLeft")
                    enemies(UBound(enemies)).minMoveLeft = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("maxMoveLeft")
                    enemies(UBound(enemies)).maxMoveLeft = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("maxHealth")
                    enemies(UBound(enemies)).maxHealth = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("weaponName")
                    enemies(UBound(enemies)).weaponName = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("pics")
                    Set nodelist2 = node2.selectNodes("pic")
                        e = 0
                        While e < nodelist2.length
                            ReDim Preserve enemies(UBound(enemies)).pics(UBound(enemies(UBound(enemies)).pics) + 1)
                            enemies(UBound(enemies)).pics(UBound(enemies(UBound(enemies)).pics)) = LoadPic(nodelist2.Item(e).Text)
                            e = e + 1
                        Wend
                Set node2 = nodelist1.Item(i).selectSingleNode("timeBetweenPics")
                    enemies(UBound(enemies)).timeBetweenPics = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("picsDead")
                    Set nodelist2 = node2.selectNodes("pic")
                        e = 0
                        While e < nodelist2.length
                            ReDim Preserve enemies(UBound(enemies)).picsDead(UBound(enemies(UBound(enemies)).picsDead) + 1)
                            enemies(UBound(enemies)).picsDead(UBound(enemies(UBound(enemies)).picsDead)) = LoadPic(nodelist2.Item(e).Text)
                            e = e + 1
                        Wend
                Set node2 = nodelist1.Item(i).selectSingleNode("timeBetweenPicsDead")
                    enemies(UBound(enemies)).timeBetweenPicsDead = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("cash")
                    enemies(UBound(enemies)).cash = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("initReloadTime")
                    enemies(UBound(enemies)).initReloadTime = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("entryPoint")
                    enemies(UBound(enemies)).entryPoint = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("entrySide")
                    enemies(UBound(enemies)).entrySide = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("entrySideChange")
                    enemies(UBound(enemies)).entrySideChange = node2.Text
                i = i + 1
            Wend

    'get ships
    Set node1 = root.selectSingleNode("ships")
        Set nodelist1 = node1.selectNodes("ship")
            i = 0
            While i < nodelist1.length
                ReDim Preserve ships(UBound(ships) + 1)
                ReDim ships(UBound(ships)).pics(0)
                ReDim ships(UBound(ships)).picsDead(0)
                Set node2 = nodelist1.Item(i).selectSingleNode("obName")
                    ships(UBound(ships)).obName = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("moveSpeed")
                    ships(UBound(ships)).moveSpeed = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("maxHealth")
                    ships(UBound(ships)).maxHealth = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("reloadTimeMultiplier")
                    ships(UBound(ships)).reloadTimeMultiplier = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("collisionWeaponName")
                    ships(UBound(ships)).collisionWeaponName = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("collisionProjectileFromWeaponName")
                    ships(UBound(ships)).collisionProjectileFromWeaponName = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("pics")
                    Set nodelist2 = node2.selectNodes("pic")
                        e = 0
                        While e < nodelist2.length
                            ReDim Preserve ships(UBound(ships)).pics(UBound(ships(UBound(ships)).pics) + 1)
                            ships(UBound(ships)).pics(UBound(ships(UBound(ships)).pics)) = LoadPic(nodelist2.Item(e).Text)
                            e = e + 1
                        Wend
                Set node2 = nodelist1.Item(i).selectSingleNode("timeBetweenPics")
                    ships(UBound(ships)).timeBetweenPics = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("picsDead")
                    Set nodelist2 = node2.selectNodes("pic")
                        e = 0
                        While e < nodelist2.length
                            ReDim Preserve ships(UBound(ships)).picsDead(UBound(ships(UBound(ships)).picsDead) + 1)
                            ships(UBound(ships)).picsDead(UBound(ships(UBound(ships)).picsDead)) = LoadPic(nodelist2.Item(e).Text)
                            e = e + 1
                        Wend
                Set node2 = nodelist1.Item(i).selectSingleNode("timeBetweenPicsDead")
                    ships(UBound(ships)).timeBetweenPicsDead = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("cost")
                    ships(UBound(ships)).cost = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("description")
                    ships(UBound(ships)).description = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("title")
                    ships(UBound(ships)).title = node2.Text
                i = i + 1
            Wend

    'get weapons
    Set node1 = root.selectSingleNode("weapons")
        Set nodelist1 = node1.selectNodes("weapon")
            i = 0
            While i < nodelist1.length
                ReDim Preserve weapons(UBound(weapons) + 1)
                ReDim weapons(UBound(weapons)).projectiles(0)
                Set node2 = nodelist1.Item(i).selectSingleNode("obName")
                    weapons(UBound(weapons)).obName = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("projectiles")
                    Set nodelist2 = node2.selectNodes("projectile")
                        e = 0
                        While e < nodelist2.length
                            ReDim Preserve weapons(UBound(weapons)).projectiles(UBound(weapons(UBound(weapons)).projectiles) + 1)
                            ReDim weapons(UBound(weapons)).projectiles(UBound(weapons(UBound(weapons)).projectiles)).pics(0)
                            ReDim weapons(UBound(weapons)).projectiles(UBound(weapons(UBound(weapons)).projectiles)).picsDead(0)
                            Set node3 = nodelist2.Item(e).selectSingleNode("obName")
                                weapons(UBound(weapons)).projectiles(UBound(weapons(UBound(weapons)).projectiles)).obName = node3.Text
                            Set node3 = nodelist2.Item(e).selectSingleNode("column")
                                weapons(UBound(weapons)).projectiles(UBound(weapons(UBound(weapons)).projectiles)).column = node3.Text
                            Set node3 = nodelist2.Item(e).selectSingleNode("row")
                                weapons(UBound(weapons)).projectiles(UBound(weapons(UBound(weapons)).projectiles)).row = node3.Text
                            Set node3 = nodelist2.Item(e).selectSingleNode("customColumn")
                                weapons(UBound(weapons)).projectiles(UBound(weapons(UBound(weapons)).projectiles)).customColumn = node3.Text
                            Set node3 = nodelist2.Item(e).selectSingleNode("customRow")
                                weapons(UBound(weapons)).projectiles(UBound(weapons(UBound(weapons)).projectiles)).customRow = node3.Text
                            Set node3 = nodelist2.Item(e).selectSingleNode("moveTop")
                                weapons(UBound(weapons)).projectiles(UBound(weapons(UBound(weapons)).projectiles)).moveTop = node3.Text
                            Set node3 = nodelist2.Item(e).selectSingleNode("moveLeft")
                                weapons(UBound(weapons)).projectiles(UBound(weapons(UBound(weapons)).projectiles)).moveLeft = node3.Text
                            Set node3 = nodelist2.Item(e).selectSingleNode("tracksPlayer")
                                weapons(UBound(weapons)).projectiles(UBound(weapons(UBound(weapons)).projectiles)).tracksPlayer = node3.Text
                            Set node3 = nodelist2.Item(e).selectSingleNode("pics")
                                Set nodelist3 = node3.selectNodes("pic")
                                p = 0
                                While p < nodelist3.length
                                    ReDim Preserve weapons(UBound(weapons)).projectiles(UBound(weapons(UBound(weapons)).projectiles)).pics(UBound(weapons(UBound(weapons)).projectiles(UBound(weapons(UBound(weapons)).projectiles)).pics) + 1)
                                    weapons(UBound(weapons)).projectiles(UBound(weapons(UBound(weapons)).projectiles)).pics(UBound(weapons(UBound(weapons)).projectiles(UBound(weapons(UBound(weapons)).projectiles)).pics)) = LoadPic(nodelist3.Item(p).Text)
                                    p = p + 1
                                Wend
                            Set node3 = nodelist2.Item(e).selectSingleNode("timeBetweenPics")
                                weapons(UBound(weapons)).projectiles(UBound(weapons(UBound(weapons)).projectiles)).timeBetweenPics = node3.Text
                            Set node3 = nodelist2.Item(e).selectSingleNode("picsDead")
                                Set nodelist3 = node3.selectNodes("pic")
                                p = 0
                                While p < nodelist3.length
                                    ReDim Preserve weapons(UBound(weapons)).projectiles(UBound(weapons(UBound(weapons)).projectiles)).picsDead(UBound(weapons(UBound(weapons)).projectiles(UBound(weapons(UBound(weapons)).projectiles)).picsDead) + 1)
                                    weapons(UBound(weapons)).projectiles(UBound(weapons(UBound(weapons)).projectiles)).picsDead(UBound(weapons(UBound(weapons)).projectiles(UBound(weapons(UBound(weapons)).projectiles)).picsDead)) = LoadPic(nodelist3.Item(p).Text)
                                    p = p + 1
                                Wend
                            Set node3 = nodelist2.Item(e).selectSingleNode("timeBetweenPicsDead")
                                weapons(UBound(weapons)).projectiles(UBound(weapons(UBound(weapons)).projectiles)).timeBetweenPicsDead = node3.Text
                            Set node3 = nodelist2.Item(e).selectSingleNode("damage")
                                weapons(UBound(weapons)).projectiles(UBound(weapons(UBound(weapons)).projectiles)).damage = node3.Text
                            Set node3 = nodelist2.Item(e).selectSingleNode("reloadTime")
                                weapons(UBound(weapons)).projectiles(UBound(weapons(UBound(weapons)).projectiles)).reloadTime = node3.Text
                            Set node3 = nodelist2.Item(e).selectSingleNode("initReloadTime")
                                weapons(UBound(weapons)).projectiles(UBound(weapons(UBound(weapons)).projectiles)).initReloadTime = node3.Text
                            e = e + 1
                        Wend
                Set node2 = nodelist1.Item(i).selectSingleNode("cost")
                    weapons(UBound(weapons)).cost = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("description")
                    weapons(UBound(weapons)).description = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("title")
                    weapons(UBound(weapons)).title = node2.Text
                i = i + 1
            Wend
    
    'get backgroundobjects
    Set node1 = root.selectSingleNode("backgroundobjects")
        Set nodelist1 = node1.selectNodes("backgroundobject")
            i = 0
            While i < nodelist1.length
                ReDim Preserve backgroundobjects(UBound(backgroundobjects) + 1)
                ReDim backgroundobjects(UBound(backgroundobjects)).pics(0)
                Set node2 = nodelist1.Item(i).selectSingleNode("obName")
                    backgroundobjects(UBound(backgroundobjects)).obName = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("moveSpeed")
                    backgroundobjects(UBound(backgroundobjects)).moveSpeed = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("pics")
                    Set nodelist2 = node2.selectNodes("pic")
                        e = 0
                        While e < nodelist2.length
                            ReDim Preserve backgroundobjects(UBound(backgroundobjects)).pics(UBound(backgroundobjects(UBound(backgroundobjects)).pics) + 1)
                            backgroundobjects(UBound(backgroundobjects)).pics(UBound(backgroundobjects(UBound(backgroundobjects)).pics)) = LoadPic(nodelist2.Item(e).Text)
                            e = e + 1
                        Wend
                Set node2 = nodelist1.Item(i).selectSingleNode("timeBetweenPics")
                    backgroundobjects(UBound(backgroundobjects)).timeBetweenPics = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("chanceOfCreation")
                    backgroundobjects(UBound(backgroundobjects)).chanceOfCreation = node2.Text
                i = i + 1
            Wend
            
    'get levels
    Set node1 = root.selectSingleNode("levels")
        Set nodelist1 = node1.selectNodes("level")
            i = 0
            While i < nodelist1.length
                ReDim Preserve levels(UBound(levels) + 1)
                ReDim levels(UBound(levels)).dependenciesNames(0)
                ReDim levels(UBound(levels)).enemiesPresent(0)
                ReDim levels(UBound(levels)).backgroundobjectsPresentNames(0)
                Set node2 = nodelist1.Item(i).selectSingleNode("obName")
                    levels(UBound(levels)).obName = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("title")
                    levels(UBound(levels)).title = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("description")
                    levels(UBound(levels)).description = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("dependenciesNames")
                    Set nodelist2 = node2.selectNodes("dependencieName")
                        e = 0
                        While e < nodelist2.length
                            ReDim Preserve levels(UBound(levels)).dependenciesNames(UBound(levels(UBound(levels)).dependenciesNames) + 1)
                            levels(UBound(levels)).dependenciesNames(UBound(levels(UBound(levels)).dependenciesNames)) = nodelist2.Item(e).Text
                            e = e + 1
                        Wend
                Set node2 = nodelist1.Item(i).selectSingleNode("enemiesPresent")
                    Set nodelist2 = node2.selectNodes("enemyPresent")
                        e = 0
                        While e < nodelist2.length
                            ReDim Preserve levels(UBound(levels)).enemiesPresent(UBound(levels(UBound(levels)).enemiesPresent) + 1)
                            Set node3 = nodelist2.Item(e).selectSingleNode("enemyName")
                                levels(UBound(levels)).enemiesPresent(UBound(levels(UBound(levels)).enemiesPresent)).enemyName = node3.Text
                            Set node3 = nodelist2.Item(e).selectSingleNode("chanceOfCreation")
                                levels(UBound(levels)).enemiesPresent(UBound(levels(UBound(levels)).enemiesPresent)).chanceOfCreation = node3.Text
                            Set node3 = nodelist2.Item(e).selectSingleNode("maxOnScreen")
                                levels(UBound(levels)).enemiesPresent(UBound(levels(UBound(levels)).enemiesPresent)).maxOnScreen = node3.Text
                            e = e + 1
                        Wend
                Set node2 = nodelist1.Item(i).selectSingleNode("levelTime")
                    levels(UBound(levels)).levelTime = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("backgroundobjectsPresentNames")
                    Set nodelist2 = node2.selectNodes("backgroundobjectPresentName")
                        e = 0
                        While e < nodelist2.length
                            ReDim Preserve levels(UBound(levels)).backgroundobjectsPresentNames(UBound(levels(UBound(levels)).backgroundobjectsPresentNames) + 1)
                            levels(UBound(levels)).backgroundobjectsPresentNames(UBound(levels(UBound(levels)).backgroundobjectsPresentNames)) = nodelist2.Item(e).Text
                            e = e + 1
                        Wend
                Set node2 = nodelist1.Item(i).selectSingleNode("bossName")
                    levels(UBound(levels)).bossName = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("backgroundColour")
                    levels(UBound(levels)).backgroundColour = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("backImageBackgroundobjectName")
                    levels(UBound(levels)).backImageBackgroundobjectName = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("enemiesStartTime")
                    levels(UBound(levels)).enemiesStartTime = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("bossFinalX")
                    levels(UBound(levels)).bossFinalX = node2.Text
                Set node2 = nodelist1.Item(i).selectSingleNode("bossFinalY")
                    levels(UBound(levels)).bossFinalY = node2.Text
                i = i + 1
            Wend


    Set node1 = Nothing
    Set node2 = Nothing
    Set node3 = Nothing
    Set nodelist1 = Nothing
    Set nodelist2 = Nothing
    Set nodelist3 = Nothing
    
    'translate names into numbers
    For i = LBound(enemies) To UBound(enemies)
        If i <> 0 Then
            enemies(i).weapon = -1
            For e = LBound(weapons) To UBound(weapons)
                If e <> 0 Then
                    If enemies(i).weaponName = weapons(e).obName Then
                        enemies(i).weapon = e
                    End If
                End If
            Next e
        End If
    Next i
    For i = LBound(ships) To UBound(ships)
        If i <> 0 Then
            ships(i).collisionWeapon = -1
            For e = LBound(weapons) To UBound(weapons)
                If e <> 0 Then
                    If ships(i).collisionWeaponName = weapons(e).obName Then
                        ships(i).collisionWeapon = e
                        ships(i).collisionProjectileFromWeapon = -1
                        For p = LBound(weapons(e).projectiles) To UBound(weapons(e).projectiles)
                            If p <> 0 Then
                                If ships(i).collisionProjectileFromWeaponName = weapons(e).projectiles(p).obName Then
                                    ships(i).collisionProjectileFromWeapon = p
                                End If
                            End If
                        Next p
                    End If
                End If
            Next e
        End If
    Next i
    For i = LBound(levels) To UBound(levels)
        If i <> 0 Then
            For e = LBound(levels(i).enemiesPresent) To UBound(levels(i).enemiesPresent)
                If e <> 0 Then
                    levels(i).enemiesPresent(e).enemy = -1
                    For p = LBound(enemies) To UBound(enemies)
                        If p <> 0 Then
                            If levels(i).enemiesPresent(e).enemyName = enemies(p).obName Then
                                levels(i).enemiesPresent(e).enemy = p
                            End If
                        End If
                    Next p
                End If
            Next e
            levels(i).backImageBackgroundobject = -1
            For e = LBound(backgroundobjects) To UBound(backgroundobjects)
                If e <> 0 Then
                    If levels(i).backImageBackgroundobjectName = backgroundobjects(e).obName Then
                        levels(i).backImageBackgroundobject = e
                    End If
                End If
            Next e
            levels(i).boss = -1
            For e = LBound(enemies) To UBound(enemies)
                If e <> 0 Then
                    If levels(i).bossName = enemies(e).obName Then
                        levels(i).boss = e
                    End If
                End If
            Next e
            ReDim levels(i).backgroundobjectsPresent(UBound(levels(i).backgroundobjectsPresentNames))
            For e = LBound(levels(i).backgroundobjectsPresentNames) To UBound(levels(i).backgroundobjectsPresentNames)
                If e <> 0 Then
                    levels(i).backgroundobjectsPresent(e) = -1
                    For p = LBound(backgroundobjects) To UBound(backgroundobjects)
                        If p <> 0 Then
                            If levels(i).backgroundobjectsPresentNames(e) = backgroundobjects(p).obName Then
                                levels(i).backgroundobjectsPresent(e) = p
                            End If
                        End If
                    Next p
                End If
            Next e
            ReDim levels(i).dependencies(UBound(levels(i).dependenciesNames))
            For e = LBound(levels(i).dependenciesNames) To UBound(levels(i).dependenciesNames)
                If e <> 0 Then
                    levels(i).dependencies(e) = -1
                    For p = LBound(levels) To UBound(levels)
                        If p <> 0 Then
                            If levels(i).dependenciesNames(e) = levels(p).obName Then
                                levels(i).dependencies(e) = p
                            End If
                        End If
                    Next p
                End If
            Next e
        End If
    Next i

    modLoaded = True
End Sub

Function LoadPic(fileName As String) As Integer
    'check if the pic has already been loaded
    Dim i As Integer
    For i = LBound(pics) To UBound(pics)
        If i <> 0 Then
            If pics(i).filePath = fileName Then
                LoadPic = i
                Exit Function
            End If
        End If
    Next i
    'the pic has not been loaded yet
    ReDim Preserve pics(UBound(pics) + 1)
    pics(UBound(pics)).filePath = fileName
    On Error Resume Next
    imgLoad.Picture = LoadPicture(App.Path & "\" & modName & "\" & fileName)
    On Error GoTo 0
    If imgLoad.Picture = Empty Then
        MsgBox "Could not open " & App.Path & "\" & modName & "\" & fileName
    End If
    pics(UBound(pics)).height = imgLoad.height
    pics(UBound(pics)).width = imgLoad.width
    imgLoad.Picture = LoadPicture
    LoadPic = UBound(pics)
End Function

Sub StartMission()
    If player.shipSelected > 0 And player.level > 0 And (player.weapon1 > 0 Or player.weapon2 > 0) Then
        cmbPlayerShip.Clear
        cmbWeapon1.Clear
        cmbWeapon2.Clear
        lstLevels.Clear
        lstLevels.SetFocus
    
        Me.Hide
        frmGame.Show
        frmGame.StartGame
    Else
        MsgBox "Before you can start a mission you must have selected a mission, a ship and at least one weapon"
    End If
End Sub

Private Sub cmbPlayerShip_Click()
    If cmbPlayerShip.ListIndex >= 0 Then
        player.shipSelected = cmbPlayerShip.ItemData(cmbPlayerShip.ListIndex)
    End If
    Update
End Sub

Private Sub cmbShips_Click()
    Update
End Sub

Private Sub cmbWeapon1_Click()
    If cmbWeapon1.ListIndex >= 0 Then
        player.weapon1 = cmbWeapon1.ItemData(cmbWeapon1.ListIndex)
    End If
    Update
End Sub

Private Sub cmbWeapon2_Click()
    If cmbWeapon2.ListIndex >= 0 Then
        player.weapon2 = cmbWeapon2.ItemData(cmbWeapon2.ListIndex)
    End If
    Update
End Sub

Private Sub cmbWeapons_Click()
    Update
End Sub

Private Sub cmdAbout_Click()
    frmAbout.Show
End Sub

Private Sub cmdBuyHealth_Click()
    If player.shipSelected <> -1 Then
        If player.shipHealth(player.shipSelected) < ships(player.shipsOwned(player.shipSelected)).maxHealth Then
            If player.cash >= healthPriceCost Then
                player.cash = player.cash - healthPriceCost
                player.shipHealth(player.shipSelected) = player.shipHealth(player.shipSelected) + healthPriceQty
                If player.shipHealth(player.shipSelected) > ships(player.shipsOwned(player.shipSelected)).maxHealth Then
                    player.shipHealth(player.shipSelected) = ships(player.shipsOwned(player.shipSelected)).maxHealth
                End If
            Else
                MsgBox "You cant afford any more shield!"
            End If
        Else
            MsgBox "This ship already has full shield"
        End If
        Update
    Else
        MsgBox "Select a ship first"
    End If
End Sub

Private Sub cmdBuyShip_Click()
    If cmbShips.ListIndex >= 0 Then
        Dim i As Integer
        Dim alreadyOwned As Boolean
        alreadyOwned = False
        For i = LBound(player.shipsOwned) To UBound(player.shipsOwned)
            If player.shipsOwned(i) = cmbShips.ItemData(cmbShips.ListIndex) Then
                alreadyOwned = True
            End If
        Next i
        If alreadyOwned = False Then
            If player.cash >= ships(cmbShips.ItemData(cmbShips.ListIndex)).cost Then
                player.cash = player.cash - ships(cmbShips.ItemData(cmbShips.ListIndex)).cost
                ReDim Preserve player.shipsOwned(UBound(player.shipsOwned) + 1)
                player.shipsOwned(UBound(player.shipsOwned)) = cmbShips.ItemData(cmbShips.ListIndex)
                player.shipSelected = UBound(player.shipsOwned)
                ReDim Preserve player.shipHealth(UBound(player.shipHealth) + 1)
                player.shipHealth(UBound(player.shipHealth)) = ships(cmbShips.ItemData(cmbShips.ListIndex)).maxHealth
                cmbPlayerShip.Clear
            Else
                MsgBox "You cant afford this ship!"
            End If
        Else
            MsgBox "You already own this ship!"
        End If
    Else
        MsgBox "Please select a ship from the list"
    End If
    Update
End Sub

Private Sub cmdBuyWeapon_Click()
    If cmbWeapons.ListIndex >= 0 Then
        Dim i As Integer
        Dim alreadyOwned As Boolean
        alreadyOwned = False
        For i = LBound(player.weaponsOwned) To UBound(player.weaponsOwned)
            If player.weaponsOwned(i) = cmbWeapons.ItemData(cmbWeapons.ListIndex) Then
                alreadyOwned = True
            End If
        Next i
        If alreadyOwned = False Then
            If player.cash >= weapons(cmbWeapons.ItemData(cmbWeapons.ListIndex)).cost Then
                player.cash = player.cash - weapons(cmbWeapons.ItemData(cmbWeapons.ListIndex)).cost
                ReDim Preserve player.weaponsOwned(UBound(player.weaponsOwned) + 1)
                player.weaponsOwned(UBound(player.weaponsOwned)) = cmbWeapons.ItemData(cmbWeapons.ListIndex)
                player.weaponSelected = UBound(player.weaponsOwned)
                player.weapon2 = player.weapon1
                player.weapon1 = UBound(player.weaponsOwned)
                cmbWeapon1.Clear
                cmbWeapon2.Clear
            Else
                MsgBox "You cant afford this weapon!"
            End If
        Else
            MsgBox "You already own this weapon!"
        End If
    Else
        MsgBox "Please select a weapon from the list"
    End If
    Update
End Sub

Sub WriteShipSave()
    Dim i As Integer
        
    Dim openDl As Boolean
    openDl = True
    If savedGamePath <> Empty Then
        If MsgBox("Save to " & savedGamePath & "? Click Cancel to open a save dialogue", vbOKCancel) = vbOK Then
            openDl = False
        End If
    End If
    If openDl = True Then
        dlgMain.fileName = ""
        dlgMain.DialogTitle = "Save Game"
        dlgMain.InitDir = "h:"
        dlgMain.ShowSave
        If Len(dlgMain.fileName) = 0 Then
            Exit Sub
        End If
        savedGamePath = dlgMain.fileName
    End If
    
    'there is no error checking in this code fragment (check for -1)
    backupPlayer.weapon1Name = Empty
    backupPlayer.weapon2Name = Empty
    ReDim backupPlayer.weaponsOwnedNames(UBound(backupPlayer.weaponsOwned))
    For i = LBound(backupPlayer.weaponsOwned) To UBound(backupPlayer.weaponsOwned)
        If i <> 0 Then
            backupPlayer.weaponsOwnedNames(i) = Empty
            If backupPlayer.weaponsOwned(i) > 0 Then
                backupPlayer.weaponsOwnedNames(i) = weapons(player.weaponsOwned(i)).obName
                If backupPlayer.weapon1 = i Then
                    backupPlayer.weapon1Name = weapons(player.weaponsOwned(i)).obName
                End If
                If backupPlayer.weapon2 = i Then
                    backupPlayer.weapon2Name = weapons(player.weaponsOwned(i)).obName
                End If
            End If
        End If
    Next i
    backupPlayer.shipSelectedName = Empty
    ReDim backupPlayer.shipsOwnedNames(UBound(backupPlayer.shipsOwned))
    For i = LBound(backupPlayer.shipsOwned) To UBound(backupPlayer.shipsOwned)
        If i <> 0 Then
            backupPlayer.shipsOwnedNames(i) = Empty
            If backupPlayer.shipsOwned(i) Then
                backupPlayer.shipsOwnedNames(i) = ships(backupPlayer.shipsOwned(i)).obName
                If backupPlayer.shipSelected = i Then
                    backupPlayer.shipSelectedName = ships(backupPlayer.shipsOwned(i)).obName
                End If
            End If
        End If
    Next i
    ReDim backupPlayer.levelsPassedNames(UBound(backupPlayer.levelsPassed))
    For i = LBound(backupPlayer.levelsPassed) To UBound(backupPlayer.levelsPassed)
        If i <> 0 Then
            backupPlayer.levelsPassedNames(i) = Empty
            If backupPlayer.levelsPassed(i) Then
                backupPlayer.levelsPassedNames(i) = levels(backupPlayer.levelsPassed(i)).obName
            End If
        End If
    Next i
    
    Dim xml_document As DOMDocument
    Set xml_document = New DOMDocument
    Dim node As IXMLDOMNode
    Dim node2 As IXMLDOMNode
    Dim node3 As IXMLDOMNode
    Dim node4 As IXMLDOMNode
    Dim attrib As IXMLDOMAttribute
    
    'Set node = xml_document.createProcessingInstruction("xml", "version='1.0'")
    'xml_document.appendChild node
    
    'figure out how to do doctype
    
    Set node = xml_document.createElement("shipsave")
        Set attrib = xml_document.createAttribute("modname")
            attrib.Text = modName
        node.Attributes.setNamedItem attrib
        Set attrib = xml_document.createAttribute("shipsaveversion")
            attrib.Text = saveVersion
        node.Attributes.setNamedItem attrib
        
        Set node2 = xml_document.createElement("shipsOwned")
            'change this to shipsOwnedNames
            For i = LBound(backupPlayer.shipsOwnedNames) To UBound(backupPlayer.shipsOwnedNames)
                If i <> 0 Then
                    Set node3 = xml_document.createElement("ship")
                        Set node4 = xml_document.createElement("obName")
                            node4.Text = backupPlayer.shipsOwnedNames(i)
                        node3.appendChild node4
                        Set node4 = xml_document.createElement("health")
                            node4.Text = backupPlayer.shipHealth(i)
                        node3.appendChild node4
                    node2.appendChild node3
                End If
            Next i
        node.appendChild node2
    
        Set node2 = xml_document.createElement("shipSelected")
            node2.Text = backupPlayer.shipSelectedName
        node.appendChild node2
        
        Set node2 = xml_document.createElement("weaponsOwned")
            For i = LBound(backupPlayer.weaponsOwnedNames) To UBound(backupPlayer.weaponsOwnedNames)
                If i <> 0 Then
                    Set node3 = xml_document.createElement("weapon")
                        node3.Text = backupPlayer.weaponsOwnedNames(i)
                    node2.appendChild node3
                End If
            Next i
        node.appendChild node2
    
        Set node2 = xml_document.createElement("weaponSelected1")
            node2.Text = backupPlayer.weapon1Name
        node.appendChild node2
        
        Set node2 = xml_document.createElement("weaponSelected2")
            node2.Text = backupPlayer.weapon2Name
        node.appendChild node2
        
        Set node2 = xml_document.createElement("levelsPassed")
            For i = LBound(backupPlayer.levelsPassedNames) To UBound(backupPlayer.levelsPassedNames)
                If i <> 0 Then
                    Set node3 = xml_document.createElement("level")
                        node3.Text = backupPlayer.levelsPassedNames(i)
                    node2.appendChild node3
                End If
            Next i
        node.appendChild node2
    
        Set node2 = xml_document.createElement("cash")
            node2.Text = backupPlayer.cash
        node.appendChild node2
    
    xml_document.appendChild node
    
    xml_document.save savedGamePath
    
    'write xml version and dtd refernce lines
    Dim sfn As Integer
    sfn = FreeFile
    Open savedGamePath For Output As #sfn
    Print #sfn, "<?xml version='1.0'?>" & vbNewLine & "<!DOCTYPE shipsave SYSTEM 'shipsave.dtd'>" & vbNewLine & xml_document.xml
    Close #sfn
End Sub

Sub LoadShipSave()
    If gameLoaded = True Then
        If MsgBox("Are you sure you want to load a game?", vbOKCancel) = vbCancel Then
            Exit Sub
        End If
    End If

    'load into both player then backup right away
    dlgMain.fileName = ""
    dlgMain.DialogTitle = "Load Saved Game"
    'dlgMain.InitDir = App.Path
    dlgMain.ShowOpen
    If Len(dlgMain.fileName) = 0 Then
        Exit Sub
    End If

    Dim xml_document As DOMDocument
    Dim root As IXMLDOMNode
    Set xml_document = New DOMDocument
    
    savedGamePath = dlgMain.fileName
    FileCopy dlgMain.fileName, App.Path & "\shiptemp.xml"
    xml_document.Load App.Path & "\shiptemp.xml"
    Kill App.Path & "\shiptemp.xml"

    'check if the file was opened correctly
    If xml_document.documentElement Is Nothing Then
        MsgBox dlgMain.fileName & " is not a valid ship saved game file"
        Exit Sub
    End If
    Dim dtd As IXMLDOMDocumentType
    Set dtd = xml_document.doctype
    If dtd Is Nothing Then
        MsgBox dlgMain.fileName & " does not reference any DTD. It should reference shipsave.dtd"
        Exit Sub
    Else
        If dtd.Name <> "shipsave" Then
            MsgBox dlgMain.fileName & " references the " & dtd.Name & " dtd. It should reference shipsave.dtd"
            Exit Sub
        End If
    End If

    'find the shipmod node
    Set root = xml_document.selectSingleNode("shipsave")
    
    'load shipmod name
    Dim nameAttribute As IXMLDOMAttribute
    Set nameAttribute = root.Attributes.getNamedItem("modname")
    If nameAttribute.Text <> modName Then
        MsgBox dlgMain.fileName & " is a saved game for the " & nameAttribute.Text & " ship mod. The mod you have loaded currently is " & modName & ". Click 'Switch Mod' to view a list of available mods."
        Exit Sub
    End If
    Set nameAttribute = Nothing
    
    'load shipsave version
    Dim versionAttribute As IXMLDOMAttribute
    Set versionAttribute = root.Attributes.getNamedItem("shipsaveversion")
    If versionAttribute.Text <> saveVersion Then
        MsgBox dlgMain.fileName & " is a version " & versionAttribute.Text & " ship save file. This version of ship loads version " & saveVersion & " ship saved game files only."
        Exit Sub
    End If
    Set versionAttribute = Nothing
    
    gameLoaded = False
    Update
    
    Dim node1 As IXMLDOMElement
    Dim node2 As IXMLDOMElement
    Dim nodelist1 As IXMLDOMNodeList
    Dim i As Integer
    Dim e As Integer
    
    Set node1 = root.selectSingleNode("shipsOwned")
        Set nodelist1 = node1.selectNodes("ship")
            i = 1
            ReDim player.shipsOwnedNames(nodelist1.length)
            ReDim player.shipHealth(nodelist1.length)
            While i <= nodelist1.length
                Set node2 = nodelist1.Item(i - 1).selectSingleNode("obName")
                    player.shipsOwnedNames(i) = node2.Text
                Set node2 = nodelist1.Item(i - 1).selectSingleNode("health")
                    player.shipHealth(i) = node2.Text
                i = i + 1
            Wend
                    
    Set node1 = root.selectSingleNode("shipSelected")
        player.shipSelectedName = node1.Text
        
    Set node1 = root.selectSingleNode("weaponsOwned")
        Set nodelist1 = node1.selectNodes("weapon")
            i = 1
            ReDim player.weaponsOwnedNames(nodelist1.length)
            While i <= nodelist1.length
                player.weaponsOwnedNames(i) = nodelist1.Item(i - 1).Text
                i = i + 1
            Wend
            
    Set node1 = root.selectSingleNode("weaponSelected1")
        player.weapon1Name = node1.Text
        
    Set node1 = root.selectSingleNode("weaponSelected2")
        player.weapon2Name = node1.Text
        
    Set node1 = root.selectSingleNode("levelsPassed")
        Set nodelist1 = node1.selectNodes("level")
            i = 1
            ReDim player.levelsPassedNames(nodelist1.length)
            While i <= nodelist1.length
                player.levelsPassedNames(i) = nodelist1.Item(i - 1).Text
                i = i + 1
            Wend
    
    Set node1 = root.selectSingleNode("cash")
        player.cash = node1.Text
        
    player.weapon1 = -1
    player.weapon2 = -1
    ReDim player.weaponsOwned(UBound(player.weaponsOwnedNames))
    For i = LBound(player.weaponsOwnedNames) To UBound(player.weaponsOwnedNames)
        If i <> 0 Then
            player.weaponsOwned(i) = -1
            For e = LBound(weapons) To UBound(weapons)
                If e <> 0 Then
                    If player.weaponsOwnedNames(i) = weapons(e).obName Then
                        player.weaponsOwned(i) = e
                    End If
                End If
            Next e
            If player.weaponsOwned(i) = -1 Then
                MsgBox "Error in save file. This file is either corrupt or was intended for use with a different ship mod with the name " & modName
                Exit Sub
            End If
            If player.weapon1Name = player.weaponsOwnedNames(i) Then
                player.weapon1 = i
            End If
            If player.weapon2Name = player.weaponsOwnedNames(i) Then
                player.weapon2 = i
            End If
        End If
    Next i
    player.shipSelected = -1
    ReDim player.shipsOwned(UBound(player.shipsOwnedNames))
    For i = LBound(player.shipsOwnedNames) To UBound(player.shipsOwnedNames)
        If i <> 0 Then
            player.shipsOwned(i) = -1
            For e = LBound(ships) To UBound(ships)
                If e <> 0 Then
                    If player.shipsOwnedNames(i) = ships(e).obName Then
                        player.shipsOwned(i) = e
                    End If
                End If
            Next e
            If player.shipsOwned(i) = -1 Then
                MsgBox "Error in save file. This file is either corrupt or was intended for use with a different ship mod with the name " & modName
                Exit Sub
            End If
            If player.shipSelectedName = player.shipsOwnedNames(i) Then
                player.shipSelected = i
            End If
        End If
    Next i
    ReDim player.levelsPassed(UBound(player.levelsPassedNames))
    For i = LBound(player.levelsPassedNames) To UBound(player.levelsPassedNames)
        If i <> 0 Then
            player.levelsPassed(i) = -1
            For e = LBound(levels) To UBound(levels)
                If e <> 0 Then
                    If player.levelsPassedNames(i) = levels(e).obName Then
                        player.levelsPassed(i) = e
                    End If
                End If
            Next e
            If player.levelsPassed(i) = -1 Then
                MsgBox "Error in save file. This file is either corrupt or was intended for use with a different ship mod with the name " & modName
                Exit Sub
            End If
        End If
    Next i
    
    backupPlayer = player
    
    gameLoaded = True
    
    Update
End Sub

Private Sub cmdExit_Click()
    If gameLoaded = True Then
        If MsgBox("Are you sure u want to quit?", vbOKCancel) = vbCancel Then
            Exit Sub
        End If
    End If
    End
End Sub

Private Sub cmdLoad_Click()
    LoadShipSave
End Sub

Private Sub cmdMod_Click()
    Me.Hide
    frmMods.Show
End Sub

Private Sub cmdNewGame_Click()
    newgame
End Sub

Private Sub cmdSave_Click()
    WriteShipSave
End Sub

Private Sub cmdStartMission_Click()
    StartMission
End Sub

Private Sub Form_Load()
    On Error GoTo fileerror
    Open App.Path & "\defaultmod" For Input As #1
    Line Input #1, defaultModFileName
    Close #1
    modFileName = defaultModFileName
    LoadMod
    
fileerror:
    Update
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub

Private Sub lstLevels_Click()
    player.level = lstLevels.ItemData(lstLevels.ListIndex)
    Update
End Sub
