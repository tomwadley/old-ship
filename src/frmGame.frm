VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGame 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "basecaption - modname - levelname"
   ClientHeight    =   8205
   ClientLeft      =   285
   ClientTop       =   450
   ClientWidth     =   7575
   FillColor       =   &H00808080&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar pbHealth 
      Height          =   7935
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   13996
      _Version        =   393216
      Appearance      =   0
      Max             =   1
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.Label lblShipLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ship"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   13
      Top             =   7875
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblShip 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Name of ship"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1080
      TabIndex        =   12
      Top             =   7875
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblWeapon 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Name of weapon"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4800
      TabIndex        =   11
      Top             =   7875
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblWeaponLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Weapon"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   7875
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblHealthLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Shield"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblHealth 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblFpsLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "FPS"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblFps 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   6600
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPaused 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Paused"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblMissionFailed 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mission Failed.."
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblMissionPassed 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mission Passed!"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblCashLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cash"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblCash 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "$ 000000000000000"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const borderLeft = 450
Const borderTop = 300
Const borderRight = 0
Const borderBottom = 350

Const baseFormCaption = "ship 0.4"

Dim fps As Integer
Dim progressiveFps As Integer
Dim etThisSecond As Double

Dim fpsShowing As Boolean
Dim pausedShowing As Boolean
Dim failedShowing As Boolean
Dim passedShowing As Boolean

'Const numberOfPics = 27 'starts at 0
Dim imglstArray() As MSComctlLib.ImageList

Const gameEndTime = 4

Dim playerCollisionTimer As Double
Dim nextBackImageTop As Integer 'top for next back image, -1 for not needed yet

Dim windowHeight As Integer
Dim windowWidth As Integer

'level stuff
Dim missionOver As Boolean
Dim gameEndTimer As Double
Dim timeSoFar As Double
Dim goalReached As Boolean
Dim bossDead As Boolean
Dim enemiesStartTimer As Double
Dim levelIndex As Integer
''''''''''''

Private Type typeRuntimeBackgroundobject
    left As Integer
    top As Integer
    pic As Integer
    picTimer As Double
    backgoundobjectTemplate As Integer 'index into backgroundobjects
    backImage As Boolean
    backImageTop As Boolean
End Type

Private Type typeRuntimeEnemy
    left As Integer
    top As Integer
    moveLeft As Integer
    moveTop As Integer
    health As Integer
    enemyTemplate As Integer 'index into enemies
    reloadTimer() As Double
    pic As Integer
    picTimer As Double
    alive As Boolean
    boss As Boolean 'is this enemy the level boss
End Type

Private Type typeRuntimeProjectile
    left As Integer
    top As Integer
    moveLeft As Integer 'these are here because of tracking
    moveTop As Integer
    weaponFiredFrom As Integer 'index into weapons
    projectileFromWeapon As Integer 'index into weapons.projectiles
    rtEnemyFiredFrom As Integer 'index into rtEnemies, -1 for enemy no longer exists or owned by player
    pic As Integer
    picTimer As Double
    alive As Boolean
    ownedByPlayer As Boolean
End Type

Dim rtBackgroundobjects() As typeRuntimeBackgroundobject
Dim rtEnemies() As typeRuntimeEnemy
Dim projectiles() As typeRuntimeProjectile

Dim keyLeftDown As Boolean
Dim keyRightDown As Boolean
Dim keyUpDown As Boolean
Dim keyDownDown As Boolean
Dim keyFireDown As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyLeft
            keyLeftDown = True
        Case vbKeyA
            keyLeftDown = True
        Case vbKeyRight
            keyRightDown = True
        Case vbKeyD
            keyRightDown = True
        Case vbKeyUp
            keyUpDown = True
        Case vbKeyW
            keyUpDown = True
        Case vbKeyDown
            keyDownDown = True
        Case vbKeyS
            keyDownDown = True
        Case vbKeySpace
            keyFireDown = True
        Case vbKeyControl
            SwitchWeapon
        Case vbKeyReturn
            PauseGame
        Case vbKeyBack
            ToggleFps
        Case vbKeyEscape
            PauseGame
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyLeft
            keyLeftDown = False
        Case vbKeyA
            keyLeftDown = False
        Case vbKeyRight
            keyRightDown = False
        Case vbKeyD
            keyRightDown = False
        Case vbKeyUp
            keyUpDown = False
        Case vbKeyW
            keyUpDown = False
        Case vbKeyDown
            keyDownDown = False
        Case vbKeyS
            keyDownDown = False
        Case vbKeySpace
            keyFireDown = False
    End Select
End Sub

Private Sub Form_Load()
    ReDim imglstArray(UBound(pics))
    Dim i As Integer
    Dim e As Integer
    Dim f As Integer
    For i = 0 To UBound(levels(player.level).backgroundobjectsPresent)
        If i <> 0 Then
            For e = 0 To UBound(backgroundobjects(levels(player.level).backgroundobjectsPresent(i)).pics)
                If e <> 0 Then
                    LoadPic backgroundobjects(levels(player.level).backgroundobjectsPresent(i)).pics(e)
                End If
            Next e
        End If
    Next i
    If levels(player.level).backImageBackgroundobject > 0 Then
        For e = 0 To UBound(backgroundobjects(levels(player.level).backImageBackgroundobject).pics)
            If e <> 0 Then
                LoadPic backgroundobjects(levels(player.level).backImageBackgroundobject).pics(e)
            End If
        Next e
    End If
    If levels(player.level).boss > 0 Then
        For e = 0 To UBound(enemies(levels(player.level).boss).pics)
            If e <> 0 Then
                LoadPic enemies(levels(player.level).boss).pics(e)
            End If
        Next e
        For e = 0 To UBound(enemies(levels(player.level).boss).picsDead)
            If e <> 0 Then
                LoadPic enemies(levels(player.level).boss).picsDead(e)
            End If
        Next e
        If enemies(levels(player.level).boss).weapon > 0 Then
            For i = 0 To UBound(weapons(enemies(levels(player.level).boss).weapon).projectiles)
                If i <> 0 Then
                    For e = 0 To UBound(weapons(enemies(levels(player.level).boss).weapon).projectiles(i).pics)
                        If e <> 0 Then
                            LoadPic weapons(enemies(levels(player.level).boss).weapon).projectiles(i).pics(e)
                        End If
                    Next e
                    For e = 0 To UBound(weapons(enemies(levels(player.level).boss).weapon).projectiles(i).picsDead)
                        If e <> 0 Then
                            LoadPic weapons(enemies(levels(player.level).boss).weapon).projectiles(i).picsDead(e)
                        End If
                    Next e
                End If
            Next i
        End If
    End If
    For i = 0 To UBound(levels(player.level).enemiesPresent)
        If i <> 0 Then
            For e = 0 To UBound(enemies(levels(player.level).enemiesPresent(i).enemy).pics)
                If e <> 0 Then
                    LoadPic enemies(levels(player.level).enemiesPresent(i).enemy).pics(e)
                End If
            Next e
            For e = 0 To UBound(enemies(levels(player.level).enemiesPresent(i).enemy).picsDead)
                If e <> 0 Then
                    LoadPic enemies(levels(player.level).enemiesPresent(i).enemy).picsDead(e)
                End If
            Next e
            If enemies(levels(player.level).enemiesPresent(i).enemy).weapon > 0 Then
                For e = 0 To UBound(weapons(enemies(levels(player.level).enemiesPresent(i).enemy).weapon).projectiles)
                    If e <> 0 Then
                        For f = 0 To UBound(weapons(enemies(levels(player.level).enemiesPresent(i).enemy).weapon).projectiles(e).pics)
                            If f <> 0 Then
                                LoadPic weapons(enemies(levels(player.level).enemiesPresent(i).enemy).weapon).projectiles(e).pics(f)
                            End If
                        Next f
                        For f = 0 To UBound(weapons(enemies(levels(player.level).enemiesPresent(i).enemy).weapon).projectiles(e).picsDead)
                            If f <> 0 Then
                                LoadPic weapons(enemies(levels(player.level).enemiesPresent(i).enemy).weapon).projectiles(e).picsDead(f)
                            End If
                        Next f
                    End If
                Next e
            End If
        End If
    Next i
    For e = 0 To UBound(ships(player.shipsOwned(player.shipSelected)).pics)
        If e <> 0 Then
            LoadPic ships(player.shipsOwned(player.shipSelected)).pics(e)
        End If
    Next e
    For e = 0 To UBound(ships(player.shipsOwned(player.shipSelected)).picsDead)
        If e <> 0 Then
            LoadPic ships(player.shipsOwned(player.shipSelected)).picsDead(e)
        End If
    Next e
    If ships(player.shipsOwned(player.shipSelected)).collisionWeapon > 0 Then
        For e = 0 To UBound(weapons(ships(player.shipsOwned(player.shipSelected)).collisionWeapon).projectiles(ships(player.shipsOwned(player.shipSelected)).collisionProjectileFromWeapon).pics)
            If e <> 0 Then
                LoadPic weapons(ships(player.shipsOwned(player.shipSelected)).collisionWeapon).projectiles(ships(player.shipsOwned(player.shipSelected)).collisionProjectileFromWeapon).pics(e)
            End If
        Next e
    End If
    If ships(player.shipsOwned(player.shipSelected)).collisionProjectileFromWeapon > 0 Then
        For e = 0 To UBound(weapons(ships(player.shipsOwned(player.shipSelected)).collisionWeapon).projectiles(ships(player.shipsOwned(player.shipSelected)).collisionProjectileFromWeapon).picsDead)
            If e <> 0 Then
                LoadPic weapons(ships(player.shipsOwned(player.shipSelected)).collisionWeapon).projectiles(ships(player.shipsOwned(player.shipSelected)).collisionProjectileFromWeapon).picsDead(e)
            End If
        Next e
    End If
    If player.weapon1 > 0 Then
        For i = 0 To UBound(weapons(player.weaponsOwned(player.weapon1)).projectiles)
            If i <> 0 Then
                For e = 0 To UBound(weapons(player.weaponsOwned(player.weapon1)).projectiles(i).pics)
                    If e <> 0 Then
                        LoadPic weapons(player.weaponsOwned(player.weapon1)).projectiles(i).pics(e)
                    End If
                Next e
                For e = 0 To UBound(weapons(player.weaponsOwned(player.weapon1)).projectiles(i).picsDead)
                    If e <> 0 Then
                        LoadPic weapons(player.weaponsOwned(player.weapon1)).projectiles(i).picsDead(e)
                    End If
                Next e
            End If
        Next i
    End If
    If player.weapon2 > 0 Then
        For i = 0 To UBound(weapons(player.weaponsOwned(player.weapon2)).projectiles)
            If i <> 0 Then
                For e = 0 To UBound(weapons(player.weaponsOwned(player.weapon2)).projectiles(i).pics)
                    If e <> 0 Then
                        LoadPic weapons(player.weaponsOwned(player.weapon2)).projectiles(i).pics(e)
                    End If
                Next e
                For e = 0 To UBound(weapons(player.weaponsOwned(player.weapon2)).projectiles(i).picsDead)
                    If e <> 0 Then
                        LoadPic weapons(player.weaponsOwned(player.weapon2)).projectiles(i).picsDead(e)
                    End If
                Next e
            End If
        Next i
    End If
End Sub

Sub LoadPic(pic As Integer)
    If imglstArray(pic) Is Nothing Then
        Set imglstArray(pic) = Controls.Add("MSComctlLib.ImageListCtrl.2", "imglstArray" & pic)
        imglstArray(pic).ListImages.Clear
        imglstArray(pic).ListImages.Add 1, , LoadPicture(App.Path & "\" & modName & "\" & pics(pic).filePath)
        imglstArray(pic).MaskColor = picMaskColour
        imglstArray(pic).UseMaskColor = True
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    player.alive = False
    EndGame
End Sub

Sub SwitchWeapon()
    Dim oldWeapon As Integer
    oldWeapon = player.weaponSelected
    If player.weaponSelected = player.weapon1 Then
        player.weaponSelected = player.weapon2
    Else
        player.weaponSelected = player.weapon1
    End If
    If player.weaponSelected <= 0 Then
        player.weaponSelected = oldWeapon
    End If
    Dim i As Integer
    ReDim player.reloadTimer(UBound(weapons(player.weaponsOwned(player.weaponSelected)).projectiles))
    For i = LBound(player.reloadTimer) To UBound(player.reloadTimer)
        If i <> 0 Then
            player.reloadTimer(i) = weapons(player.weaponsOwned(player.weaponSelected)).projectiles(i).reloadTime
        End If
    Next i
End Sub

Sub ToggleFps()
    If fpsShowing = False Then
        fpsShowing = True
    Else
        fpsShowing = False
    End If
End Sub

Sub FillScreenWithBackgroundobjects()
    Dim i As Integer
    For i = LBound(levels(levelIndex).backgroundobjectsPresent) To UBound(levels(levelIndex).backgroundobjectsPresent)
        If i <> 0 Then
            If levels(levelIndex).backgroundobjectsPresent(i) >= 0 Then
                If backgroundobjects(i).moveSpeed <> 0 Then
                    Dim e As Integer
                    For e = 1 To (windowHeight / backgroundobjects(i).moveSpeed) * backgroundobjects(i).chanceOfCreation
                        CreateBackgroundobject levels(levelIndex).backgroundobjectsPresent(i), True
                    Next e
                End If
            End If
        End If
    Next i
    
    If levels(levelIndex).backImageBackgroundobject >= 0 Then
        Dim firstOneCreated As Integer
        nextBackImageTop = windowWidth '- imgMain(backgroundobjects(levels(levelIndex).backImageBackgroundobject).pics(0)).height
        While nextBackImageTop > 0
            Dim leftUpto As Integer
            leftUpto = 0
            firstOneCreated = UBound(rtBackgroundobjects) + 1
            While leftUpto <= windowWidth
                ReDim Preserve rtBackgroundobjects(UBound(rtBackgroundobjects) + 1)
                rtBackgroundobjects(UBound(rtBackgroundobjects)).backgoundobjectTemplate = levels(levelIndex).backImageBackgroundobject
                rtBackgroundobjects(UBound(rtBackgroundobjects)).pic = 1
                rtBackgroundobjects(UBound(rtBackgroundobjects)).picTimer = backgroundobjects(levels(levelIndex).backImageBackgroundobject).timeBetweenPics
                rtBackgroundobjects(UBound(rtBackgroundobjects)).top = nextBackImageTop
                rtBackgroundobjects(UBound(rtBackgroundobjects)).left = leftUpto
                leftUpto = leftUpto + pics(backgroundobjects(levels(levelIndex).backImageBackgroundobject).pics(1)).width
                rtBackgroundobjects(UBound(rtBackgroundobjects)).backImage = True
                rtBackgroundobjects(UBound(rtBackgroundobjects)).backImageTop = False
            Wend
            nextBackImageTop = nextBackImageTop - pics(backgroundobjects(levels(levelIndex).backImageBackgroundobject).pics(1)).height
        Wend
        rtBackgroundobjects(firstOneCreated).backImageTop = True
        nextBackImageTop = 1
    End If
End Sub

Sub CreateBackgroundobject(backgroundobject As Integer, randomTop As Boolean) 'this cannot create background images
    Dim objectLeft As Integer
    Dim objectTop As Integer
    If randomTop = True Then
        objectTop = Int(Rnd() * (windowHeight - pics(backgroundobjects(backgroundobject).pics(1)).height))
    Else
        objectTop = -pics(backgroundobjects(backgroundobject).pics(1)).height
    End If
    objectLeft = Int(Rnd() * (windowWidth - pics(backgroundobjects(backgroundobject).pics(1)).width))
    Dim proceed As Boolean
    proceed = True
    Dim e As Integer
    For e = LBound(rtBackgroundobjects) To UBound(rtBackgroundobjects)
        If e <> 0 Then
            If rtBackgroundobjects(e).backImage = False Then
                If objectLeft + pics(backgroundobjects(backgroundobject).pics(1)).width > rtBackgroundobjects(e).left And objectLeft < rtBackgroundobjects(e).left + pics(backgroundobjects(rtBackgroundobjects(e).backgoundobjectTemplate).pics(rtBackgroundobjects(e).pic)).width Then
                    If objectTop + pics(backgroundobjects(backgroundobject).pics(1)).height > rtBackgroundobjects(e).top And objectTop < rtBackgroundobjects(e).top + pics(backgroundobjects(rtBackgroundobjects(e).backgoundobjectTemplate).pics(rtBackgroundobjects(e).pic)).height Then
                        proceed = False
                    End If
                End If
            End If
        End If
    Next e
    If proceed = True Then
        ReDim Preserve rtBackgroundobjects(UBound(rtBackgroundobjects) + 1)
        rtBackgroundobjects(UBound(rtBackgroundobjects)).backgoundobjectTemplate = backgroundobject
        rtBackgroundobjects(UBound(rtBackgroundobjects)).pic = 1
        rtBackgroundobjects(UBound(rtBackgroundobjects)).picTimer = backgroundobjects(backgroundobject).timeBetweenPics
        rtBackgroundobjects(UBound(rtBackgroundobjects)).top = objectTop
        rtBackgroundobjects(UBound(rtBackgroundobjects)).left = objectLeft
        rtBackgroundobjects(UBound(rtBackgroundobjects)).backImage = False
        rtBackgroundobjects(UBound(rtBackgroundobjects)).backImageTop = False
    End If
End Sub

Sub CreateEnemy(enemy As Integer, boss As Boolean)
    If UBound(enemies(enemy).pics) = 0 Then
        'this enemy has no pics and therefore will not be created
        Exit Sub
    End If
    Dim e As Integer
    ReDim Preserve rtEnemies(UBound(rtEnemies) + 1)
    rtEnemies(UBound(rtEnemies)).enemyTemplate = enemy
    rtEnemies(UBound(rtEnemies)).health = enemies(rtEnemies(UBound(rtEnemies)).enemyTemplate).maxHealth
    rtEnemies(UBound(rtEnemies)).pic = 1
    rtEnemies(UBound(rtEnemies)).picTimer = enemies(rtEnemies(UBound(rtEnemies)).enemyTemplate).timeBetweenPics
    rtEnemies(UBound(rtEnemies)).moveTop = Int(Rnd() * (enemies(rtEnemies(UBound(rtEnemies)).enemyTemplate).maxMoveTop - enemies(rtEnemies(UBound(rtEnemies)).enemyTemplate).minMoveTop)) + enemies(rtEnemies(UBound(rtEnemies)).enemyTemplate).minMoveTop
    rtEnemies(UBound(rtEnemies)).moveLeft = Int(Rnd() * (enemies(rtEnemies(UBound(rtEnemies)).enemyTemplate).maxMoveLeft - enemies(rtEnemies(UBound(rtEnemies)).enemyTemplate).minMoveLeft)) + enemies(rtEnemies(UBound(rtEnemies)).enemyTemplate).minMoveLeft
    rtEnemies(UBound(rtEnemies)).alive = True
    rtEnemies(UBound(rtEnemies)).boss = boss
    If enemies(rtEnemies(UBound(rtEnemies)).enemyTemplate).weapon <> -1 Then
        ReDim rtEnemies(UBound(rtEnemies)).reloadTimer(UBound(weapons(enemies(rtEnemies(UBound(rtEnemies)).enemyTemplate).weapon).projectiles))
        For e = LBound(rtEnemies(UBound(rtEnemies)).reloadTimer) To UBound(rtEnemies(UBound(rtEnemies)).reloadTimer)
            If e <> 0 Then
                rtEnemies(UBound(rtEnemies)).reloadTimer(e) = enemies(rtEnemies(UBound(rtEnemies)).enemyTemplate).initReloadTime + weapons(enemies(rtEnemies(UBound(rtEnemies)).enemyTemplate).weapon).projectiles(e).initReloadTime
            End If
        Next e
    End If
    Dim rndTop As Integer
    Dim rndLeft As Integer
    Select Case enemies(rtEnemies(UBound(rtEnemies)).enemyTemplate).entrySide
        Case epAnywhereSide
            rndTop = Int(Rnd() * (windowHeight - pics(enemies(rtEnemies(UBound(rtEnemies)).enemyTemplate).pics(1)).height))
            rndLeft = Int(Rnd() * (windowWidth - pics(enemies(rtEnemies(UBound(rtEnemies)).enemyTemplate).pics(1)).width))
        Case epMiddleSide
            rndTop = (((windowWidth) / 2)) + Int(Rnd() * ((enemies(rtEnemies(UBound(rtEnemies)).enemyTemplate).entrySideChange * windowHeight) * 2) - (enemies(rtEnemies(UBound(rtEnemies)).enemyTemplate).entrySideChange * windowHeight) - (pics(enemies(rtEnemies(UBound(rtEnemies)).enemyTemplate).pics(rtEnemies(UBound(rtEnemies)).pic)).height / 2))
            rndLeft = (((windowWidth) / 2)) + Int(Rnd() * ((enemies(rtEnemies(UBound(rtEnemies)).enemyTemplate).entrySideChange * windowHeight) * 2) - (enemies(rtEnemies(UBound(rtEnemies)).enemyTemplate).entrySideChange * windowHeight) - (pics(enemies(rtEnemies(UBound(rtEnemies)).enemyTemplate).pics(rtEnemies(UBound(rtEnemies)).pic)).width / 2))
        Case epLeftSide
            rndTop = Int(Rnd() * ((enemies(rtEnemies(UBound(rtEnemies)).enemyTemplate).entrySideChange * windowHeight)))
            rndLeft = Int(Rnd() * ((enemies(rtEnemies(UBound(rtEnemies)).enemyTemplate).entrySideChange * windowHeight)))
        Case epRightSide
            rndTop = windowHeight - Int(Rnd() * ((enemies(rtEnemies(UBound(rtEnemies)).enemyTemplate).entrySideChange * windowHeight)))
            rndLeft = windowWidth - Int(Rnd() * ((enemies(rtEnemies(UBound(rtEnemies)).enemyTemplate).entrySideChange * windowHeight)))
    End Select
    Select Case enemies(rtEnemies(UBound(rtEnemies)).enemyTemplate).entryPoint
        Case epTop
            rtEnemies(UBound(rtEnemies)).top = -pics(enemies(rtEnemies(UBound(rtEnemies)).enemyTemplate).pics(1)).height
            rtEnemies(UBound(rtEnemies)).left = rndLeft
        Case epBottom
            rtEnemies(UBound(rtEnemies)).top = windowHeight
            rtEnemies(UBound(rtEnemies)).left = rndLeft
        Case epLeft
            rtEnemies(UBound(rtEnemies)).top = rndTop
            rtEnemies(UBound(rtEnemies)).left = -pics(enemies(rtEnemies(UBound(rtEnemies)).enemyTemplate).pics(1)).width
        Case epRight
            rtEnemies(UBound(rtEnemies)).top = rndTop
            rtEnemies(UBound(rtEnemies)).left = windowWidth
    End Select
End Sub

Sub CreateProjectile(weapon As Integer, projectile As Integer, ownedByPlayer As Boolean, left As Integer, top As Integer, height As Integer, width As Integer, rtenemy As Integer)
    ReDim Preserve projectiles(UBound(projectiles) + 1)
    projectiles(UBound(projectiles)).weaponFiredFrom = weapon
    projectiles(UBound(projectiles)).projectileFromWeapon = projectile
    If ownedByPlayer = True Then
        projectiles(UBound(projectiles)).rtEnemyFiredFrom = -1 'perhaps make this 0
    Else
        projectiles(UBound(projectiles)).rtEnemyFiredFrom = rtenemy
    End If
    projectiles(UBound(projectiles)).pic = 1
    projectiles(UBound(projectiles)).picTimer = weapons(weapon).projectiles(projectile).timeBetweenPics
    Select Case weapons(weapon).projectiles(projectile).column
        Case wpFarLeft
            projectiles(UBound(projectiles)).left = left - pics(weapons(weapon).projectiles(projectile).pics(1)).width
        Case wpLeft
            projectiles(UBound(projectiles)).left = left
        Case wpRight
            projectiles(UBound(projectiles)).left = left + width - pics(weapons(weapon).projectiles(projectile).pics(1)).width
        Case wpFarRight
            projectiles(UBound(projectiles)).left = left + width
        Case wpCustom
            projectiles(UBound(projectiles)).left = left + (weapons(weapon).projectiles(projectile).customColumn * width) - (pics(weapons(weapon).projectiles(projectile).pics(1)).width / 2)
    End Select
    Select Case weapons(weapon).projectiles(projectile).row
        Case wpFarFront
            projectiles(UBound(projectiles)).top = top - pics(weapons(weapon).projectiles(projectile).pics(1)).height
        Case wpFront
            projectiles(UBound(projectiles)).top = top
        Case wpBack
            projectiles(UBound(projectiles)).top = top + height - pics(weapons(weapon).projectiles(projectile).pics(1)).height
        Case wpFarBack
            projectiles(UBound(projectiles)).top = top + height
        Case wpCustom
            projectiles(UBound(projectiles)).top = top + (weapons(weapon).projectiles(projectile).customRow * height) - (pics(weapons(weapon).projectiles(projectile).pics(1)).height / 2)
    End Select
    If weapons(weapon).projectiles(projectile).tracksPlayer = True And ownedByPlayer = False Then
        If player.alive = True Then
            'track the player
            Dim difX As Integer
            Dim difY As Integer
            difX = player.left + (pics(ships(player.shipsOwned(player.shipSelected)).pics(player.pic)).width / 2) - projectiles(UBound(projectiles)).left + (pics(weapons(projectiles(UBound(projectiles)).weaponFiredFrom).projectiles(projectiles(UBound(projectiles)).projectileFromWeapon).pics(projectiles(UBound(projectiles)).pic)).width / 2)
            difY = player.top + (pics(ships(player.shipsOwned(player.shipSelected)).pics(player.pic)).height / 2) - projectiles(UBound(projectiles)).top + (pics(weapons(projectiles(UBound(projectiles)).weaponFiredFrom).projectiles(projectiles(UBound(projectiles)).projectileFromWeapon).pics(projectiles(UBound(projectiles)).pic)).height / 2)
            Dim totalSpeed As Integer
            totalSpeed = weapons(weapon).projectiles(projectile).moveTop
            projectiles(UBound(projectiles)).moveLeft = (difX / (Abs(difX) + Abs(difY))) * totalSpeed
            projectiles(UBound(projectiles)).moveTop = (difY / (Abs(difX) + Abs(difY))) * totalSpeed
        Else
            removeProjectile UBound(projectiles)
            Exit Sub
        End If
    Else
        projectiles(UBound(projectiles)).moveLeft = weapons(weapon).projectiles(projectile).moveLeft
        projectiles(UBound(projectiles)).moveTop = weapons(weapon).projectiles(projectile).moveTop
    End If
    projectiles(UBound(projectiles)).alive = True
    projectiles(UBound(projectiles)).ownedByPlayer = ownedByPlayer
End Sub

Sub EnemyKilled(enemy As Integer, byPlayer As Boolean)
    rtEnemies(enemy).alive = False
    If rtEnemies(enemy).boss = True Then
        bossDead = True
    End If
    
    Dim leftCenter As Integer
    leftCenter = rtEnemies(enemy).left + (pics(enemies(rtEnemies(enemy).enemyTemplate).pics(rtEnemies(enemy).pic)).width / 2)
    Dim topCenter As Integer
    topCenter = rtEnemies(enemy).top + (pics(enemies(rtEnemies(enemy).enemyTemplate).pics(rtEnemies(enemy).pic)).height / 2)
    
    rtEnemies(enemy).pic = 1
    rtEnemies(enemy).picTimer = enemies(rtEnemies(enemy).enemyTemplate).timeBetweenPicsDead
    
    rtEnemies(enemy).left = leftCenter - (pics(enemies(rtEnemies(enemy).enemyTemplate).picsDead(rtEnemies(enemy).pic)).width / 2)
    rtEnemies(enemy).top = topCenter - (pics(enemies(rtEnemies(enemy).enemyTemplate).picsDead(rtEnemies(enemy).pic)).height / 2)
    
    If byPlayer = True Then
        player.cash = player.cash + enemies(rtEnemies(enemy).enemyTemplate).cash
    End If
End Sub

Sub PlayerKilled()
    player.alive = False
    
    Dim leftCenter As Integer
    leftCenter = player.left + (pics(ships(player.shipsOwned(player.shipSelected)).pics(player.pic)).width / 2)
    Dim topCenter As Integer
    topCenter = player.top + (pics(ships(player.shipsOwned(player.shipSelected)).pics(player.pic)).height / 2)
    
    player.pic = 1
    player.picTimer = ships(player.shipsOwned(player.shipSelected)).timeBetweenPicsDead
    
    player.left = leftCenter - (pics(ships(player.shipsOwned(player.shipSelected)).picsDead(player.pic)).width / 2)
    player.top = topCenter - (pics(ships(player.shipsOwned(player.shipSelected)).picsDead(player.pic)).height / 2)
End Sub

Sub DrawScreen()
    'fps
    progressiveFps = progressiveFps + 1
    etThisSecond = etThisSecond + et
    If etThisSecond >= 1 Then
        etThisSecond = etThisSecond - 1
        fps = progressiveFps
        progressiveFps = 0
    End If
    'set labels
    If player.shipHealth(player.shipSelected) >= 0 Then
        pbHealth.Value = player.shipHealth(player.shipSelected)
    Else
        pbHealth.Value = 0
    End If
    lblHealth.Caption = player.shipHealth(player.shipSelected)
    lblCash.Caption = "$" & player.cash
    lblWeapon.Caption = weapons(player.weaponsOwned(player.weaponSelected)).title
    lblShip.Caption = ships(player.shipsOwned(player.shipSelected)).title
    lblFps.Caption = fps
    If missionOver = True Then
        If player.alive = True Then
            passedShowing = True
            failedShowing = False
        Else
            passedShowing = False
            failedShowing = True
        End If
    Else
        passedShowing = False
        failedShowing = False
    End If
    
    'draw
    Me.Cls
    Dim i As Integer 'for looping
    
    'background objects
    For i = LBound(rtBackgroundobjects) To UBound(rtBackgroundobjects)
        If i > 0 Then
            If rtBackgroundobjects(i).backImage = True Then
                Me.ScaleLeft = 0 - rtBackgroundobjects(i).left - borderLeft
                Me.ScaleTop = 0 - rtBackgroundobjects(i).top - borderTop
                imglstArray(backgroundobjects(rtBackgroundobjects(i).backgoundobjectTemplate).pics(rtBackgroundobjects(i).pic)).ListImages(1).Draw Me.hDC, 0, 0, 1
                Scale
            End If
        End If
    Next i
    For i = LBound(rtBackgroundobjects) To UBound(rtBackgroundobjects)
        If i > 0 Then
            If rtBackgroundobjects(i).backImage = False Then
                Me.ScaleLeft = 0 - rtBackgroundobjects(i).left - borderLeft
                Me.ScaleTop = 0 - rtBackgroundobjects(i).top - borderTop
                imglstArray(backgroundobjects(rtBackgroundobjects(i).backgoundobjectTemplate).pics(rtBackgroundobjects(i).pic)).ListImages(1).Draw Me.hDC, 0, 0, 1
                Scale
            End If
        End If
    Next i
    
    'enemies
    For i = LBound(rtEnemies) To UBound(rtEnemies)
        If i > 0 Then
            Me.ScaleLeft = 0 - rtEnemies(i).left - borderLeft
            Me.ScaleTop = 0 - rtEnemies(i).top - borderTop
            If rtEnemies(i).alive = True Then
                imglstArray(enemies(rtEnemies(i).enemyTemplate).pics(rtEnemies(i).pic)).ListImages(1).Draw Me.hDC, 0, 0, 1
            Else
                imglstArray(enemies(rtEnemies(i).enemyTemplate).picsDead(rtEnemies(i).pic)).ListImages(1).Draw Me.hDC, 0, 0, 1
            End If
            Scale
        End If
    Next i
    
    'player
    If missionOver = True And player.alive = False Then
        'do not draw player
    Else
        Me.ScaleLeft = 0 - player.left - borderLeft
        Me.ScaleTop = 0 - player.top - borderTop
        If player.alive = True Then
            imglstArray(ships(player.shipsOwned(player.shipSelected)).pics(player.pic)).ListImages(1).Draw Me.hDC, 0, 0, 1
        Else
            imglstArray(ships(player.shipsOwned(player.shipSelected)).picsDead(player.pic)).ListImages(1).Draw Me.hDC, 0, 0, 1
        End If
        Scale
    End If
    
    'projectiles
    For i = LBound(projectiles) To UBound(projectiles)
        If i > 0 Then
            Me.ScaleLeft = 0 - projectiles(i).left - borderLeft
            Me.ScaleTop = 0 - projectiles(i).top - borderTop
            If projectiles(i).alive = True Then
                imglstArray(weapons(projectiles(i).weaponFiredFrom).projectiles(projectiles(i).projectileFromWeapon).pics(projectiles(i).pic)).ListImages(1).Draw Me.hDC, 0, 0, 1
            Else
                imglstArray(weapons(projectiles(i).weaponFiredFrom).projectiles(projectiles(i).projectileFromWeapon).picsDead(projectiles(i).pic)).ListImages(1).Draw Me.hDC, 0, 0, 1
            End If
            Scale
        End If
    Next i
    
    'border
    Me.ForeColor = &H404040
    Line (0, 0)-Step(Me.ScaleWidth, borderTop), , BF
    Line (0, 0)-Step(borderLeft, Me.ScaleHeight), , BF
    Line (0, Me.ScaleHeight - borderBottom)-Step(Me.ScaleWidth, Me.ScaleHeight), , BF
    Line (Me.ScaleWidth - borderRight, 0)-Step(borderRight, Me.ScaleHeight), , BF
    
    'labels
    CurrentX = lblCashLabel.left
    CurrentY = lblCashLabel.top
    Me.ForeColor = lblCashLabel.ForeColor
    Me.FontSize = lblCashLabel.FontSize
    Print lblCashLabel.Caption
    CurrentX = lblCash.left
    CurrentY = lblCash.top
    Me.ForeColor = lblCash.ForeColor
    Me.FontSize = lblCash.FontSize
    Print lblCash.Caption
    CurrentX = lblShipLabel.left
    CurrentY = lblShipLabel.top
    Me.ForeColor = lblShipLabel.ForeColor
    Me.FontSize = lblShipLabel.FontSize
    Print lblShipLabel.Caption
    CurrentX = lblShip.left
    CurrentY = lblShip.top
    Me.ForeColor = lblShip.ForeColor
    Me.FontSize = lblShip.FontSize
    Print lblShip.Caption
    CurrentX = lblWeaponLabel.left
    CurrentY = lblWeaponLabel.top
    Me.ForeColor = lblWeaponLabel.ForeColor
    Me.FontSize = lblWeaponLabel.FontSize
    Print lblWeaponLabel.Caption
    CurrentX = lblWeapon.left
    CurrentY = lblWeapon.top
    Me.ForeColor = lblWeapon.ForeColor
    Me.FontSize = lblWeapon.FontSize
    Print lblWeapon.Caption
    CurrentX = lblHealthLabel.left
    CurrentY = lblHealthLabel.top
    Me.ForeColor = lblHealthLabel.ForeColor
    Me.FontSize = lblHealthLabel.FontSize
    Print lblHealthLabel.Caption
    CurrentX = lblHealth.left
    CurrentY = lblHealth.top
    Me.ForeColor = lblHealth.ForeColor
    Me.FontSize = lblHealth.FontSize
    Print lblHealth.Caption
    If fpsShowing = True Then
        CurrentX = lblFpsLabel.left
        CurrentY = lblFpsLabel.top
        Me.ForeColor = lblFpsLabel.ForeColor
        Me.FontSize = lblFpsLabel.FontSize
        Print lblFpsLabel.Caption
        CurrentX = lblFps.left
        CurrentY = lblFps.top
        Me.ForeColor = lblFps.ForeColor
        Me.FontSize = lblFps.FontSize
        Print lblFps.Caption
    End If
    If failedShowing = True Then
        CurrentX = lblMissionFailed.left
        CurrentY = lblMissionFailed.top
        Me.ForeColor = lblMissionFailed.ForeColor
        Me.FontSize = lblMissionFailed.FontSize
        Print lblMissionFailed.Caption
    End If
    If passedShowing = True Then
        CurrentX = lblMissionPassed.left
        CurrentY = lblMissionPassed.top
        Me.ForeColor = lblMissionPassed.ForeColor
        Me.FontSize = lblMissionPassed.FontSize
        Print lblMissionPassed.Caption
    End If
    If pausedShowing = True Then
        CurrentX = lblPaused.left
        CurrentY = lblPaused.top
        Me.ForeColor = lblPaused.ForeColor
        Me.FontSize = lblPaused.FontSize
        Print lblPaused.Caption
    End If
    
    
End Sub

Sub GameLoop()
    Dim i As Integer 'for looping
    Dim j As Integer
    Dim e As Integer
    
    'advance timeSoFar
    If goalReached = False Then
        timeSoFar = timeSoFar + et
    End If

    'end the mission
    If missionOver = True Then
        gameEndTimer = gameEndTimer - et
        If gameEndTimer <= 0 Then
            EndGame
            Exit Sub
        End If
    End If
    
    'check if time is up
    If timeSoFar >= levels(levelIndex).levelTime Then
        If levels(levelIndex).boss >= 0 Then
            If goalReached = False Then
                CreateEnemy levels(levelIndex).boss, True
                bossDead = False
            End If
        Else
            If UBound(rtEnemies) = 0 Then
                bossDead = True
            End If
        End If
        goalReached = True
    End If
    
    'check if mission over
    If goalReached = True And bossDead = True Then
        missionOver = True
    End If
    
    'enemies start timer
    enemiesStartTimer = enemiesStartTimer - et
    If enemiesStartTimer < 0 Then
        enemiesStartTimer = 0
    End If
    
    'move
    If player.alive = True Then
        Dim numberOfDirectionKeysDown As Integer
        numberOfDirectionKeysDown = 0
        If keyLeftDown = True Then
            numberOfDirectionKeysDown = numberOfDirectionKeysDown + 1
        End If
        If keyRightDown = True Then
            numberOfDirectionKeysDown = numberOfDirectionKeysDown + 1
        End If
        If keyUpDown = True Then
            numberOfDirectionKeysDown = numberOfDirectionKeysDown + 1
        End If
        If keyDownDown = True Then
            numberOfDirectionKeysDown = numberOfDirectionKeysDown + 1
        End If
        Dim diagMoveMinus As Integer
        diagMoveMinus = 0
        Dim roundedMoveSpeed As Double
        roundedMoveSpeed = RoundTo5((ships(player.shipsOwned(player.shipSelected)).moveSpeed * et))
        If numberOfDirectionKeysDown = 2 Then
            diagMoveMinus = (Sqr(2 * roundedMoveSpeed * roundedMoveSpeed) - roundedMoveSpeed) / Sqr(2)
        End If
        If keyLeftDown = True Then
            If player.left - (roundedMoveSpeed - diagMoveMinus) < 0 Then
                player.left = 0
            Else
                player.left = player.left - (roundedMoveSpeed - diagMoveMinus)
            End If
        End If
        If keyRightDown = True Then
            If player.left + (roundedMoveSpeed - diagMoveMinus) > (windowWidth - pics(ships(player.shipsOwned(player.shipSelected)).pics(player.pic)).width) Then
                player.left = (windowWidth - pics(ships(player.shipsOwned(player.shipSelected)).pics(player.pic)).width)
            Else
                player.left = player.left + (roundedMoveSpeed - diagMoveMinus)
            End If
        End If
        If keyUpDown = True Then
            If player.top - (roundedMoveSpeed - diagMoveMinus) < 0 Then
                player.top = 0
            Else
                player.top = player.top - (roundedMoveSpeed - diagMoveMinus)
            End If
        End If
        If keyDownDown = True Then
            If player.top + (roundedMoveSpeed - diagMoveMinus) > (windowHeight - pics(ships(player.shipsOwned(player.shipSelected)).pics(player.pic)).height) Then
                player.top = (windowHeight - pics(ships(player.shipsOwned(player.shipSelected)).pics(player.pic)).height)
            Else
                player.top = player.top + (roundedMoveSpeed - diagMoveMinus)
            End If
        End If
    End If
    
    'player pic timer (missionOver is set here)
    player.picTimer = player.picTimer - et
    If player.picTimer <= 0 Then
        If player.alive = True Then
            player.picTimer = ships(player.shipsOwned(player.shipSelected)).timeBetweenPics
            player.pic = player.pic + 1
            If player.pic > UBound(ships(player.shipsOwned(player.shipSelected)).pics) Then
                player.pic = 1
            End If
        Else
            player.picTimer = ships(player.shipsOwned(player.shipSelected)).timeBetweenPicsDead
            player.pic = player.pic + 1
            If player.pic > UBound(ships(player.shipsOwned(player.shipSelected)).picsDead) Then
                missionOver = True
            End If
        End If
    End If
    
    'player reload and fire
    If player.alive = True Then
        For i = LBound(player.reloadTimer) To UBound(player.reloadTimer)
            If i <> 0 Then
                player.reloadTimer(i) = player.reloadTimer(i) - et
                If player.reloadTimer(i) <= 0 Then
                    player.reloadTimer(i) = 0
                    If keyFireDown = True Then
                        CreateProjectile player.weaponsOwned(player.weaponSelected), i, True, player.left, player.top, pics(ships(player.shipsOwned(player.shipSelected)).pics(player.pic)).height, pics(ships(player.shipsOwned(player.shipSelected)).pics(player.pic)).width, -1
                        player.reloadTimer(i) = weapons(player.weaponsOwned(player.weaponSelected)).projectiles(i).reloadTime * ships(player.shipsOwned(player.shipSelected)).reloadTimeMultiplier
                    End If
                End If
            End If
        Next i
    End If
    
    'projectile movement and pic change and projectile collision
    i = 1 '0 is not used
    While i <= UBound(projectiles)
        Dim removeThisProjectile As Boolean
        removeThisProjectile = False
        
        projectiles(i).picTimer = projectiles(i).picTimer - et
        If projectiles(i).picTimer <= 0 Then
            If projectiles(i).alive = True Then
                projectiles(i).picTimer = weapons(projectiles(i).weaponFiredFrom).projectiles(projectiles(i).projectileFromWeapon).timeBetweenPics
                projectiles(i).pic = projectiles(i).pic + 1
                If projectiles(i).pic > UBound(weapons(projectiles(i).weaponFiredFrom).projectiles(projectiles(i).projectileFromWeapon).pics) Then
                    projectiles(i).pic = 1
                End If
            Else
                projectiles(i).picTimer = weapons(projectiles(i).weaponFiredFrom).projectiles(projectiles(i).projectileFromWeapon).timeBetweenPicsDead
                projectiles(i).pic = projectiles(i).pic + 1
                If projectiles(i).pic > UBound(weapons(projectiles(i).weaponFiredFrom).projectiles(projectiles(i).projectileFromWeapon).picsDead) Then
                    removeThisProjectile = True
                    projectiles(i).pic = 1
                End If
            End If
        End If
        
        Dim projectileHeight As Integer
        If projectiles(i).alive = True Then
            projectileHeight = pics(weapons(projectiles(i).weaponFiredFrom).projectiles(projectiles(i).projectileFromWeapon).pics(projectiles(i).pic)).height
        Else
            projectileHeight = pics(weapons(projectiles(i).weaponFiredFrom).projectiles(projectiles(i).projectileFromWeapon).picsDead(projectiles(i).pic)).height
        End If
        Dim projectileWidth As Integer
        If projectiles(i).alive = True Then
            projectileWidth = pics(weapons(projectiles(i).weaponFiredFrom).projectiles(projectiles(i).projectileFromWeapon).pics(projectiles(i).pic)).width
        Else
            projectileWidth = pics(weapons(projectiles(i).weaponFiredFrom).projectiles(projectiles(i).projectileFromWeapon).picsDead(projectiles(i).pic)).width
        End If
        
        Dim leftCenter As Integer
        Dim topCenter As Integer
        If projectiles(i).alive = True Then
            For j = LBound(rtEnemies) To UBound(rtEnemies)
                If j <> 0 Then
                    If rtEnemies(j).alive = True Then
                        If projectiles(i).rtEnemyFiredFrom <> j Then
                            If projectiles(i).left + projectileWidth > rtEnemies(j).left And projectiles(i).left < rtEnemies(j).left + pics(enemies(rtEnemies(j).enemyTemplate).pics(rtEnemies(j).pic)).width And projectiles(i).top + projectileHeight > rtEnemies(j).top And projectiles(i).top < rtEnemies(j).top + pics(enemies(rtEnemies(j).enemyTemplate).pics(rtEnemies(j).pic)).height Then
                                'collsion with enemy ship!
                                rtEnemies(j).health = rtEnemies(j).health - weapons(projectiles(i).weaponFiredFrom).projectiles(projectiles(i).projectileFromWeapon).damage
                                If rtEnemies(j).health <= 0 Then
                                    If projectiles(i).ownedByPlayer = True Then
                                        EnemyKilled j, True
                                    Else
                                        EnemyKilled j, False
                                    End If
                                End If
                                projectiles(i).alive = False
                                
                                leftCenter = projectiles(i).left + (pics(weapons(projectiles(i).weaponFiredFrom).projectiles(projectiles(i).projectileFromWeapon).pics(projectiles(i).pic)).width / 2)
                                topCenter = projectiles(i).top + (pics(weapons(projectiles(i).weaponFiredFrom).projectiles(projectiles(i).projectileFromWeapon).pics(projectiles(i).pic)).height / 2)
                                
                                projectiles(i).pic = 1
                                projectiles(i).picTimer = weapons(projectiles(i).weaponFiredFrom).projectiles(projectiles(i).projectileFromWeapon).timeBetweenPicsDead
                                
                                projectiles(i).left = leftCenter - (pics(weapons(projectiles(i).weaponFiredFrom).projectiles(projectiles(i).projectileFromWeapon).picsDead(projectiles(i).pic)).width / 2)
                                projectiles(i).top = topCenter - (pics(weapons(projectiles(i).weaponFiredFrom).projectiles(projectiles(i).projectileFromWeapon).picsDead(projectiles(i).pic)).height / 2)
                            End If
                        End If
                    End If
                End If
            Next j
            If player.alive = True Then
                If projectiles(i).ownedByPlayer = False Then
                    If projectiles(i).left + projectileWidth > player.left And projectiles(i).left < player.left + pics(ships(player.shipsOwned(player.shipSelected)).pics(player.pic)).width And projectiles(i).top + projectileHeight > player.top And projectiles(i).top < player.top + pics(ships(player.shipsOwned(player.shipSelected)).pics(player.pic)).height Then
                        'enemy projectile collides with player!
                        player.shipHealth(player.shipSelected) = player.shipHealth(player.shipSelected) - weapons(projectiles(i).weaponFiredFrom).projectiles(projectiles(i).projectileFromWeapon).damage
                        If player.shipHealth(player.shipSelected) <= 0 Then
                            PlayerKilled
                        End If
                        projectiles(i).alive = False
                        
                        leftCenter = projectiles(i).left + (pics(weapons(projectiles(i).weaponFiredFrom).projectiles(projectiles(i).projectileFromWeapon).pics(projectiles(i).pic)).width / 2)
                        topCenter = projectiles(i).top + (pics(weapons(projectiles(i).weaponFiredFrom).projectiles(projectiles(i).projectileFromWeapon).pics(projectiles(i).pic)).height / 2)
                        
                        projectiles(i).pic = 1
                        projectiles(i).picTimer = weapons(projectiles(i).weaponFiredFrom).projectiles(projectiles(i).projectileFromWeapon).timeBetweenPicsDead
                        
                        projectiles(i).left = leftCenter - (pics(weapons(projectiles(i).weaponFiredFrom).projectiles(projectiles(i).projectileFromWeapon).picsDead(projectiles(i).pic)).width / 2)
                        projectiles(i).top = topCenter - (pics(weapons(projectiles(i).weaponFiredFrom).projectiles(projectiles(i).projectileFromWeapon).picsDead(projectiles(i).pic)).height / 2)
                    End If
                End If
            End If
        End If
        
        If projectiles(i).alive = True Then
            projectiles(i).top = projectiles(i).top + RoundTo5(projectiles(i).moveTop * et)
            projectiles(i).left = projectiles(i).left + RoundTo5(projectiles(i).moveLeft * et)
        End If
        If projectiles(i).top < -projectileHeight Or projectiles(i).top > windowHeight Or projectiles(i).left < -projectileWidth Or projectiles(i).left > windowWidth Then
            removeThisProjectile = True
        End If
        
        If removeThisProjectile = True Then
            removeProjectile (i)
        Else
            i = i + 1
        End If
    Wend
    
    'enemy generation
    If enemiesStartTimer <= 0 Then
        If goalReached = False Then
            For i = LBound(levels(levelIndex).enemiesPresent) To UBound(levels(levelIndex).enemiesPresent)
                If i <> 0 Then
                    If levels(levelIndex).enemiesPresent(i).enemy > 0 Then
                        If Rnd() < levels(levelIndex).enemiesPresent(i).chanceOfCreation * et Then
                            Dim onScreenCounter As Integer
                            onScreenCounter = 0
                            For e = LBound(rtEnemies) To UBound(rtEnemies)
                                If e <> 0 Then
                                    If rtEnemies(e).enemyTemplate = levels(levelIndex).enemiesPresent(i).enemy Then
                                        onScreenCounter = onScreenCounter + 1
                                    End If
                                End If
                            Next e
                            If onScreenCounter < levels(levelIndex).enemiesPresent(i).maxOnScreen Then
                                CreateEnemy levels(levelIndex).enemiesPresent(i).enemy, False
                                e = 1
                                While e <= UBound(rtEnemies) - 1
                                    If rtEnemies(e).alive = True Then
                                        If rtEnemies(UBound(rtEnemies)).left + pics(enemies(rtEnemies(UBound(rtEnemies)).enemyTemplate).pics(rtEnemies(UBound(rtEnemies)).pic)).width > rtEnemies(e).left And rtEnemies(UBound(rtEnemies)).left < rtEnemies(e).left + pics(enemies(rtEnemies(e).enemyTemplate).pics(rtEnemies(e).pic)).width Then
                                            If rtEnemies(UBound(rtEnemies)).top + pics(enemies(rtEnemies(UBound(rtEnemies)).enemyTemplate).pics(rtEnemies(UBound(rtEnemies)).pic)).height > rtEnemies(e).top And rtEnemies(UBound(rtEnemies)).top < rtEnemies(e).top + pics(enemies(rtEnemies(e).enemyTemplate).pics(rtEnemies(e).pic)).height Then
                                                e = UBound(rtEnemies)
                                                removeEnemy (UBound(rtEnemies))
                                            End If
                                        End If
                                    End If
                                    e = e + 1
                                Wend
                            End If
                        End If
                    End If
                End If
            Next i
        End If
    End If
    
    'enemy movement and pic change and reload and fire
    i = 1 '0 is not used
    While i <= UBound(rtEnemies)
        Dim removeThisEnemy As Boolean
        removeThisEnemy = False
        
        rtEnemies(i).picTimer = rtEnemies(i).picTimer - et
        If rtEnemies(i).picTimer <= 0 Then
            If rtEnemies(i).alive = True Then
                rtEnemies(i).picTimer = enemies(rtEnemies(i).enemyTemplate).timeBetweenPics
                rtEnemies(i).pic = rtEnemies(i).pic + 1
                If rtEnemies(i).pic > UBound(enemies(rtEnemies(i).enemyTemplate).pics) Then
                    rtEnemies(i).pic = 1
                End If
            Else
                rtEnemies(i).picTimer = enemies(rtEnemies(i).enemyTemplate).timeBetweenPicsDead
                rtEnemies(i).pic = rtEnemies(i).pic + 1
                If rtEnemies(i).pic > UBound(enemies(rtEnemies(i).enemyTemplate).picsDead) Then
                    removeThisEnemy = True
                    rtEnemies(i).pic = 1
                End If
            End If
        End If
        
        Dim enemyHeight As Integer
        If rtEnemies(i).alive = True Then
            enemyHeight = pics(enemies(rtEnemies(i).enemyTemplate).pics(rtEnemies(i).pic)).height
        Else
            enemyHeight = pics(enemies(rtEnemies(i).enemyTemplate).picsDead(rtEnemies(i).pic)).height
        End If
        Dim enemyWidth As Integer
        If rtEnemies(i).alive = True Then
            enemyWidth = pics(enemies(rtEnemies(i).enemyTemplate).pics(rtEnemies(i).pic)).width
        Else
            enemyWidth = pics(enemies(rtEnemies(i).enemyTemplate).picsDead(rtEnemies(i).pic)).width
        End If
        
        If rtEnemies(i).boss = False Then
            rtEnemies(i).top = rtEnemies(i).top + (rtEnemies(i).moveTop * et)
            rtEnemies(i).left = rtEnemies(i).left + (rtEnemies(i).moveLeft * et)
        Else
            If rtEnemies(i).top + (rtEnemies(i).moveTop * et) > rtEnemies(i).top Then
                If rtEnemies(i).top + (enemyHeight / 2) < windowHeight * levels(levelIndex).bossFinalY Then
                    rtEnemies(i).top = rtEnemies(i).top + RoundTo5(rtEnemies(i).moveTop * et)
                End If
            Else
                If rtEnemies(i).top + (enemyHeight / 2) > windowHeight * levels(levelIndex).bossFinalY Then
                    rtEnemies(i).top = rtEnemies(i).top + RoundTo5(rtEnemies(i).moveTop * et)
                End If
            End If
            If rtEnemies(i).left + (rtEnemies(i).moveLeft * et) > rtEnemies(i).left Then
                If rtEnemies(i).left + (enemyWidth / 2) < windowWidth * levels(levelIndex).bossFinalX Then
                    rtEnemies(i).left = rtEnemies(i).left + RoundTo5(rtEnemies(i).moveLeft * et)
                End If
            Else
                If rtEnemies(i).left + (enemyWidth / 2) > windowWidth * levels(levelIndex).bossFinalX Then
                    rtEnemies(i).left = rtEnemies(i).left + RoundTo5(rtEnemies(i).moveLeft * et)
                End If
            End If
        End If
        If rtEnemies(i).top < -enemyHeight Or rtEnemies(i).top > windowHeight Or rtEnemies(i).left < -enemyWidth Or rtEnemies(i).left > windowWidth Then
            removeThisEnemy = True
            If rtEnemies(i).boss = True Then
                bossDead = True
            End If
        End If
        
        If enemies(rtEnemies(i).enemyTemplate).weapon >= 0 Then
            If rtEnemies(i).alive = True Then
                Dim m As Integer
                For m = LBound(rtEnemies(i).reloadTimer) To UBound(rtEnemies(i).reloadTimer)
                    If m <> 0 Then
                        rtEnemies(i).reloadTimer(m) = rtEnemies(i).reloadTimer(m) - et
                        If rtEnemies(i).reloadTimer(m) <= 0 Then
                            CreateProjectile enemies(rtEnemies(i).enemyTemplate).weapon, m, False, rtEnemies(i).left, rtEnemies(i).top, enemyHeight, enemyWidth, i
                            rtEnemies(i).reloadTimer(m) = weapons(projectiles(UBound(projectiles)).weaponFiredFrom).projectiles(projectiles(UBound(projectiles)).projectileFromWeapon).reloadTime
                        End If
                    End If
                Next m
            End If
        End If
        
        If removeThisEnemy = True Then
            removeEnemy (i)
        Else
            i = i + 1
        End If
    Wend
    
    'player enemy collision and collision timer
    If player.alive = True Then
        playerCollisionTimer = playerCollisionTimer - et
        If playerCollisionTimer <= 0 Then
            playerCollisionTimer = 0
            If ships(player.shipSelected).collisionWeapon >= 0 And ships(player.shipSelected).collisionProjectileFromWeapon >= 0 Then
                For i = LBound(rtEnemies) To UBound(rtEnemies)
                    If i <> 0 Then
                        If rtEnemies(i).alive = True Then
                            If player.left + pics(ships(player.shipsOwned(player.shipSelected)).pics(1)).width > rtEnemies(i).left And player.left < rtEnemies(i).left + pics(enemies(rtEnemies(i).enemyTemplate).pics(1)).width Then
                                If player.top + pics(ships(player.shipsOwned(player.shipSelected)).pics(1)).height > rtEnemies(i).top And player.top < rtEnemies(i).top + pics(enemies(rtEnemies(i).enemyTemplate).pics(1)).height Then
                                    ReDim Preserve projectiles(UBound(projectiles) + 1)
                                    projectiles(UBound(projectiles)).alive = False
                                    projectiles(UBound(projectiles)).weaponFiredFrom = ships(player.shipSelected).collisionWeapon
                                    projectiles(UBound(projectiles)).projectileFromWeapon = ships(player.shipSelected).collisionProjectileFromWeapon
                                    projectiles(UBound(projectiles)).pic = 1
                                    projectiles(UBound(projectiles)).picTimer = weapons(ships(player.shipSelected).collisionWeapon).projectiles(ships(player.shipSelected).collisionProjectileFromWeapon).timeBetweenPicsDead
                                    projectiles(UBound(projectiles)).ownedByPlayer = True 'this is irelevant
                                    projectiles(UBound(projectiles)).left = rtEnemies(i).left + ((player.left + pics(ships(player.shipsOwned(player.shipSelected)).pics(player.pic)).width - rtEnemies(i).left) / 2) - (pics(weapons(projectiles(UBound(projectiles)).weaponFiredFrom).projectiles(projectiles(UBound(projectiles)).projectileFromWeapon).pics(projectiles(UBound(projectiles)).pic)).width / 2)
                                    projectiles(UBound(projectiles)).top = rtEnemies(i).top + ((player.top + pics(ships(player.shipsOwned(player.shipSelected)).pics(player.pic)).height - rtEnemies(i).top) / 2) - (pics(weapons(projectiles(UBound(projectiles)).weaponFiredFrom).projectiles(projectiles(UBound(projectiles)).projectileFromWeapon).pics(projectiles(UBound(projectiles)).pic)).height / 2)
                                    
                                    player.shipHealth(player.shipSelected) = player.shipHealth(player.shipSelected) - weapons(ships(player.shipSelected).collisionWeapon).projectiles(ships(player.shipSelected).collisionProjectileFromWeapon).damage
                                    If player.shipHealth(player.shipSelected) <= 0 Then
                                        PlayerKilled
                                    End If
                                    rtEnemies(i).health = rtEnemies(i).health - weapons(ships(player.shipSelected).collisionWeapon).projectiles(ships(player.shipSelected).collisionProjectileFromWeapon).damage
                                    If rtEnemies(i).health <= 0 Then
                                        EnemyKilled i, True
                                    End If
                                    playerCollisionTimer = weapons(ships(player.shipSelected).collisionWeapon).projectiles(ships(player.shipSelected).collisionProjectileFromWeapon).reloadTime
                                End If
                            End If
                        End If
                    End If
                Next i
            End If
        End If
    End If
                
    'background objects movement
    i = 1
    While i <= UBound(rtBackgroundobjects)
        Dim removeThisBackgroundobject As Boolean
        removeThisBackgroundobject = False
        
        rtBackgroundobjects(i).picTimer = rtBackgroundobjects(i).picTimer - et
        If rtBackgroundobjects(i).picTimer <= 0 Then
            rtBackgroundobjects(i).picTimer = backgroundobjects(rtBackgroundobjects(i).backgoundobjectTemplate).timeBetweenPics
            rtBackgroundobjects(i).pic = rtBackgroundobjects(i).pic + 1
            If rtBackgroundobjects(i).pic > UBound(backgroundobjects(rtBackgroundobjects(i).backgoundobjectTemplate).pics) Then
                rtBackgroundobjects(i).pic = 1
            End If
        End If
        
        
        rtBackgroundobjects(i).top = rtBackgroundobjects(i).top + RoundTo5(backgroundobjects(rtBackgroundobjects(i).backgoundobjectTemplate).moveSpeed * et)
        If rtBackgroundobjects(i).top > windowHeight Or rtBackgroundobjects(i).top < -pics(backgroundobjects(rtBackgroundobjects(i).backgoundobjectTemplate).pics(rtBackgroundobjects(i).pic)).height Then
            removeThisBackgroundobject = True
        End If
        If rtBackgroundobjects(i).backImageTop = True Then
            If rtBackgroundobjects(i).top > 0 Then
                rtBackgroundobjects(i).backImageTop = False
                nextBackImageTop = rtBackgroundobjects(i).top - pics(backgroundobjects(levels(levelIndex).backImageBackgroundobject).pics(1)).height
            End If
        End If
        
        If removeThisBackgroundobject = True Then
            removeBackgroundobject (i)
        Else
            i = i + 1
        End If
    Wend
    
    'background objects generation
    For i = LBound(levels(levelIndex).backgroundobjectsPresent) To UBound(levels(levelIndex).backgroundobjectsPresent)
        If i <> 0 Then
            If levels(levelIndex).backgroundobjectsPresent(i) > 0 Then
                If Rnd() < backgroundobjects(levels(levelIndex).backgroundobjectsPresent(i)).chanceOfCreation * et Then
                    CreateBackgroundobject levels(levelIndex).backgroundobjectsPresent(i), False
                End If
            End If
        End If
    Next i
    
    'backimage generation
    If levels(levelIndex).backImageBackgroundobject >= 0 Then
        If nextBackImageTop <= 0 Then
            Dim leftUpto As Integer
            leftUpto = 0
            Dim firstOneCreated As Integer
            firstOneCreated = UBound(rtBackgroundobjects) + 1
            While leftUpto <= windowWidth
                ReDim Preserve rtBackgroundobjects(UBound(rtBackgroundobjects) + 1)
                rtBackgroundobjects(UBound(rtBackgroundobjects)).backgoundobjectTemplate = levels(levelIndex).backImageBackgroundobject
                rtBackgroundobjects(UBound(rtBackgroundobjects)).pic = 1
                rtBackgroundobjects(UBound(rtBackgroundobjects)).picTimer = backgroundobjects(levels(levelIndex).backImageBackgroundobject).timeBetweenPics
                rtBackgroundobjects(UBound(rtBackgroundobjects)).top = nextBackImageTop
                rtBackgroundobjects(UBound(rtBackgroundobjects)).left = leftUpto
                leftUpto = leftUpto + pics(backgroundobjects(levels(levelIndex).backImageBackgroundobject).pics(1)).width
                rtBackgroundobjects(UBound(rtBackgroundobjects)).backImage = True
                rtBackgroundobjects(UBound(rtBackgroundobjects)).backImageTop = False
            Wend
            rtBackgroundobjects(firstOneCreated).backImageTop = True
            nextBackImageTop = 1
        End If
    End If
    
    DrawScreen
End Sub

Sub StartGame()
    fps = 0
    progressiveFps = 0
    etThisSecond = 0

    fpsShowing = False
    pausedShowing = False
    failedShowing = False
    passedShowing = False

    playerCollisionTimer = 0
    nextBackImageTop = 0

    windowWidth = Me.ScaleWidth - borderLeft - borderRight
    windowHeight = Me.ScaleHeight - borderTop - borderBottom

    missionOver = False
    gameEndTimer = gameEndTime
    gamePaused = False
    timeSoFar = 0
    bossDead = False
    goalReached = False
    'If player.level > UBound(levels) Then
    '    levelIndex = 0
    'Else
    '    levelIndex = player.level
    'End If
    levelIndex = player.level
    Me.Caption = baseFormCaption & " - " & modName & " - " & levels(levelIndex).title
    enemiesStartTimer = levels(levelIndex).enemiesStartTime
    
    keyLeftDown = False
    keyRightDown = False
    keyUpDown = False
    keyDownDown = False
    keyFireDown = False
    
    Me.BackColor = levels(levelIndex).backgroundColour
    
    lblMissionPassed.top = borderTop + (windowHeight / 2.5) - (lblMissionPassed.height / 2)
    lblMissionPassed.left = borderLeft + (windowWidth / 2) - (lblMissionPassed.width / 2)
    lblMissionFailed.top = borderTop + (windowHeight / 2.5) - (lblMissionFailed.height / 2)
    lblMissionFailed.left = borderLeft + (windowWidth / 2) - (lblMissionFailed.width / 2)
    lblPaused.top = lblMissionPassed.top + lblMissionPassed.height + (lblPaused.height * 2)
    lblPaused.left = borderLeft + (windowWidth / 2) - (lblPaused.width / 2)
    
    pbHealth.Max = ships(player.shipsOwned(player.shipSelected)).maxHealth
    player.weaponSelected = player.weapon1
    player.alive = True
    player.picTimer = ships(player.shipsOwned(player.shipSelected)).timeBetweenPics
    player.pic = 1
    player.left = windowWidth / 2
    player.top = windowHeight / 2
    
    'this may be put in seperate function when player can switch weapons
    Dim i As Integer
    ReDim player.reloadTimer(UBound(weapons(player.weaponsOwned(player.weaponSelected)).projectiles))
    For i = LBound(player.reloadTimer) To UBound(player.reloadTimer)
        If i <> 0 Then
            player.reloadTimer(i) = 0
        End If
    Next i
    
    ReDim rtBackgroundobjects(0)
    ReDim rtEnemies(0)
    ReDim projectiles(0)
    
    FillScreenWithBackgroundobjects
    
    StartTheTimer
End Sub

Sub EndGame()
    If player.alive = False Then
        player = backupPlayer
    Else
        'player.level = player.level + 1
        If UBound(player.levelsPassed) < UBound(levels) - 1 Then
            ReDim Preserve player.levelsPassed(UBound(player.levelsPassed) + 1)
            player.levelsPassed(UBound(player.levelsPassed)) = player.level
        End If
        
        backupPlayer = player
    End If
    gamePaused = False
    StopTheTimer

    Me.Hide
    Unload Me
    frmMenu.Show
    frmMenu.Update
End Sub

Sub PauseGame()
    If gamePaused = True Then
        gamePaused = False
        pausedShowing = False
    Else
        gamePaused = True
        pausedShowing = True
    End If
    DrawScreen
End Sub

Sub removeProjectile(elementOfProjectileArray As Integer)
    If elementOfProjectileArray = 0 Then
        MsgBox "You shouldnt delete the 0th element of the projectile array!"
    End If
    Dim newProjectileArray() As typeRuntimeProjectile
    ReDim newProjectileArray(UBound(projectiles) - 1)
    Dim counter As Integer
    counter = LBound(projectiles)
    Dim i As Integer
    For i = LBound(projectiles) To UBound(projectiles)
        If i <> elementOfProjectileArray Then
            newProjectileArray(counter) = projectiles(i)
            counter = counter + 1
        End If
    Next i
    ReDim projectiles(LBound(newProjectileArray) To UBound(newProjectileArray))
    For i = LBound(projectiles) To UBound(projectiles)
        projectiles(i) = newProjectileArray(i)
    Next i
End Sub

Sub removeEnemy(elementOfEnemyArray As Integer)
    If elementOfEnemyArray = 0 Then
        MsgBox "You shouldnt delete the 0th element of the enemy array!"
    End If
    Dim i As Integer
    For i = LBound(projectiles) To UBound(projectiles)
        If i > 0 Then
            If projectiles(i).rtEnemyFiredFrom = elementOfEnemyArray Then
                projectiles(i).rtEnemyFiredFrom = -1
            End If
        End If
    Next i
    Dim newEnemyArray() As typeRuntimeEnemy
    ReDim newEnemyArray(UBound(rtEnemies) - 1)
    Dim counter As Integer
    counter = LBound(rtEnemies)
    For i = LBound(rtEnemies) To UBound(rtEnemies)
        If i <> elementOfEnemyArray Then
            newEnemyArray(counter) = rtEnemies(i)
            counter = counter + 1
        End If
    Next i
    ReDim rtEnemies(LBound(newEnemyArray) To UBound(newEnemyArray))
    For i = LBound(rtEnemies) To UBound(rtEnemies)
        rtEnemies(i) = newEnemyArray(i)
    Next i
End Sub

Sub removeBackgroundobject(elementOfBackgroundobjectArray As Integer)
    If elementOfBackgroundobjectArray = 0 Then
        MsgBox "You shouldnt delete the 0th element of the backgroundobject array!"
    End If
    Dim newBackgroundobjectArray() As typeRuntimeBackgroundobject
    ReDim newBackgroundobjectArray(UBound(rtBackgroundobjects) - 1)
    Dim counter As Integer
    counter = LBound(rtBackgroundobjects)
    Dim i As Integer
    For i = LBound(rtBackgroundobjects) To UBound(rtBackgroundobjects)
        If i <> elementOfBackgroundobjectArray Then
            newBackgroundobjectArray(counter) = rtBackgroundobjects(i)
            counter = counter + 1
        End If
    Next i
    ReDim rtBackgroundobjects(LBound(newBackgroundobjectArray) To UBound(newBackgroundobjectArray))
    For i = LBound(rtBackgroundobjects) To UBound(rtBackgroundobjects)
        rtBackgroundobjects(i) = newBackgroundobjectArray(i)
    Next i
End Sub

