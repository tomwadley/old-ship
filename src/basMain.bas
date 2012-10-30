Attribute VB_Name = "basMain"
'ship 0.4 - tom wadley

'todo:
'load data from disk - make dev tool
'save/load games
'track the nearest enemy?
'make level semi-scripting better (do something clever after levels end) - how about start from the begining but goto harder skill
'maybe make background objects a little better, prob ok how they are now
'possibility of tilting ships
'only load pics that are needed?
'marathon mode - classic ship?
'use a schema insted of dtd?
'level end message - maybe display this instead of "mission passed"

'bugs: * - not sure if they still exist
'its possible that tracking isnt perfect yet
'weapon switching and reload times. might have to make reloadtime array for each weapon
'*possible bug that causes health to be less than it should be after death
'*player can sometimes exist on 0 health
'*could be casued by for looping my arrays with dummy 0 badly (edit: nah)
'*holding fire when dying causes subscript out of range (array not initialised) (not sure if this still happens)

Option Explicit

Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function timeBeginPeriod Lib "winmm.dll" (uPeriod) As Long
Public Declare Function timeEndPeriod Lib "winmm.dll" (uPeriod) As Long
  '16-bit programs must use MMSYSTEM.DLL instead of WINMM.DLL

Public twTimerStarted As Boolean
Public et As Double 'elapsed time since last cycle (seconds)

Const frameLimit = 90 'on some computers if this is high the player will experience control lag

Public Const fileVersion = 1
Public Const saveVersion = 1

Public defaultModFileName As String
Public modFileName As String
Public modLoaded As Boolean
Public gameLoaded As Boolean
Public onMission As Boolean
Public gamePaused As Boolean
Public modName As String
Public modDescription As String
Public picMaskColour As Long
Public healthPriceQty As Integer
Public healthPriceCost As Long
Public savedGamePath As String

Private Type typePic
    filePath As String
    height As Integer
    width As Integer
End Type

Private Type typeBackgroundObject
    obName As String
    moveSpeed As Integer
    pics() As Integer
    timeBetweenPics As Double
    chanceOfCreation As Double 'chance per second
End Type

Public Const wpCustom = 0 'used in row and column
Public Const wpFarLeft = 1
Public Const wpLeft = 2
Public Const wpRight = 3
Public Const wpFarRight = 4
Public Const wpFarFront = 1
Public Const wpFront = 2
Public Const wpBack = 3
Public Const wpFarBack = 4
Private Type typeProjectile
    obName As String
    column As Integer 'wpLeft wpCenter wpRight
    row As Integer 'wpFront wpMiddle wpBack
    customColumn As Double 'these are ratios of the ship/enemy firing the projectile
    customRow As Double
    moveTop As Integer 'pixels per second. +or-
    moveLeft As Integer 'pixels per second. +or-
    tracksPlayer As Boolean 'use moveTop as the movement speed
    pics() As Integer
    timeBetweenPics As Double
    picsDead() As Integer
    timeBetweenPicsDead As Double
    damage As Integer
    reloadTime As Double
    initReloadTime As Double
End Type

Private Type typeWeapon
    obName As String
    'reloadTime As Double
    projectiles() As typeProjectile
    cost As Long 'negative number means not in shop
    description As String
    title As String
End Type

Public Const epTop = 0
Public Const epBottom = 1
Public Const epLeft = 2
Public Const epRight = 3
Public Const epLeftSide = 0
Public Const epRightSide = 1
Public Const epMiddleSide = 2
Public Const epAnywhereSide = 3
Private Type typeEnemy
    obName As String
    minMoveTop As Integer
    maxMoveTop As Integer
    minMoveLeft As Integer
    maxMoveLeft As Integer
    maxHealth As Integer
    weapon As Integer 'index into weapons
    weaponName As String
    pics() As Integer
    timeBetweenPics As Double
    picsDead() As Integer
    timeBetweenPicsDead As Double
    cash As Long
    initReloadTime As Double 'this gets added to each of this enemies projectiles initReloadTimes (currently this does nothing to human player)
    'enemy generation info
    entryPoint As Integer 'epTop epBottom epLeft epRight
    entrySide As Integer 'epLeftSide epRightSide epMiddleSide epAnywhereSide
    entrySideChange As Double 'amount entryside can change by (doesnt affect epAnywhereSide)  'ratio of form
End Type

Private Type typeShip
    obName As String
    moveSpeed As Integer 'pixels per second
    maxHealth As Integer
    reloadTimeMultiplier As Double
    collisionWeapon As Integer 'index into weapons
    collisionWeaponName As String
    collisionProjectileFromWeapon As Integer 'index into weapons(collisionWeapon).projectiles()
    collisionProjectileFromWeaponName As String
    pics() As Integer
    timeBetweenPics As Double
    picsDead() As Integer
    timeBetweenPicsDead As Double
    cost As Long 'negative number means not in shop
    description As String
    title As String
End Type

'TODO: must convert player to the way i do everything else
Private Type typePlayer
    left As Integer
    top As Integer
    alive As Boolean
    shipHealth() As Integer 'this array is the same size as shipsOwned as stores the health for each ship
    reloadTimer() As Double 'time till next shot can be made (for each projectile in players current weapon)
    weaponSelected As Integer 'index into weaponsOwned
    weaponsOwned() As Integer 'index into weapons
    weaponsOwnedNames() As String
    weapon1 As Integer 'index into weaponsOwned
    weapon1Name As String
    weapon2 As Integer 'index into weaponsOwned
    weapon2Name As String
    pic As Integer
    picTimer As Double
    cash As Long
    level As Integer
    levelsPassed() As Integer
    levelsPassedNames() As String
    shipSelected As Integer 'index into shipsOwned
    shipSelectedName As String
    shipsOwned() As Integer 'index into ships
    shipsOwnedNames() As String
End Type

Private Type typeLevelEnemies
    enemy As Integer 'index into enemies()
    enemyName As String
    chanceOfCreation As Double
    maxOnScreen As Integer
End Type

Private Type typeLevel
    obName As String
    title As String
    description As String
    dependencies() As Integer
    dependenciesNames() As String
    enemiesPresent() As typeLevelEnemies
    'cashRequired As Long 'cash scored in this level to pass (del this later)
    levelTime As Integer 'seconds till level passed (or upto boss)
    backgroundobjectsPresent() As Integer 'index into backgroundobjects()
    backgroundobjectsPresentNames() As String
    boss As Integer 'index into enemies() -1 for no boss
    bossName As String 'empty string for no boss
    backgroundColour As Long
    backImageBackgroundobject As Integer 'index into backgroundobjects()
    backImageBackgroundobjectName As String
    enemiesStartTime As Integer 'seconds
    'the following are ratios of the window height and width (eg 0.25 is a quarter of the form) these values refer to the center of the boss, not the bosses left and top
    bossFinalX As Double 'where the boss sits after entering screen
    bossFinalY As Double 'where the boss sits after entering screen
End Type

Private Type typeNewGame
    weaponsOwnedNames() As String
    shipsOwnedNames() As String
End Type

Public backgroundobjects() As typeBackgroundObject
Public weapons() As typeWeapon
Public enemies() As typeEnemy
Public levels() As typeLevel
Public ships() As typeShip

Public newgameinfo As typeNewGame
Public player As typePlayer
Public backupPlayer As typePlayer

Public pics() As typePic

'call this to start the gameloop (and set twTimerStarted=True)
'to stop set twTimerStarted=False
Sub TimedCode()
    Dim prevTime As Long
    prevTime = timeGetTime
    et = 0
    timeBeginPeriod 1
    Do
        While gamePaused = True
            DoEvents
            prevTime = timeGetTime
        Wend
        If twTimerStarted = False Then
            timeEndPeriod 1
            Exit Sub
        End If
        
        Dim frameLimitTest As Double
        frameLimitTest = frameLimit + 1
        Dim NewTime As Long
        While frameLimitTest > frameLimit
            NewTime = timeGetTime
            et = (NewTime - prevTime)
            If et = 0 Then
                et = 1
            End If
            frameLimitTest = 1000 / et
            DoEvents
        Wend
        et = et / 1000
        If et < 0 Then
            et = 0
        End If
        If et > 0.25 Then
            et = 0
        End If
        
        If twTimerStarted = False Then
            Exit Sub
        End If
        frmGame.GameLoop
        prevTime = NewTime
        DoEvents
    Loop
End Sub

Public Sub StartTheTimer()
    twTimerStarted = True
    TimedCode
End Sub

Public Sub StopTheTimer()
    twTimerStarted = False
End Sub

Function RoundTo5(number As Double) As Double
'    If Round(number / 10) * 10 = 0 Then
'        RoundTo5 = number
'    Else
'        RoundTo5 = Round(number / 10) * 10
'    End If
    Dim r As Double
    r = -1
    Dim d As Integer
    d = ((number / 10) - Int(number / 10)) * 10 * 2
    If d < 5 Then
        r = 0
    ElseIf d >= 5 And d < 15 Then
        r = 0.5
    ElseIf d >= 15 Then
        r = 1
    End If
    If r = -1 Then
        MsgBox number & " " & d
    End If
    RoundTo5 = (Int(number / 10) + r) * 10
    If RoundTo5 = 0 Then
        RoundTo5 = number
    End If
End Function

Sub Main()
    Randomize
    
    modLoaded = False
    gameLoaded = False
    onMission = False
    savedGamePath = Empty
    
    frmMenu.Show
End Sub
