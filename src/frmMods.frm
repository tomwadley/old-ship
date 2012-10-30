VERSION 5.00
Begin VB.Form frmMods 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Ship Mod"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   4005
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdSetDefault 
      Caption         =   "Set Default"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Load Mod"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.FileListBox flbMods 
      Height          =   1260
      Left            =   1440
      Pattern         =   "*.xml"
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lstMods 
      Height          =   2010
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label lblModDescription 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
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
      Height          =   2055
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mod Description:"
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
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lblDefaultMod 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "The default mod"
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
      Left            =   1320
      TabIndex        =   7
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label14 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Default Mod:"
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
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Mods detected:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmMods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mods() As String
Dim modSelected As Integer

Private Sub cmdCancel_Click()
    CloseForm
End Sub

Private Sub cmdOk_Click()
    If modSelected >= 0 Then
        If mods(0, modSelected) = modFileName And modLoaded = True Then
            CloseForm
        Else
            If gameLoaded = True Then
                If MsgBox("If you switch mods you will lose any unsaved progress in the game you have loaded now. Continue?", vbOKCancel) = vbCancel Then
                    Exit Sub
                End If
            End If
            modFileName = mods(0, modSelected)
            frmMenu.LoadMod
            If modLoaded = True Then
                CloseForm
                frmMenu.Update
            End If
        End If
    Else
        CloseForm
    End If
End Sub

Private Sub cmdSetDefault_Click()
    If modSelected >= 0 Then
        defaultModFileName = mods(0, modSelected)
        On Error GoTo fileerror
        Open App.Path & "\defaultmod" For Output As #1
        Print #1, defaultModFileName
        Close #1
        lblDefaultMod.Caption = mods(1, modSelected)
    Else
        MsgBox "Select a mod first.."
    End If
    Exit Sub
fileerror:
    MsgBox "error setting default mod"
End Sub

Private Sub Form_Load()
    modSelected = -1
    UpdateModList
End Sub

Sub UpdateModList()
    lblDefaultMod.Caption = Empty
    Dim xml_document As DOMDocument
    Dim root As IXMLDOMNode
    flbMods.Path = App.Path
    flbMods.Refresh
    lstMods.Clear
    ReDim mods(0 To 2, 0)
    Dim i As Integer
    i = flbMods.ListCount - 1
    While i >= 0
        Set xml_document = New DOMDocument
        xml_document.Load App.Path & "\" & flbMods.List(i)
        'check if the file was opened correctly
        If Not xml_document.documentElement Is Nothing Then
            'find the shipmod node
            Set root = xml_document.selectSingleNode("shipmod")
            If Not root Is Nothing Then
                'get name
                Dim nameAttribute As IXMLDOMAttribute
                Set nameAttribute = root.Attributes.getNamedItem("name")
                'load description
                Dim descriptionNode As IXMLDOMElement
                Set descriptionNode = root.selectSingleNode("description")
                'insert enry
                ReDim Preserve mods(0 To 2, UBound(mods, 2) + 1)
                mods(0, UBound(mods, 2)) = flbMods.List(i)
                mods(1, UBound(mods, 2)) = nameAttribute.Text
                mods(2, UBound(mods, 2)) = descriptionNode.Text
                lstMods.AddItem nameAttribute.Text, UBound(mods, 2) - 1
                If modFileName = flbMods.List(i) Then
                    lstMods.ListIndex = lstMods.NewIndex
                    lblModDescription.Caption = mods(2, UBound(mods, 2))
                End If
                If defaultModFileName = flbMods.List(i) Then
                    lblDefaultMod.Caption = mods(1, UBound(mods, 2))
                End If
            End If
        End If
        Set xml_document = Nothing
        Set root = Nothing
        Set nameAttribute = Nothing
        Set descriptionNode = Nothing
        i = i - 1
    Wend
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    CloseForm
End Sub

Private Sub lstMods_Click()
    modSelected = lstMods.ListIndex + 1
    
    If modSelected >= 0 Then
        lblModDescription.Caption = mods(2, modSelected)
    End If
End Sub

Sub CloseForm()
    Me.Hide
    frmMenu.Show
End Sub
