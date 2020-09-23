VERSION 5.00
Begin VB.Form frmTutorListBox 
   Caption         =   "ListBox Tutorial"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Simple Tutorial on ListBox"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3255
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   6690
      Begin VB.TextBox txtAddItem 
         Height          =   315
         Left            =   270
         TabIndex        =   1
         Top             =   2805
         Width           =   2475
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   360
         Left            =   2940
         TabIndex        =   4
         Top             =   2025
         Width           =   840
      End
      Begin VB.ListBox lstItems 
         Height          =   1230
         Left            =   240
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1170
         Width           =   2500
      End
      Begin VB.ListBox lstThings 
         Height          =   1230
         Left            =   3960
         MultiSelect     =   1  'Simple
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1170
         Width           =   2500
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&>"
         Height          =   360
         Left            =   2940
         TabIndex        =   2
         Top             =   1155
         Width           =   840
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "<"
         Height          =   360
         Left            =   2940
         TabIndex        =   3
         Top             =   1590
         Width           =   840
      End
      Begin VB.Label Label3 
         Caption         =   "Press <ENTER> Key to add item"
         Height          =   195
         Left            =   270
         TabIndex        =   9
         Top             =   2580
         Width           =   2355
      End
      Begin VB.Label Label2 
         Caption         =   "Simple MultiSelect - Click to select and to deselect an item."
         Height          =   390
         Left            =   3990
         TabIndex        =   8
         Top             =   690
         Width           =   2580
      End
      Begin VB.Label Label1 
         Caption         =   "Extended MultiSelect - combine mouse-click with either SHIFT or CTRL key"
         Height          =   600
         Left            =   285
         TabIndex        =   7
         Top             =   525
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmTutorListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Very Basic tutorial on multi-select property of listbox, how to add and remove items from it
' To add items to a list box, use "AddItem" method  (See the form's Load event for an example)
' To Remove an item from list box, use "RemoveItems" method (see the command button's click event)
'
'   Multi-Select property of list box
'   0 - Mone     -> No multiselection feature, you can only select one item  (this is the default)
'   1 - Simple   -> you can multiselect by clicking an item, and deselecting, by clicking also  (see lstThing object)
'   2 - Extended -> Multiselect by combining mouse-click with either SHIFT key or CTRL          (see lstItems object)
'
' This is intended for absolute beginners.
' If you find this useful, please vote.
' thanks and if you have questions, just contact me.
'  my email : cMorning30@hotmail.com
'
Option Explicit

Private Sub cmdBack_Click()
    Dim intSelect As Integer    ' declare a variable representing the index of the listbox
    intSelect = 0               ' of course, set it to 0, the starting point of its index value
                                ' it's ok if you dont assign values, it will be set to its default value, which is 0
                                ' but its always a good programming practice to assign values to a variable
    
    ' Loop inside the list box and test the item if it is selected
    While intSelect < (lstThings.ListCount)
        If lstThings.Selected(intSelect) = True Then
            lstItems.AddItem lstThings.List(intSelect)
            lstThings.RemoveItem intSelect
        Else
            intSelect = intSelect + 1   ' increment this by one, so we can move on the the next item in the list
        End If
    Wend
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdRemove_Click()
    Dim intSelect As Integer    ' declare a variable representing the index of the listbox
    intSelect = 0               ' of course, set it to 0, the starting point of its index value
                                ' it's ok if you dont assign values, it will be set to its default value, which is 0
                                ' but its always a good programming practice to assign values to a variable
    
    ' Loop inside the list box and test the item if it is selected
    While intSelect < (lstItems.ListCount)
        If lstItems.Selected(intSelect) = True Then
            lstThings.AddItem lstItems.List(intSelect)      'it is important to add the items first to the other listbox
                                                            ' otherwise, it will add the wrong value!
            lstItems.RemoveItem intSelect                   ' then remove it from the original list box
        Else
            intSelect = intSelect + 1   ' increment this by one, so we can move on the the next item in the list
        End If
    Wend

' note that we can not use FOR Loop here because if we remove an
' item from the listbox, the listcount acually decreases by one,
' therefore making the FOR value inconsistent with the index.
' you may want try it, if want to experiment.

End Sub

Private Sub Form_Load()
    
    'Center the form
    Top = (Screen.Height - ScaleHeight) / 2
    Left = (Screen.Width - ScaleWidth) / 2
    
    ' Make sure that the SORTED Property of the list box is set to true, and the
    ' MULTISELECTED Property set to 2 - Extended, meaning you can multiselect by holding down SHIFT or CTRL key
    ' because you can't assign values to these properties at run time
    
    ' The other list box, named lstThings, has a multiselect property of 1 - Simple
    ' you'll know the difference between 1 and 2 once you run the program.
    
    'Initialize Items listbox
    With lstItems
        .AddItem "CPU"
        .AddItem "Monitor"
        .AddItem "Mouse"
        .AddItem "KeyBoard"
        .AddItem "Printer"
        .AddItem "Scanner"
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)  ' this is a pretty straight-forward thing
    Dim intEx As Integer
    
    intEx = MsgBox("Are you sure you want to exit?", vbYesNo + vbCritical + vbDefaultButton1, "Exit")
    If intEx = vbYes Then       ' vbYes is a VB constant that has an integer equivalent
                                ' that's the reason why we declared intEx as integer
        End
    Else
        Cancel = 1              ' set this to true so the form won't exit
    End If
End Sub

Private Sub txtAddItem_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then           ' <ENTER> Key was pressed
    If Trim(txtAddItem.Text) <> "" Then     ' check if textbox is empty
        lstItems.AddItem txtAddItem.Text    ' if not, add it to listbox
        txtAddItem.Text = ""                ' and set the textbox empty for another item
    Else
        MsgBox "Unable to add empty item.", vbOKOnly + vbCritical, "Add"
    End If
    txtAddItem.SetFocus
End If
End Sub
