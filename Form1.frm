VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "ListVew Drag And Drop"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   6480
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   720
      Width           =   540
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Symbol"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Move ListView Items with Mouse
'by Matt Jonker
'mattjonker@usa.net


' Declare global variables.
Dim indrag As Boolean ' Flag that signals a Drag Drop operation.
Dim selX As Object ' Item that is being dragged.
Public LastItemIndex As String
Private MousePosX As Single
Private MousePosY As Single

Private Sub Form_Load()
    Dim itmX As ListItem    ' Create a tree.
    Set itmX = ListView1.ListItems.Add(1, , "SYKE")
    Set itmX = ListView1.ListItems.Add(2, , "KO")
    Set itmX = ListView1.ListItems.Add(3, , "ATMI")
    Set itmX = ListView1.ListItems.Add(4, , "ADI")
    Set itmX = ListView1.ListItems.Add(5, , "CSCO")
    Set itmX = ListView1.ListItems.Add(6, , "NOVL")
    Set itmX = ListView1.ListItems.Add(7, , "MSFT")
    Set itmX = ListView1.ListItems.Add(8, , "AOL")
    Set itmX = ListView1.ListItems.Add(9, , "NSCP")
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set selX = ListView1.SelectedItem
    If Button = vbLeftButton And ListView1.ListItems.Count <> 0 Then ' Signal a Drag operation.
        indrag = True ' Set the flag to true.
        LastItemIndex = ListView1.SelectedItem.Index
        ' Set the drag icon with the CreateDragImage method.
        ListView1.DragIcon = Picture1.Picture
        ListView1.Drag vbBeginDrag ' Drag operation.
    End If
End Sub

Private Sub ListView1_DragDrop(Source As Control, x As Single, y As Single)
    If ListView1.DropHighlight Is Nothing Then
        Set ListView1.DropHighlight = Nothing
        indrag = False
        Exit Sub
    Else
        If selX = ListView1.DropHighlight Then Exit Sub
        Debug.Print selX.Text & " dropped on " & ListView1.DropHighlight.Text
        Set ListView1.DropHighlight = Nothing
        indrag = False
End If
End Sub

Private Sub ListView1_DragOver(Source As Control, x As Single, y As Single, State As Integer)

MousePosX = x
MousePosY = y

    If indrag = True Then
        ' Set DropHighlight to the mouse's coordinates.
        
        If MousePosX < 0 Then
            MousePosX = 0
        End If
        
        If MousePosX > ListView1.Width Then
            MousePosX = ListView1.Width
        End If
        
        If MousePosY < ListView1.ListItems(1).Top Then
            MousePosY = ListView1.ListItems(1).Top
        End If
        
        If MousePosY > (ListView1.ListItems(1).Height * ListView1.ListItems.Count + 100) Then
            MousePosY = (ListView1.ListItems(1).Height * ListView1.ListItems.Count + 100)
        End If
               
        If ListView1.SelectedItem.Index <> LastItemIndex Then
            If ListView1.HitTest(40, MousePosY).Index < ListView1.SelectedItem.Index Then
                Set itmX = ListView1.ListItems.Add(ListView1.SelectedItem.Index - 1, , ListView1.SelectedItem.Text)
                If ListView1.SelectedItem.Index = ListView1.ListItems.Count Then
                    ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
                    Set ListView1.SelectedItem = ListView1.ListItems(ListView1.SelectedItem.Index - 1)
                Else
                    ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
                    Set ListView1.SelectedItem = ListView1.ListItems(ListView1.SelectedItem.Index - 2)
                End If
            Else
                Set itmX = ListView1.ListItems.Add(ListView1.SelectedItem.Index + 2, , ListView1.SelectedItem.Text)
                ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
                Set ListView1.SelectedItem = ListView1.ListItems(ListView1.SelectedItem.Index + 1)
            End If
        End If
        LastItemIndex = ListView1.HitTest(ListView1.ListItems(1).Left, MousePosY).Index
    End If
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    Set ListView1.DropHighlight = Nothing
    
End Sub
