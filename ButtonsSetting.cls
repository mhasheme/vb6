VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ButtonsSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public Property Let Visible(bTypeOrKey As String, ByVal vData As Boolean)
Dim xButton As Object
Dim xType As String, xKey As String
Dim xMenu As Menu
If InStr(bTypeOrKey, "Type:") <> 0 Then
    xType = Mid(bTypeOrKey, 6)
Else
    xKey = bTypeOrKey
End If
For Each xButton In MDIMain.MainToolBar.ButtonS
    If Len(xType) = 0 Then
        If xButton.Key = xKey Then
            xButton.Visible = vData
            If Not xKey = "help" And Not xKey = "hrsoft" And Not xKey = "NewContractEmp" Then
                Set xMenu = LinkMenu(xButton.Key)
                xMenu.Visible = vData
            End If
        End If
    Else
        If xButton.Description = xType Then
            xButton.Visible = vData
            If xButton.Key <> "" Then
                Set xMenu = LinkMenu(xButton.Key)
                If xType = "move" Then
                    Set xMenu = MDIMain.mnuMove
                ElseIf xType = "mass" Then
                    Set xMenu = MDIMain.mnu_Mass
                ElseIf xType = "edit" Then
                    Set xMenu = MDIMain.mnuEdit
                End If
                xMenu.Visible = vData
            End If
        End If
    End If
Next
End Property

Public Property Let Enabled(bTypeOrKey As String, ByVal vData As Boolean)
Dim xButton As Object
Dim xType As String, xKey As String
Dim xMenu As Menu
If InStr(bTypeOrKey, "Type:") <> 0 Then
    xType = Mid(bTypeOrKey, 6)
Else
    xKey = bTypeOrKey
End If

For Each xButton In MDIMain.MainToolBar.ButtonS
    If Len(xType) = 0 Then
        If xButton.Key = xKey Then
            'Ticket #22682 - Release 8.0
            If xKey = "NewEmployee" Then
                If gSec_Add_NewHire Then
                    xButton.Enabled = vData
                    Set xMenu = LinkMenu(xButton.Key)
                    xMenu.Enabled = vData
                End If
            Else
                xButton.Enabled = vData
                Set xMenu = LinkMenu(xButton.Key)
                xMenu.Enabled = vData
            End If
        End If
    Else
        If xButton.Description = xType Then
            xButton.Enabled = vData
            Set xMenu = LinkMenu(xButton.Key)
            If xButton.Key <> "" Then xMenu.Enabled = vData
        End If
    End If
Next
End Property


Private Function LinkMenu(bKey) As Menu

    Select Case bKey
    Case "close":           Set LinkMenu = MDIMain.mnu_F_Close
    
    Case "preview":         Set LinkMenu = MDIMain.mnu_F_Preview
    Case "print":           Set LinkMenu = MDIMain.mnu_F_Print
    
    Case "NewEmployee":    Set LinkMenu = MDIMain.mnu_NewEmployee
    
    Case "NewRecord":       Set LinkMenu = MDIMain.mnu_E_NewRecord
    Case "save":            Set LinkMenu = MDIMain.mnu_E_Save
    Case "cancel":          Set LinkMenu = MDIMain.mnu_E_Cancel
    Case "delete":          Set LinkMenu = MDIMain.mnu_E_Delete
    
    Case "down":            Set LinkMenu = MDIMain.mnu_M_Down
    Case "up":              Set LinkMenu = MDIMain.mnu_M_Up
    
    Case "FindEmployee":    Set LinkMenu = MDIMain.mnu_F_Employee
    Case "FindTerm":        Set LinkMenu = MDIMain.mnu_F_TEmployee
'    Case "FindREQ":         Set LinkMenu = MDIMain.mnu_F_OpenREQ
    Case "FindPosition":    Set LinkMenu = MDIMain.mnu_F_Position
    
    Case "massadd":         Set LinkMenu = MDIMain.mnu_M_Add
    Case "massupdate":      Set LinkMenu = MDIMain.mnu_M_Update
    Case "massdelete":      Set LinkMenu = MDIMain.mnu_M_Delete
    Case "mail":            Set LinkMenu = MDIMain.mmnu_mail
End Select
End Function







