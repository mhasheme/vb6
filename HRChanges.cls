VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HRChanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarOldValue As Variant 'local copy
Private mvarNewValue As Variant 'local copy
Private mvarHRField As Variant 'local copy
Public Property Let HRField(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HRField = 5
    mvarHRField = vData
End Property


Public Property Set HRField(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.HRField = Form1
    Set mvarHRField = vData
End Property


Public Property Get HRField() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HRField
    If IsObject(mvarHRField) Then
        Set HRField = mvarHRField
    Else
        HRField = mvarHRField
    End If
End Property



Public Property Let NewValue(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NewValue = 5
    mvarNewValue = vData
End Property


Public Property Set NewValue(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.NewValue = Form1
    Set mvarNewValue = vData
End Property


Public Property Get NewValue() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NewValue
    If IsObject(mvarNewValue) Then
        Set NewValue = mvarNewValue
    Else
        NewValue = mvarNewValue
    End If
End Property



Public Property Let OldValue(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OldValue = 5
    mvarOldValue = vData
End Property


Public Property Set OldValue(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.OldValue = Form1
    Set mvarOldValue = vData
End Property


Public Property Get OldValue() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OldValue
    If IsObject(mvarOldValue) Then
        Set OldValue = mvarOldValue
    Else
        OldValue = mvarOldValue
    End If
End Property



