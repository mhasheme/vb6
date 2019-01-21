VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FieldInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarfdName As String 'local copy
Private mvarfdTable As String 'local copy
Private mvarfdDesc As String 'local copy
Private mvarfdValue As Variant 'local copy
Private mvarfdLength As Integer 'local copy
Private mvarfdType As FieldTypeEnum
Private mvarfdDecimal As Integer 'local copy
Private mvarfdIHR As String 'local copy
Private mvarfdREQ As Boolean
Public Property Let fdDecimal(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fDecimal = 5
    mvarfdDecimal = vData
End Property


Public Property Get fdDecimal() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fDecimal
    fdDecimal = mvarfdDecimal
End Property
'Public Property Let fdAUDIT(ByVal vData As String)
'    mvarfdAUDIT = vData
'End Property
'
'
'Public Property Get fdAUDIT() As String
'    fdAUDIT = mvarfdAUDIT
'End Property
Public Property Let fdIHR(ByVal vData As String)
    mvarfdIHR = vData
End Property


Public Property Get fdIHR() As String
    fdIHR = mvarfdIHR
End Property

Public Property Let fdREQ(ByVal vData As Boolean)
    mvarfdREQ = vData
End Property


Public Property Get fdREQ() As Boolean
    fdREQ = mvarfdREQ
End Property
Public Property Let fdType(ByVal vData As FieldTypeEnum)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fType = 5
    mvarfdType = vData
End Property


Public Property Get fdType() As FieldTypeEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fType
    fdType = mvarfdType
End Property



Public Property Let fdLength(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fLength = 5
    mvarfdLength = vData
End Property


Public Property Get fdLength() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fLength
    fdLength = mvarfdLength
End Property



Public Property Let fdValue(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fValue = 5
    mvarfdValue = vData
End Property


Public Property Set fdValue(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.fValue = Form1
    Set mvarfdValue = vData
End Property


Public Property Get fdValue() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fValue
    If IsObject(mvarfdValue) Then
        Set fdValue = mvarfdValue
    Else
        fdValue = mvarfdValue
    End If
End Property



Public Property Let fdName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fName = 5
    mvarfdName = vData
End Property


Public Property Get fdName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fName
    fdName = mvarfdName
End Property


Public Property Let fdTable(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fName = 5
    mvarfdTable = vData
End Property


Public Property Get fdTable() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fName
    fdTable = mvarfdTable
End Property

Public Property Let fdDesc(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fName = 5
    mvarfdDesc = vData
End Property


Public Property Get fdDesc() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fName
    fdDesc = mvarfdDesc
End Property
