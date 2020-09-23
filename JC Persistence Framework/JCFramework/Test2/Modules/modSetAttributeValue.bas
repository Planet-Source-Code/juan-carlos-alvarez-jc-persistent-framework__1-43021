Attribute VB_Name = "modSetAttributeValue"
Option Explicit

Public Sub setAttributeValue(ByVal objPers As Object, ByVal pName As String, ByRef Value As Variant)
  Select Case VarType(Value)
  Case vbInteger '=2
    CallByName objPers, pName, VbLet, CInt(Value)
  Case vbLong '=3
    CallByName objPers, pName, VbLet, CLng(Value)
  Case vbSingle '=4
    CallByName objPers, pName, VbLet, CSng(Value)
  Case vbDouble '=5
    CallByName objPers, pName, VbLet, CDbl(Value)
  Case vbCurrency '=6
    CallByName objPers, pName, VbLet, CCur(Value)
  Case vbDate '=7
    CallByName objPers, pName, VbLet, CDate(Value)
  Case vbString '=8
    CallByName objPers, pName, VbLet, CStr(Value)
  Case vbObject '=9
    CallByName objPers, pName, VbSet, Value
  Case vbError '=10
    'TO DO: Raise an exception
    MsgBox "Error in setAttributeValue::modSetAttributeValue"
  Case vbBoolean '=11
    CallByName objPers, pName, VbLet, CBool(Value)
  Case vbDecimal '=14
    CallByName objPers, pName, VbLet, CDec(Value)
  Case vbByte '=17
    CallByName objPers, pName, VbLet, CByte(Value)
  Case vbNull
    'Don´t do anything
    'No hacer nada
  Case Else
    'TO DO: Raise an exception
    MsgBox "Type not found, error in setAttributeValue::modSetAttributeValue"
  End Select
End Sub

