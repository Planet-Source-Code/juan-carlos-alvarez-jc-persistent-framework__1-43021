Attribute VB_Name = "modAliasNull"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3818C6D600DA"
Option Base 0
Option Explicit

'-----------------------------------------------------------
'VBPJ Junio 2000
'Titulo: Maintain Null Values
'SubTitulo:Ensure data integrity by maintaining database Null values properly.
'Autor: Eric Litwin
'Explicacion:
'   1-Usar "NullToAlias" para convertir valores de la base de datos nulos cuando cargamos los datos
'   2-Chequear si es un alias de valor nulo usando "IsAliasNull" cuando tengamos que mostrar o utilizar los valores en la interface
'   3-Convert UI control values to Alias Nulls when you set object property
'   4-Usar "AliasToNull" cuando grabamos los datos del objeto a la base de datos
'-----------------------------------------------------------


Public Const NULL_INTEGER = -32767 '-32768
Public Const NULL_LONG = -2147483647# '-2147483648#
Public Const NULL_SINGLE = -3.402823E+38
Public Const NULL_DOUBLE = -1.7976931348623E+308
Public Const NULL_CURRENCY = -922337203685477#
Public Const NULL_STRING = ""
Public Const NULL_DATE = #1/1/100#
Public Const NULL_BYTE = 0


'Constantes para devolver DATE y NOW
Public Const DEFAULT_DATE = #1/1/101#
Public Const DEFAULT_NOW = #1/1/101 1:01:01 AM#


'Returns true if the value (variant) is an aliased
'null number. For example, if the inValue is a date
'with a value of #01/01/100", this function will return
'true. If the value is #01/25/2000#, this function will return false.
Public Function IsNullAlias(ByVal InValue As Variant) As Boolean
    Dim blnValue As Boolean
    
    blnValue = False
    Select Case VarType(InValue)
        Case vbInteger
            If InValue = NULL_INTEGER Then blnValue = True
        Case vbLong
            If InValue = NULL_LONG Then blnValue = True
        Case vbSingle
            If InValue = NULL_SINGLE Then blnValue = True
        Case vbDouble
            If InValue = NULL_DOUBLE Then blnValue = True
        Case vbCurrency
            If InValue = NULL_CURRENCY Then blnValue = True
        Case vbString
            If Trim$(InValue) = NULL_STRING Then blnValue = True
        Case vbByte
            If InValue = NULL_BYTE Then blnValue = True
        Case Else
            ' bytes and boleans
            'Debug.Assert False
    End Select
    IsNullAlias = blnValue
End Function


' Replaces a Null value with an application defined null value.
' For example, instead of a Date being Null, an alias null
' would be #01/01/100#
' Primarily used when null values are returned from
' the database.  Returns the appropriate alias defined
' null value based on the variant type passed in.
' Can also force an override of this null value with a variant value.

'Reemplaza un valor NULL con un valor definido como nulo dentro de la aplicacion.
'Por ejemplo, en ves de que una fecha comience con NULL, podemos usar el alias para fecha nula #01/01/100#.
'Principalmente usado cuando retornamos valores NULL de una base de datos.
'Devuelve el alias apropiado al tipo de dato pasado.
'Tambien puede forzar sobreescribir este valor nulo con otro valor variant
Public Sub NullToAlias(ByVal InValue As Variant, _
                       ByRef OutValue As Variant, _
                       Optional Override As Variant = Null)
    Dim vntOutValue As Variant
    
    ' If the InValue is not Null, then just return that value
    If Not IsNull(InValue) Then
        vntOutValue = InValue
    ' Otherwise, check if we are to override the Null out value
    ' If not, then return a null alias specific for the variable type
    ElseIf IsNull(Override) Then
        Select Case VarType(OutValue)
            Case vbInteger
                vntOutValue = CInt(NULL_INTEGER)
            Case vbLong
                vntOutValue = CLng(NULL_LONG)
            Case vbSingle
                vntOutValue = CSng(NULL_SINGLE)
            Case vbDouble
                vntOutValue = CDbl(NULL_DOUBLE)
            Case vbCurrency
                vntOutValue = CCur(NULL_CURRENCY)
            Case vbString
                vntOutValue = CStr(NULL_STRING)
            Case vbDate
                vntOutValue = CDate(NULL_DATE)
            Case vbByte
                vntOutValue = CByte(NULL_BYTE)
            Case vbBoolean
                vntOutValue = CBool(False)
            Case Else
        End Select
    ' Otherwise, return the Null override value
    Else
        vntOutValue = Override
    End If
    
    OutValue = vntOutValue
End Sub





' Replaces an application defined null value with a
' true VB Null value. For example, a Date with a value
' of #01/01/100# would be replaces with 'Null'
' Primarily used when sending null data back to the
' database.  Can also override the VB null value with a variant value.
Public Function AliasToNull(ByVal InValue As Variant, Optional NullValue As Variant = Null) As Variant
    Dim vntOutValue As Variant
    
    vntOutValue = InValue
    Select Case VarType(InValue)
        Case vbInteger
            If InValue = NULL_INTEGER Then vntOutValue = NullValue
        Case vbLong
            If InValue = NULL_LONG Then vntOutValue = NullValue
        Case vbSingle
            If InValue = NULL_SINGLE Then vntOutValue = NullValue
        Case vbDouble
            If InValue = NULL_DOUBLE Then vntOutValue = NullValue
        Case vbCurrency
            If InValue = NULL_CURRENCY Then vntOutValue = NullValue
        Case vbString
            If Trim$(InValue) = NULL_STRING Then
                vntOutValue = NullValue
            Else
                vntOutValue = Trim$(InValue)
            End If
        Case vbDate
            If InValue = NULL_DATE Then vntOutValue = NullValue
        Case vbByte
            If InValue = NULL_BYTE Then vntOutValue = 0
        Case Else
'            Debug.Assert False
    End Select
    
    AliasToNull = vntOutValue
End Function


' Returns the number of decimals in a number.
Public Function NumberOfDecimals(InNumber As Variant) As Integer
    ' finds the length of the string after the first decimal
    Dim iPos As Integer
    Dim strTemp As String
    
    If IsNumeric(InNumber) Then
        strTemp = CStr(InNumber)
        iPos = InStr(1, Trim$(strTemp), ".")
        If iPos > 0 Then
            NumberOfDecimals = Len(strTemp) - iPos
        Else
            NumberOfDecimals = 0
        End If
    Else
        NumberOfDecimals = 0
    End If
End Function


' Returns the number of digits in a number (string value).
' Does not include decimals in the count.
Public Function NumberOfDigits(InNumber As Variant) As Integer
    ' finds the length of the string after the first decimal
    Dim iPos As Integer
    Dim strTemp As String
    Dim iNum As Integer
    
    If IsNumeric(InNumber) Then
        strTemp = CStr(InNumber)
        ' Only count digits to left of decimal point
        iPos = InStr(1, LTrim$(strTemp), ".")
        If iPos > 0 Then
            iNum = iPos - 1
        Else
            iNum = Len(Trim(strTemp))
        End If
        ' Do not count a negative sign
        iPos = InStr(1, LTrim$(strTemp), "-")
        If iPos > 0 Then
            iNum = iNum - 1
        End If
    Else
        iNum = 0
    End If
    
    NumberOfDigits = iNum
End Function
