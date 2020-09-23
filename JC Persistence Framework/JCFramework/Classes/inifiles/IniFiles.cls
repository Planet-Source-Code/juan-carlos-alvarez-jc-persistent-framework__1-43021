VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IniFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Variables de instancia
Private m_FileName As String  'indica el nombre del archivo ini a manejar
Private m_Section As String   'indica la seccion del archivo ini
Private m_FilePath As String  'indica el path del archivo ini

' APIs para manejar archivos ini
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal SectionName As String, ByVal KeyName As String, ByVal Default As Long, ByVal FileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal SectionName As String, ByVal KeyName As String, ByVal Default As String, ByVal ReturnedString As String, ByVal StringSize As Long, ByVal FileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal SectionName As String, ByVal KeyName As String, ByVal KeyValue As String, ByVal FileName As String) As Long

' Definiciones para manejar posibles errores
Private Const mklErrFileNotIni As String = "The field FileName should be setted"
Private Const mklErrFileNotExist As String = "FileName not exist"
Private Const mklErrOffset As Long = vbObjectError + 512
Private Const mklErrNotNumeric As Long = mklErrOffset + 1
Private Const mksErrNotNumeric As String = "The setting value returned was not numeric"
Private Const mklErrWriting As Long = mklErrOffset + 2
Private Const mksErrWriting As String = "Error writing Setting "
Private Const mklErrReading As Long = mklErrOffset + 3
Private Const mksErrReading As String = "Error reading Setting "
Private Const mklErrSection As Long = mklErrOffset + 4
Private Const mksErrSection As String = "Error reading Section "
Private Const mklErrDeleteKey As Long = mklErrOffset + 5
Private Const mksErrDeleteKey As String = "Error Deleting Key "
Private Const mklErrDeleteSection As Long = mklErrOffset + 6
Private Const mksErrDeleteSection As String = "Error Deleting Section "

'Set y Get del atributo FileName
Public Property Let FileName(ByVal strFileName As String)
    CheckFileExist
    m_FileName = strFileName
End Property

Public Property Get FileName() As String
    FileName = m_FileName
End Property

'Set y Get del atributo Seccion
Public Property Let Section(ByVal strSection As String)
    m_Section = strSection
End Property

Public Property Get Section() As String
    Section = m_Section
End Property

'Set y Get del atributo FilePath
Public Property Let FilePath(ByVal strFilePath As String)
    m_FilePath = strFilePath
End Property

Public Property Get FilePath() As String
    FilePath = m_FilePath
End Property

'Esta funcion permite obtener el valor de un property dado
Public Function getProperty(ByVal strKeyName As String) As String
    Dim lRet As Long
    Dim sBuf As String * 128
    Dim strDefaultValue As String
    On Error GoTo getPropertyError
    
    strDefaultValue = "VALUE NOT FOUND"
    lRet = GetPrivateProfileString(Me.Section, strKeyName, strDefaultValue, sBuf, Len(sBuf), Me.FilePath + Me.FileName)
    getProperty = TrimNull(sBuf)
    Exit Function
    
getPropertyError:
    Err.Raise mklErrReading, "Property.getProperty", mksErrReading & "(" & Err & ", " & Error & ")"
End Function

Public Function getAllBySection(ByVal strSection As String) As String()
    'Retorna un array de Strings con todos los valores de una seccion
    'Las claves deben ser numericas en orden ascendente comenzando en 1
    Const CANT_ELEM_ARRAY = 10
    Dim vArray() As String
    Dim sBuf As String * 128
    Dim i As Integer, iter As Integer
    Dim lRet As Long
    Dim strValorPorDefecto As String
    Dim strValorAGuardar As String
    On Error GoTo getAllBySectionError
        
    'ReDim vArray(cantClaves) As String     'make variant array
    strValorPorDefecto = "" 'Esto hace que lRet sea 0 si la key no se encuentra
    i = 0
    iter = 1 'Lo que tiene vArray menos 1
    lRet = GetPrivateProfileString(strSection, Str(i + 1), strValorPorDefecto, sBuf, Len(sBuf), Me.FilePath + Me.FileName)
    strValorAGuardar = TrimNull(sBuf)
    ReDim Preserve vArray((CANT_ELEM_ARRAY * iter) - 1)
    While lRet > 0
        If (i > (CANT_ELEM_ARRAY * iter) - 1) Then
            iter = iter + 1
            ReDim Preserve vArray((CANT_ELEM_ARRAY * iter) - 1)
        End If
        vArray(i) = strValorAGuardar
        i = i + 1
        lRet = GetPrivateProfileString(strSection, Str(i + 1), strValorPorDefecto, sBuf, Len(sBuf), Me.FilePath + Me.FileName)
        strValorAGuardar = TrimNull(sBuf)
    Wend
    ReDim Preserve vArray(i - 1)
    getAllBySection = vArray      'return array
    Exit Function
   
getAllBySectionError:
    Err.Raise mklErrSection, "Property.getAllBySection", mksErrSection & "(" & Err & ", " & Error & ")"
End Function

Public Function getSectionNames() As Variant
    Dim vArray() As Variant
    Dim sKeys() As String
    Dim lRet As Long, lCnt As Long
    Dim lKeys As Long, lKey As Long
    Dim sBuf As String * 2048
    On Error GoTo getSectionNamesError
    
    'list section names
    lRet = GetPrivateProfileString(0&, 0&, 0&, sBuf, Len(sBuf), Me.FilePath + Me.FileName)
    If (lRet > 1) Then              'if any returned
        lCnt = InStr(sBuf, vbNullChar & vbNullChar)     'find end of list
        sKeys = Split(Left$(sBuf, lCnt - 1), vbNullChar) 'separate keys
        lKeys = UBound(sKeys)       'get count
        ReDim vArray(lKeys)         'make variant array
        For lKey = 0 To lKeys       'fill array
            vArray(lKey) = sKeys(lKey)
        Next
        getSectionNames = vArray      'return array
    End If
    Exit Function
    
getSectionNamesError:
    Err.Raise mklErrSection, "Property.getSectionNames", mksErrSection & "(" & Err & ", " & Error & ")"
End Function

Public Sub setProperty(ByVal strKeyName As String, ByVal strSetValue As String)
    'Guarda strSetValue como valor de strKeyName
    Dim lRet As Long
    On Error GoTo setPropertyError
    
    lRet = WritePrivateProfileString(Me.Section, strKeyName, strSetValue, Me.FilePath + Me.FileName)
    Exit Sub
    
setPropertyError:
    Err.Raise mklErrWriting, "Property.setProperty", mksErrWriting & "(" & Err & ", " & Error & ")"
End Sub

Public Sub deleteKey(ByVal KeyName As String)
    'Borra la clave KeyName
    Dim lRet As Long
    On Error GoTo deleteKeyError
    
    lRet = WritePrivateProfileString(Me.Section, KeyName, 0&, Me.FilePath + Me.FileName)
    If lRet = 0 Then Err.Raise mklErrDeleteKey, "Property.DeleteKey", mksErrDeleteKey & "(check if file is protected)"
    Exit Sub
    
deleteKeyError:
    Err.Raise mklErrDeleteKey, "Property.deleteKey", mksErrDeleteKey & "(" & Err & ", " & Error & ")"
End Sub

Private Function TrimNull(ByVal InString As String) As String
    'trims string at first Null character
    Dim lPos As Long
    TrimNull = Trim$(InString)
    lPos = InStr(TrimNull, vbNullChar)
    If lPos > 0 Then TrimNull = Left$(TrimNull, lPos - 1)
End Function

Private Function FileExist(ByVal strFileName As String) As Boolean
    On Local Error Resume Next
    FileExist = (Dir(strFileName) <> "")
End Function

Private Function PathExist(ByVal strPathName As String) As Boolean
    On Local Error Resume Next
    PathExist = (Dir(strPathName) <> "")
End Function

Private Sub CheckPathExist()
    If Not FileExist(Me.FilePath + Me.FileName) Then
        MsgBox "File: " + Me.FilePath + Me.FileName + " not found!!!"
    End If
End Sub

Private Sub CheckFileExist()
    If Not FileExist(Me.FilePath + Me.FileName) Then
        MsgBox "File: " + Me.FilePath + Me.FileName + " not found!!!"
    End If
End Sub

