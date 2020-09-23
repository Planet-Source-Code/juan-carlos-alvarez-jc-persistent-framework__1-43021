VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CXMLConfigLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'**************************************************
'Class: CXMLConfigLoader
'Author: Juan Carlos Alvarez
'**************************************************

Implements IConfigLoader

Private m_FileName As String
Private m_classMap As Scripting.Dictionary
Private m_databaseMap As Scripting.Dictionary

Public Property Get FileName() As String
  FileName = m_FileName
End Property

Public Property Let FileName(ByVal file As String)
  m_FileName = file
End Property

Public Property Get ClassMap() As Scripting.Dictionary
  Set ClassMap = m_classMap
End Property

Public Property Set ClassMap(ByVal clm As Scripting.Dictionary)
  Set m_classMap = clm
End Property

Public Property Get DatabaseMap() As Scripting.Dictionary
  Set DatabaseMap = m_databaseMap
End Property

Public Property Set DatabaseMap(ByVal dbm As Scripting.Dictionary)
  Set m_databaseMap = dbm
End Property

Private Sub Class_Initialize()
  Dim iniFile As IniFiles
  Set iniFile = New IniFiles
  iniFile.Section = "Xml Config File"
  iniFile.FilePath = App.Path + "\IniFiles\"
  iniFile.FileName = "xmlpath.ini"
  m_FileName = App.Path + iniFile.getProperty("filePath")
  Set iniFile = Nothing
End Sub

Private Sub IConfigLoader_loadConfig(ByRef dbMap As Scripting.IDictionary, clMap As Scripting.IDictionary)
  Set m_classMap = clMap
  Set m_databaseMap = dbMap
  Dim doc As New DOMDocument
  doc.Load (Me.FileName)
  Dim elementRoot As IXMLDOMElement
  Set elementRoot = doc.documentElement
  Dim node As IXMLDOMNode
  Set node = elementRoot.firstChild
  Do While Not node Is Nothing
    If node.nodeType = NODE_ELEMENT Then
      Dim elementChild As IXMLDOMElement
      Set elementChild = node
      If elementChild.tagName = "database" Then
        Dim relDb As CRelationalDatabase
        Set relDb = getRelationalDatabase(elementChild)
        dbMap.Add relDb.Name, relDb
      ElseIf elementChild.tagName = "class" Then
        Dim clm As CClassMap
        Set clm = getClassMap(elementChild)
        clMap.Add clm.Name, clm
      ElseIf elementChild.tagName = "association" Then
        processAssociation elementChild
      End If
    End If
    Set node = node.nextSibling
  Loop
End Sub

Private Function getRelationalDatabase(ByVal node As IXMLDOMElement) As CRelationalDatabase
  Dim relDb As CRelationalDatabase
  Set relDb = Nothing
  Dim attrPMName, attrClassName As IXMLDOMAttribute
  Set attrPMName = node.getAttributeNode("name")
  Set attrClassName = node.getAttributeNode("class")
  If ((Not attrPMName Is Nothing) And (Not attrClassName Is Nothing)) Then
    Set relDb = CreateObject("JCFramework." + attrClassName.Value)
    relDb.Name = attrPMName.Value
    Dim params As Scripting.Dictionary
    Set params = CreateObject("Scripting.Dictionary")
    Dim nodeChild As IXMLDOMNode
    Set nodeChild = node.firstChild
    Do While Not nodeChild Is Nothing
      If nodeChild.nodeType = NODE_ELEMENT Then
        Dim elementChild As IXMLDOMElement
        Set elementChild = nodeChild
        If elementChild.tagName = "parameter" Then
          Dim attrName As IXMLDOMAttribute, attrValue As IXMLDOMAttribute
          Set attrName = elementChild.getAttributeNode("name")
          Set attrValue = elementChild.getAttributeNode("value")
          If ((Not attrName Is Nothing) And (Not attrValue Is Nothing)) Then
            params.Add attrName.nodeValue, attrValue.Value
          End If
        End If
      End If
      Set nodeChild = nodeChild.nextSibling
    Loop
    relDb.init params
  End If
  Set getRelationalDatabase = relDb
End Function

Private Function getClassMap(ByVal node As IXMLDOMElement) As CClassMap
  Dim attrClassName As IXMLDOMAttribute, attrTable As IXMLDOMAttribute
  Dim attrDatabase As IXMLDOMAttribute, attrSuperClassName As IXMLDOMAttribute
  Set attrClassName = node.getAttributeNode("name")
  Set attrTable = node.getAttributeNode("table")
  Set attrDatabase = node.getAttributeNode("database")
  Set attrSuperClassName = node.getAttributeNode("superclass")
  If ((Not attrClassName Is Nothing) And (Not attrDatabase Is Nothing)) Then
    'Creo el ClassMap actual
    Dim ClassMap As CClassMap
    Set ClassMap = New CClassMap
    ClassMap.Name = attrClassName.Value
    If Not attrSuperClassName Is Nothing Then
      Set ClassMap.SuperClass = m_classMap.Item(attrSuperClassName.Value)
    End If
    'Creo el DatabaseMap para la tabla actual
    Dim dbMap As CDatabaseMap
    Set dbMap = New CDatabaseMap
    dbMap.Name = attrDatabase.Value
    'Creo la TableMap
    Dim tblMap As CTableMap
    Set tblMap = New CTableMap
    tblMap.Name = attrTable.Value
    tblMap.DatabaseMap = dbMap
    
    Dim nodeChild As IXMLDOMNode
    Set nodeChild = node.firstChild
    Do While Not nodeChild Is Nothing
      If nodeChild.nodeType = NODE_ELEMENT Then
        Dim elementChild As IXMLDOMElement
        Set elementChild = nodeChild
        If elementChild.tagName = "attribute" Then
          Dim AttrMap As CAttributeMap
          Set AttrMap = getAttributeMap(elementChild, ClassMap, tblMap)
          ClassMap.addAttributeMap AttrMap
        End If
      End If
      Set nodeChild = nodeChild.nextSibling
    Loop
  End If
  Set getClassMap = ClassMap
End Function

Private Function getAttributeMap(ByVal node As IXMLDOMElement, ByVal clMap As CClassMap, ByVal tblMap As CTableMap) As CAttributeMap
  Dim attrName, attrColumn, attrKey, attrReference As IXMLDOMAttribute
  Set attrName = node.getAttributeNode("name")
  Set attrColumn = node.getAttributeNode("column")
  Set attrKey = node.getAttributeNode("key")
  Set attrReference = node.getAttributeNode("reference")
  If Not attrName Is Nothing Then
    'Creo el column map para este atributo
    Dim colMap As CColumnMap
    Set colMap = Nothing
    If Not attrColumn Is Nothing Then
      Set colMap = New CColumnMap
      If Not attrKey Is Nothing Then
        If UCase(attrKey.Value) = "PRIMARY" Then
          colMap.KeyType = ColumnKeyTypes.PRIMARY_KEY
        ElseIf UCase(attrKey.Value) = "FOREIGN" Then
          colMap.KeyType = ColumnKeyTypes.FOREIGN_KEY
        End If
      End If
      colMap.Name = attrColumn.Value
      Set colMap.TableMap = tblMap
    End If
    'Creo el attribute map
    Dim AttrMap As CAttributeMap
    Set AttrMap = New CAttributeMap
    AttrMap.Name = attrName.Value
    Set AttrMap.ColumnMap = colMap
    
    If ((Not attrReference Is Nothing) And (Not clMap.SuperClass Is Nothing)) Then
      Dim refAttr As CAttributeMap
      Set refAttr = clMap.SuperClass.getAttributeMapByString(attrReference.Value, True)
      If Not refAttr Is Nothing Then
        Set AttrMap.AttributeMap = refAttr
      End If
    End If
    Set getAttributeMap = AttrMap
  End If
End Function

Private Sub processAssociation(ByVal node As IXMLDOMElement)
  Dim attrFromClass, attrToClass, attrTarget, attrCardinality As IXMLDOMAttribute
  Set attrFromClass = node.getAttributeNode("fromClass")
  Set attrToClass = node.getAttributeNode("toClass")
  Set attrTarget = node.getAttributeNode("target")
  Set attrCardinality = node.getAttributeNode("cardinality")
  If ((Not attrFromClass Is Nothing) And _
      (Not attrToClass Is Nothing) And _
      (Not attrTarget Is Nothing)) Then
    Dim fromClassMap As CClassMap
    Set fromClassMap = m_classMap.Item(attrFromClass.Value)
    If fromClassMap Is Nothing Then
      '// ** TO DO: Raise an exception here
      MsgBox "Error in association definition: fromClassMap = " + attrFromClass.Value
    End If
    Dim toClassMap As CClassMap
    Set toClassMap = m_classMap.Item(attrToClass.Value)
    If toClassMap Is Nothing Then
      '// ** TO DO: Raise an exception here
      MsgBox "Error in association definition: toClassMap = " + attrToClass.Value
    End If
    Dim udaAm As CUDAMap
    Set udaAm = New CUDAMap
    Set udaAm.ForClass = toClassMap
    Set udaAm.Target = fromClassMap.getAttributeMapByString(attrTarget.Value, False)
    udaAm.DeleteAutomatic = node.getAttributeNode("deleteAutomatic").Value
    udaAm.SaveAutomatic = node.getAttributeNode("saveAutomatic").Value
    udaAm.RetrieveAutomatic = node.getAttributeNode("retrieveAutomatic").Value
    udaAm.Inverse = node.getAttributeNode("inverse").Value
    If Not attrCardinality Is Nothing Then
      If UCase(attrCardinality.Value) = "ONETOONE" Then
        udaAm.Cardinality = CardinalityEnum.ONE_TO_ONE
      ElseIf UCase(attrCardinality.Value) = "ONETOMANY" Then
        udaAm.Cardinality = CardinalityEnum.ONE_TO_MANY
      End If
    End If
    
    Dim nodeChild As IXMLDOMNode
    Set nodeChild = node.firstChild
    Do While Not nodeChild Is Nothing
      If nodeChild.nodeType = NODE_ELEMENT Then
        Dim elementChild As IXMLDOMElement
        Set elementChild = nodeChild
        If elementChild.tagName = "entry" Then
          Dim attrFromAttribute, attrToAttribute As IXMLDOMAttribute
          Set attrFromAttribute = elementChild.getAttributeNode("fromAttribute")
          Set attrToAttribute = elementChild.getAttributeNode("toAttribute")
          If ((Not attrFromAttribute Is Nothing) And (Not attrToAttribute Is Nothing)) Then
            Dim amFromAttribute, amToAttribute As CAttributeMap
            Set amFromAttribute = fromClassMap.getAttributeMapByString(attrFromAttribute.Value, False)
            Set amToAttribute = toClassMap.getAttributeMapByString(attrToAttribute.Value, False)
            If amFromAttribute Is Nothing Then
              '// ** TO DO: Raise an exception here
              MsgBox "Error in association definition"
            End If
            If amToAttribute Is Nothing Then
              '// ** TO DO: Raise an exception here
              MsgBox "Error in association definition"
            End If
            Dim entry As CUDAMapEntry
            Set entry = New CUDAMapEntry
            If udaAm.Inverse Then
              Set entry.FromAttrMap = amFromAttribute
              Set entry.ToAttrMap = amFromAttribute
            Else
              Set entry.FromAttrMap = amFromAttribute
              Set entry.ToAttrMap = amToAttribute
            End If
            udaAm.addEntry entry
          End If
        End If
      End If
      Set nodeChild = nodeChild.nextSibling
    Loop
    fromClassMap.addAssociationMap udaAm
  End If
End Sub
