Option Compare Database
Option Explicit

Public Function SetWebBrowserControlSource()
    SetWebBrowserControlSource = "file://C:\ae\Association-Contacts\themes\ae\images\aelogo_256x256.svg"
End Function

Public Function ReadJSON()
 
    Dim root As Object
    Dim content As String
    Dim rootKeys() As String
    Dim keys() As String
    Dim i As Integer
    Dim obj As Object
    Dim prop As Variant
    
    content = basFileSys.FileToString(CurrentProject.Path & "\example0.json")
    
    content = Replace(content, vbCrLf, "")
    content = Replace(content, vbTab, "")
 
    basJsonParser.InitScriptEngine
 
    Set root = basJsonParser.DecodeJsonString(content)
  
    rootKeys = basJsonParser.GetKeys(root)
    
    For i = 0 To UBound(rootKeys)
    
        Debug.Print rootKeys(i)
        
        If basJsonParser.GetPropertyType(root, rootKeys(i)) = jptValue Then
            prop = basJsonParser.GetProperty(root, rootKeys(i))
            Debug.Print Nz(prop, "[null]")
        Else
            Set obj = basJsonParser.GetObjectProperty(root, rootKeys(i))
            RecurseProps obj, 2
        End If
        
    Next i
 
End Function

Private Function RecurseProps(obj As Object, Optional Indent As Integer = 0) As Object

    Dim nextObject As Object
    Dim propValue As Variant
    Dim keys() As String
    Dim i As Integer
    
    keys = basJsonParser.GetKeys(obj)
    
    For i = 0 To UBound(keys)
        
        If basJsonParser.GetPropertyType(obj, keys(i)) = jptValue Then
            propValue = basJsonParser.GetProperty(obj, keys(i))
            Debug.Print Space(Indent) & keys(i) & ": " & Nz(propValue, "[null]")
        Else
            Set nextObject = basJsonParser.GetObjectProperty(obj, keys(i))
            Debug.Print Space(Indent) & keys(i)
            RecurseProps nextObject, Indent + 2
        End If
    
    Next i
    
End Function