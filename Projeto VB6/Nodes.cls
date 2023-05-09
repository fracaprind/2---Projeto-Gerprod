VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Nodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mobjNodes As Collection
Private mOwner As Node
Private mobjFirstNode As Node
Private mobjLastNode As Node
Private mlngKey As Long


Public Property Get FirstNode() As Node
    Set FirstNode = mobjFirstNode
End Property

Public Property Get LastNode() As Node
    Set LastNode = mobjLastNode
End Property

Public Property Get Item(ByVal Key As Variant) As Node
    On Error Resume Next
    Set Item = mobjNodes.Item(Key)
    If Err.Number <> 0 Then
        Err.Clear
        Set Item = Nothing
    End If
End Property

Public Property Set Owner(ByVal Value As Node)
    Set mOwner = Value
End Property

Public Function Add(ByVal Text As String) As Node
    mlngKey = mlngKey + 1
    Dim objNode As New Node
    objNode.Key = "N" & mlngKey
    objNode.Text = Text
    objNode.Level = mOwner.Level + 1
    Set objNode.Parent = mOwner
    If mobjNodes.Count = 0 Then
        Set mobjFirstNode = objNode
        Set mobjLastNode = objNode
    Else
        Set mobjLastNode.NextNode = objNode
        Set objNode.PrevNode = mobjLastNode
        Set mobjLastNode = objNode
    End If
    mobjNodes.Add objNode, objNode.Key

    mOwner.UpdateChindrenCount (1)
    If mOwner.Expanded Then
        mOwner.UpdateVisibleNodesCount (1)
    End If

    Set Add = objNode
End Function

Public Sub AddChildNodes(ByVal Count As Long)
    Dim i As Long
    For i = 1 To Count
        Call Add("")
    Next
End Sub

Public Sub Clear()
    Dim i As Long
    Dim objNode As Node

    If mobjNodes Is Nothing Then
        Exit Sub
    End If

    If mobjNodes.Count = 0 Then
        Exit Sub
    End If

    mOwner.UpdateChindrenCount (-mobjNodes.Count)
    If mOwner.Expanded Then
        mOwner.UpdateVisibleNodesCount (-mobjNodes.Count)
    End If

    Set mobjLastNode = Nothing
    Set mobjFirstNode = Nothing

    For i = mobjNodes.Count To 1 Step -1
        Set objNode = mobjNodes.Item(i)
        If objNode.HasChildren Then
            objNode.Nodes.Clear
        End If
        mobjNodes.Remove (i)
    Next
End Sub

Public Function Count() As Long
    Count = mobjNodes.Count
End Function

Public Function InsertBefore(ByVal Key As String, ByVal Text As String) As Node
    mlngKey = mlngKey + 1
    Dim objNode1 As Node
    Dim objNode2 As Node
    Dim objNode As New Node

    Set objNode1 = mobjNodes.Item(Key)
    Set objNode2 = objNode1.PrevNode
    objNode.Key = "N" & mlngKey
    objNode.Text = Text
    objNode.Level = mOwner.Level + 1
    Set objNode.Parent = mOwner
    
    Set objNode.NextNode = objNode1
    Set objNode.PrevNode = objNode2
    Set objNode1.PrevNode = objNode
    If Not objNode2 Is Nothing Then
        Set objNode2.NextNode = objNode
    Else
        Set mobjFirstNode = objNode
    End If
    mobjNodes.Add objNode, objNode.Key

    mOwner.UpdateChindrenCount (1)
    If mOwner.Expanded Then
        mOwner.UpdateVisibleNodesCount (1)
    End If

    Set InsertBefore = objNode
End Function

Public Function InsertAfter(ByVal Key As String, ByVal Text As String) As Node
    mlngKey = mlngKey + 1
    Dim objNode1 As Node
    Dim objNode2 As Node
    Dim objNode As New Node

    Set objNode1 = mobjNodes.Item(Key)
    Set objNode2 = objNode1.NextNode
    objNode.Key = "N" & mlngKey
    objNode.Text = Text
    objNode.Level = mOwner.Level + 1
    Set objNode.Parent = mOwner
    
    Set objNode.PrevNode = objNode1
    Set objNode.NextNode = objNode2
    Set objNode1.NextNode = objNode
    If Not objNode2 Is Nothing Then
        Set objNode2.PrevNode = objNode
    Else
        Set mobjLastNode = objNode
    End If
    mobjNodes.Add objNode, objNode.Key

    mOwner.UpdateChindrenCount (1)
    If mOwner.Expanded Then
        mOwner.UpdateVisibleNodesCount (1)
    End If

    Set InsertAfter = objNode
End Function

Public Sub Remove(ByVal Key As Variant)
    Dim objNode As Node
    Set objNode = mobjNodes.Item(Key)

    If mobjNodes.Count <= 1 Then
        Set mobjFirstNode = Nothing
        Set mobjLastNode = Nothing
    Else
        If objNode.PrevNode Is Nothing Then    'FirstNode
            Set mobjFirstNode = objNode.NextNode
            Set mobjFirstNode.PrevNode = Nothing
        Else
            Set objNode.PrevNode.NextNode = objNode.NextNode
        End If
        If objNode.NextNode Is Nothing Then    'LastNode
            Set mobjLastNode = objNode.PrevNode
            Set mobjLastNode.NextNode = Nothing
        Else
            Set objNode.NextNode.PrevNode = objNode.PrevNode
        End If
    End If

    If objNode.HasChildren Then
        objNode.Nodes.Clear
    End If

    mobjNodes.Remove (Key)
    mOwner.UpdateChindrenCount (-1)
    If mOwner.Expanded Then
        mOwner.UpdateVisibleNodesCount (-1)
    End If
End Sub

Private Sub Class_Initialize()
    Set mobjNodes = New Collection
End Sub

Private Sub Class_Terminate()
    Clear
    Set mobjNodes = Nothing
End Sub

