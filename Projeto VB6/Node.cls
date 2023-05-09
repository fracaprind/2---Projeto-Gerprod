VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Node"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstrKey As String
Private mstrText As String
Private mblnExpanded As Boolean
Private mlngChildrenCount As Long
Private mlngVisibleNodesCount As Long
Private mlngLevel As Long
Private mobjNodes As Nodes
Private mobjParent As Node
Private mobjPrevNode As Node
Private mobjNextNode As Node

Friend Property Get AbsIndex() As Long
    Dim i As Long
    Dim tmpNode As Node
    Set tmpNode = Me
    i = 1
    Do
        If tmpNode.PrevNode Is Nothing Then
            If Not mobjParent Is Nothing Then
                i = i + mobjParent.AbsIndex
            End If
            Exit Do
        Else
            Set tmpNode = tmpNode.PrevNode
            i = i + 1 + tmpNode.ChildrenCount
        End If
    Loop
    AbsIndex = i
End Property

Friend Property Get PrevNode() As Node
    Set PrevNode = mobjPrevNode
End Property
Friend Property Set PrevNode(ByVal Value As Node)
    Set mobjPrevNode = Value
End Property

Friend Property Get NextNode() As Node
    Set NextNode = mobjNextNode
End Property
Friend Property Set NextNode(ByVal Value As Node)
    Set mobjNextNode = Value
End Property

Friend Property Get Root() As Node
    If mobjParent Is Nothing Then
        Set Root = Me
    Else
        Set Root = mobjParent.Root
    End If
End Property

Friend Property Get VisibleNodesCount() As Long
    VisibleNodesCount = mlngVisibleNodesCount
End Property

Friend Property Get Parent() As Node
    Set Parent = mobjParent
End Property
Friend Property Set Parent(ByVal Value As Node)
    Set mobjParent = Value
End Property

Public Property Get Level() As Long
    Level = mlngLevel
End Property
Public Property Let Level(ByVal Value As Long)
    mlngLevel = Value
End Property

Public Property Get Key() As String
    Key = mstrKey
End Property
Public Property Let Key(ByVal Value As String)
    mstrKey = Value
End Property

Public Property Get Text() As String
    Text = mstrText
End Property
Public Property Let Text(ByVal Value As String)
    mstrText = Value
End Property

Public Property Get HasChildren() As Boolean
    If mobjNodes Is Nothing Then
        HasChildren = False
    Else
        HasChildren = True
    End If
End Property

Public Property Get Nodes() As Nodes
    If mobjNodes Is Nothing Then
        Set mobjNodes = New Nodes
        Set mobjNodes.Owner = Me
    End If
    Set Nodes = mobjNodes
End Property

Public Property Get Expanded() As Boolean
    Expanded = mblnExpanded
End Property

Public Property Get Visible() As Boolean
    If mlngLevel = 0 Then
        Visible = False
        Exit Property
    End If
    Dim i As Long
    Dim tmpNode As Node
    Set tmpNode = mobjParent
    For i = mlngLevel To 1 Step -1
        If Not tmpNode.Expanded Then
            Visible = False
            Exit Sub
        End If
        Set tmpNode = tmpNode.Parent
    Next
    Visible = True
End Property

Public Property Get ChildrenCount() As Long
    ChildrenCount = mlngChildrenCount
End Property

Private Sub CopyNode(ByVal SourceNode As Node, ByVal TargetNode As Node)
    If Not SourceNode.Expanded Then
        TargetNode.Collapse
    End If
    If SourceNode.ChildrenCount > 0 Then
        Dim i As Long
        Dim objNode As Node
        For i = 1 To SourceNode.Nodes.Count
            With SourceNode.Nodes.Item(i)
                Set objNode = TargetNode.Nodes.Add(.Text)
                If Not .Expanded Then
                    objNode.Collapse
                End If
            End With
            CopyNode SourceNode.Nodes.Item(i), objNode
        Next
    End If
End Sub

Friend Sub UpdateChindrenCount(ByVal diff As Long)
    mlngChildrenCount = mlngChildrenCount + diff
    If Not mobjParent Is Nothing Then
        mobjParent.UpdateChindrenCount (diff)
    End If
End Sub

Friend Sub UpdateVisibleNodesCount(ByVal diff As Long)
    mlngVisibleNodesCount = mlngVisibleNodesCount + diff
    If Not mobjParent Is Nothing Then
        If mobjParent.Expanded Then
            mobjParent.UpdateVisibleNodesCount (diff)
        End If
    End If
End Sub

Friend Function FindNextVisibleNode() As Node
    Set FindNextVisibleNode = FindNextVisibleNodeEx(False)
End Function

Friend Function FindNextVisibleNodeEx(ByVal FindNext As Boolean) As Node
    Dim i As Long
    Dim CollapsedNode As Node

    If mlngLevel > 0 Then
        Dim tmpNode As Node
        Set tmpNode = mobjParent
        For i = mlngLevel To 1 Step -1
            If Not tmpNode.Expanded Then
                Set CollapsedNode = tmpNode
            End If
            Set tmpNode = tmpNode.Parent
        Next
    End If

    If Not CollapsedNode Is Nothing Then
        Set FindNextVisibleNodeEx = CollapsedNode.FindNextVisibleNode
        Exit Function
    End If

    If (Not FindNext) And mblnExpanded And mlngChildrenCount > 0 Then
        Set FindNextVisibleNodeEx = mobjNodes.FirstNode
        Exit Function
    End If

    If mobjNextNode Is Nothing Then
        If mobjParent Is Nothing Or mlngLevel = 1 Then
            Set FindNextVisibleNodeEx = Nothing
        Else
            Set FindNextVisibleNodeEx = mobjParent.FindNextVisibleNode(True)
        End If
    Else
        Set FindNextVisibleNodeEx = mobjNextNode
    End If
End Function

Friend Function FindLastVisibleChildNode() As Node
    If mlngVisibleNodesCount = 1 Then
        Set FindLastVisibleChildNode = Me
    Else
        Set FindLastVisibleChildNode = mobjNodes.LastNode.FindLastVisibleChildNode
    End If
End Function

Friend Function FindPrevVisibleNode() As Node
    Dim i As Long
    Dim CollapsedNode As Node

    If mlngLevel > 0 Then
        Dim tmpNode As Node
        Set tmpNode = mobjParent
        For i = mlngLevel To 1 Step -1
            If Not tmpNode.Expanded Then
                Set CollapsedNode = tmpNode
            End If
            Set tmpNode = tmpNode.Parent
        Next
    End If

    If Not CollapsedNode Is Nothing Then
        Set FindPrevVisibleNode = CollapsedNode
        Exit Function
    End If

    If Not mobjPrevNode Is Nothing Then
        Set FindPrevVisibleNode = mobjPrevNode.FindLastVisibleChildNode
    Else
        If mobjParent Is Nothing Then
            Set FindPrevVisibleNode = Nothing
        Else
            Set FindPrevVisibleNode = mobjParent
        End If
    End If
End Function

Public Sub Collapse()
    If (mlngLevel = 0) Or (Not mblnExpanded) Then
        Exit Sub
    End If
    mblnExpanded = False
    UpdateVisibleNodesCount 1 - mlngVisibleNodesCount
End Sub

Public Sub Expand()
    If mblnExpanded Then
        Exit Sub
    End If
    Dim i As Long
    mblnExpanded = True
    If Not mobjNodes Is Nothing Then
        For i = 1 To mobjNodes.Count
            UpdateVisibleNodesCount mobjNodes.Item(i).VisibleNodesCount
        Next
    End If
End Sub

Public Sub ExpandAll()
    Expand

    Dim i As Long
    If Not mobjNodes Is Nothing Then
        For i = 1 To mobjNodes.Count
            mobjNodes.Item(i).ExpandAll
        Next
    End If
End Sub

Public Sub LevelUp()
    If mobjParent.Level = 1 Then
        Exit Sub
    End If
    Dim objNode As Node
    Set objNode = mobjParent.Parent.Nodes.InsertAfter(mobjParent.Key, mstrText)
    CopyNode Me, objNode
    mobjParent.Nodes.Remove mstrKey
End Sub

Public Sub LevelDown()
    If mobjPrevNode Is Nothing Then
        Exit Sub
    End If
    Dim objNode As Node
    Set objNode = mobjPrevNode.Nodes.Add(mstrText)
    CopyNode Me, objNode
    mobjParent.Nodes.Remove mstrKey
End Sub

Public Sub MoveUp()
    If mobjPrevNode Is Nothing Then
        Exit Sub
    End If
    Dim objNode As Node
    Set objNode = mobjParent.Nodes.InsertBefore(mobjPrevNode.Key, mstrText)
    CopyNode Me, objNode
    mobjParent.Nodes.Remove mstrKey
End Sub

Public Sub MoveDown()
    If mobjNextNode Is Nothing Then
        Exit Sub
    End If
    Dim objNode As Node
    Set objNode = mobjParent.Nodes.InsertAfter(mobjNextNode.Key, mstrText)
    CopyNode Me, objNode
    mobjParent.Nodes.Remove mstrKey
End Sub

Public Function FindNode(ByVal Index As Long) As Node
    If Index = 1 Then
        Set FindNode = Me
        Exit Function
    End If
    Dim i As Long
    i = Index - 1
    If Me.ChildrenCount >= i Then
        Set FindNode = Me.Nodes.FirstNode.FindNode(i)
    Else
        i = i - Me.ChildrenCount
        If Me.NextNode Is Nothing Then
            Set FindNode = Nothing
        Else
            Set FindNode = Me.NextNode.FindNode(i)
        End If
    End If
End Function

Private Sub Class_Initialize()
    mstrKey = ""
    mstrText = ""
    mlngLevel = 0
    mblnExpanded = True
    mlngChildrenCount = 0
    mlngVisibleNodesCount = 1
End Sub

