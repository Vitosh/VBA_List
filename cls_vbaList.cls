VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_vbaList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mList()             As Variant
Private mError              As Error
Private mDisposed           As Boolean

'==============================
'Constructor
'==============================
Public Sub Initialize()
    Disposed = False
End Sub

Public Function CreateInstance() As cls_vbaList

    Dim oNew As New cls_vbaList
    oNew.Initialize
    Set CreateInstance = oNew
    
End Function

'==============================
'Properties
'==============================
Public Property Get Items(ByRef index As Long) As Variant
    Items = GetItemAtIndex(index)
End Property
Public Property Get Count() As Long
    Count = GetListCount()
End Property

Public Property Get GotError() As Boolean
    If ListError Is Nothing Then GotError = False Else GotError = True
End Property

Public Property Get ListItems() As Variant()

    ClearError
    On Error GoTo Err
    ListItems = mList
    
    Exit Property
    
Err:
    ListError = Err
    
End Property

Public Property Get ListError() As Error
    Set ListError = mError
End Property

Private Property Let ListError(ByRef vError As Error)
    Set mError = vError
End Property

Public Property Get Disposed() As Boolean
    Disposed = mDisposed
End Property

Private Property Let Disposed(ByRef vValue As Boolean)
    mDisposed = vValue
End Property

Public Property Get ToArray()
    ToArray = mList
End Property

'==============================
'Public Methods
'==============================

Public Sub Remove(ByRef vItem As Variant)
    DeleteElement (vItem)
End Sub

Public Sub RemoveAtIndex(ByRef index As Long)
    DeleteElementAt (index)
End Sub

Public Sub Sort()
    BubbleSort (mList)
End Sub

Public Sub Clear()
    Erase mList
End Sub

Public Function Find(ByRef vItem As Variant) As Long
    Find = FindItem(vItem)
End Function

Public Sub Dispose()

    ResetError
    Clear
    Disposed = True
    
End Sub

Public Sub ResetError()
    ClearError
End Sub

Public Function LastIndexOf(ByRef vItem As Variant)
    LastIndexOf = GetLastIndexOf(vItem)
End Function

Public Function IndexOf(ByRef vItem As Variant)
    IndexOf = FindItem(vItem)
End Function

Public Sub Reverse()
    ReverseList
End Sub

Public Function Exists(vItem As Variant)
    Exists = ItemExists(vItem)
End Function

Public Sub Add(ByRef vItem As Variant, Optional index As Long)

    If index > 0 Then
        AddItemAt index, vItem
    Else
        AddItem vItem
    End If
    
End Sub

Public Function Contains(ByRef vItem As Variant)
    Contains = Exists(vItem)
End Function

Public Function Copy() As cls_vbaList
    Set Copy = GetCopy
End Function

Public Sub RemoveAll()
    Clear
End Sub

'==============================
'Public Methods
'==============================

Private Sub ClearError()

    Set mError = Nothing

End Sub

Private Function GetListCount() As Long

    ClearError
    On Error GoTo Err
    
    GetListCount = UBound(mList) - LBound(mList) + 1
    Exit Function

Err:
    GetListCount = 0

End Function

Private Function GetItemAtIndex(ByRef index As Long) As Variant

    ClearError
    On Error GoTo Err
    
    GetItemAtIndex = mList(index)
    
    Exit Function
    
Err:
    ListError = Err
    GetItemAtIndex = Nothing

End Function

Private Sub AddItemAt(index As Long, vItem As Variant)

    Dim ar()        As Variant
    Dim i           As Long
    Dim a           As Long

    ClearError
    On Error GoTo Err

    i = Count
    ReDim ar(i)

    For a = 0 To index - 1
        ar(a) = mList(a)
    Next

    ar(index) = vItem

    For a = index + 1 To i
        ar(a) = mList(a - 1)
    Next

    mList = ar
    Exit Sub
Err:
    ListError = Err
End Sub

Private Sub BubbleSort(ByVal vArray As Variant)

    ClearError
    On Error GoTo Err

    Dim i               As Long
    Dim iMin            As Long
    Dim iMax            As Long
    Dim vSwap           As Variant
    Dim swapped         As Boolean

    iMin = LBound(vArray)
    iMax = UBound(vArray) - 1

    Do
        swapped = False
        For i = iMin To iMax
            If vArray(i) > vArray(i + 1) Then
                vSwap = vArray(i)
                vArray(i) = vArray(i + 1)
                vArray(i + 1) = vSwap
                swapped = True
            End If
        Next
        iMax = iMax - 1
    Loop Until Not swapped
    mList = vArray
    Erase vArray
    Exit Sub
Err:
    ListError = Err

End Sub

Private Sub DeleteElementAt(index As Long)
    
    Dim i As Long
    
    ClearError
    On Error GoTo Err

    For i = index + 1 To Count - 1
        mList(i - 1) = mList(i)
    Next
    ReDim Preserve mList(Count - 2)
    Exit Sub
Err:
    ListError = Err
End Sub

Private Sub DeleteElement(ByRef vItem As Variant)

    ClearError
    On Error GoTo Err

    DeleteElementAt (FindItem(vItem))
    Exit Sub
    
Err:
    ListError = Err
    
End Sub

Private Sub AddItem(vItem As Variant)

    Dim i As Long
    
    ClearError
    On Error GoTo Err

    i = Count
    ReDim Preserve mList(i)
    mList(i) = vItem
    
    Exit Sub
    
Err:
    ListError = Err
    
End Sub

Private Function FindItem(vItem As Variant) As Long

    Dim i           As Long
        
    ClearError
    On Error GoTo Err

    FindItem = -1

    For i = 0 To Count - 1
        If mList(i) = vItem Then
            FindItem = i
            Exit For
        End If
    Next i
    Exit Function
    
Err:
    ListError = Err
    FindItem = -1
    
End Function

Private Function GetLastIndexOf(vItem As Variant) As Long

    Dim i As Long

    ClearError
    On Error GoTo Err

    GetLastIndexOf = -1

    For i = Count - 1 To 0 Step -1
        If mList(i) = vItem Then
            GetLastIndexOf = i
            Exit Function
        End If
    Next i
    Exit Function
    
Err:
    ListError = Err
    GetLastIndexOf = -1
    
End Function

Private Sub ReverseList()

    Dim ar()            As Variant
    Dim a               As Long
    Dim i               As Long
    Dim j               As Long

    ClearError
    On Error GoTo Err

    If Count = 0 Then Exit Sub
    i = Count - 1
    j = i
    ReDim ar(i)

    For a = 0 To i
        ar(a) = mList(j)
        j = j - 1
    Next a

    mList = ar
    Erase ar
    Exit Sub
    
Err:
    ListError = Err
    
End Sub

Private Function ItemExists(vItem As Variant) As Boolean

    If FindItem(vItem) > -1 Then
        ItemExists = True
    Else
        ItemExists = False
    End If
    
End Function

Private Function GetCopy() As cls_vbaList

    Dim list            As New cls_vbaList
    Dim i               As Long

    Set list = list.CreateInstance

    For i = 0 To Count - 1
        list.Add mList(i)
    Next i
    
    Set GetCopy = list
    i = GetCopy.Count

End Function

