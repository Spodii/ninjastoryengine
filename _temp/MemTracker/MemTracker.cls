VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Information on each object held
Private Type MemoryObject
    LastUsed As Long    'Tick the memory was last used at
    ObjType As Byte     'Which type of object it is
    ObjIndex As Long    'Index of the object type
End Type
Private MemObj() As MemoryObject

'Number of memory objects total
Private NumMemObjs As Long

'UBound of the MemObj() array
Private MemObjUBound As Long

'Sorted (by LastUsed, ascending) pointers to the MemObj() array
Private SortArray() As Long

'Lowest index that is unsorted - used to determine if we actually need to sort the
'data before returning the sorted array. Use -1 if the array is sorted. 0 to force sorting.
Private LowestUnsortIndex As Long

'How many unused elements the array contains
Private NumUnused As Long

'Value for an unused index
Private Const MEMTYPE_UNUSED As Long = 2147483647

Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Sub IncreaseArray()

    'Increase the array size by 1000 elements
    MemObjUBound = MemObjUBound + 1000
    ReDim Preserve MemObj(0 To MemObjUBound)
    ReDim Preserve SortArray(0 To MemObjUBound)

End Sub

Public Sub GetOldest(ByVal MaxReturnSize As Long, ByRef retType() As Byte, ByRef retIndex() As Long)
Dim i As Long

    'Check for a valid return size
    If MaxReturnSize > NumMemObjs - NumUnused Then MaxReturnSize = NumMemObjs - NumUnused
    If MaxReturnSize < 0 Then Exit Sub
    
    'Check if the array needs to be sorted
    If LowestUnsortIndex > -1 Then
        If LowestUnsortIndex <= MaxReturnSize Then
            Debug.Print "Sorting " & NumMemObjs - LowestUnsortIndex
            Sort LowestUnsortIndex, NumMemObjs
            LowestUnsortIndex = -1
        End If
    End If
    
    'Return the elements
    ReDim retType(0 To MaxReturnSize)
    ReDim retIndex(0 To MaxReturnSize)
    For i = 0 To MaxReturnSize
        retType(i) = MemObj(SortArray(i)).ObjType
        retIndex(i) = MemObj(SortArray(i)).ObjIndex
    Next i

End Sub

Private Sub Sort(ByVal inLow As Long, ByVal inHi As Long)
Dim tmpSwap As Long
Dim tmpLow As Long
Dim tmpHi As Long
Dim pivot As Long

    'Sort the array using Quicksort
    'The actual array MemObj() is never modified, but rather a list of pointers
    'to the MemObj() indicies are stored, so the sorted list is MemObj(SortArray(i))
    tmpLow = inLow
    tmpHi = inHi
    pivot = MemObj(SortArray((inLow + inHi) \ 2)).LastUsed
    
    Do While (tmpLow <= tmpHi)
    
        Do While (MemObj(SortArray(tmpLow)).LastUsed < pivot And tmpLow < inHi)
            tmpLow = tmpLow + 1
        Loop
        
        Do While (pivot < MemObj(SortArray(tmpHi)).LastUsed And tmpHi > inLow)
            tmpHi = tmpHi - 1
        Loop
        
        If (tmpLow <= tmpHi) Then
            tmpSwap = SortArray(tmpLow)
            SortArray(tmpLow) = SortArray(tmpHi)
            SortArray(tmpHi) = tmpSwap
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
        End If
        
    Loop
    
    If (inLow < tmpHi) Then Sort inLow, tmpHi
    If (tmpLow < inHi) Then Sort tmpLow, inHi

End Sub

Public Function Add(ByVal ObjType As Byte, ByVal ObjIndex As Long) As Long
Dim MemIndex As Long
Dim i As Long

    If NumUnused > 0 Then
        
        'Recycle an index
        For i = 0 To NumMemObjs
            If MemObj(i).LastUsed = MEMTYPE_UNUSED Then
                MemIndex = i
                NumUnused = NumUnused - 1
                Exit For
            End If
        Next i
     
    Else
    
        'Increase the number of objects and if needed, the array
        NumMemObjs = NumMemObjs + 1
        If NumMemObjs > MemObjUBound Then IncreaseArray
        MemIndex = NumMemObjs
        
    End If
        
    'Store the values
    MemObj(MemIndex).ObjType = ObjType
    MemObj(MemIndex).ObjIndex = ObjIndex
    MemObj(MemIndex).LastUsed = timeGetTime
    
    'Set the SortArray value
    SortArray(MemIndex) = NumMemObjs
    
    'Set the lowest unsorted index
    If MemIndex < LowestUnsortIndex Or LowestUnsortIndex = -1 Then LowestUnsortIndex = MemIndex
    
    'Return the index used
    Add = MemIndex

End Function

Public Sub RemoveEX(ByVal ObjType As Byte, ByVal ObjIndex As Long)
Dim i As Long
    
    'Remove an object's life by their ObjType and ObjIndex (very slow)
    For i = 0 To NumMemObjs
        If MemObj(i).ObjIndex = ObjIndex Then
            If MemObj(i).ObjType = ObjType Then
                Remove i
                Exit For
            End If
        End If
    Next i
    
End Sub

Public Sub Remove(ByVal MemIndex As Long)

    'Set the last used time to the highest value of a long
    MemObj(MemIndex).LastUsed = MEMTYPE_UNUSED
    
    'Increase the unused count
    NumUnused = NumUnused + 1
    
    'Set the lowest unsorted index
    If MemIndex < LowestUnsortIndex Or LowestUnsortIndex = -1 Then LowestUnsortIndex = MemIndex
    
End Sub

Public Sub UpdateEX(ByVal ObjType As Byte, ByVal ObjIndex As Long)
Dim i As Long
    
    'Update an object's life by their ObjType and ObjIndex (very slow)
    For i = 0 To NumMemObjs
        If MemObj(i).ObjIndex = ObjIndex Then
            If MemObj(i).ObjType = ObjType Then
                Update i
                Exit For
            End If
        End If
    Next i

End Sub

Public Sub Update(ByVal MemIndex As Long)

    'Update an object's life by their array index
    MemObj(MemIndex).LastUsed = timeGetTime
    
    'Set the lowest unsorted index
    If MemIndex < LowestUnsortIndex Or LowestUnsortIndex = -1 Then LowestUnsortIndex = MemIndex

End Sub

Private Sub Class_Initialize()
    
    NumMemObjs = -1
    MemObjUBound = -1
    IncreaseArray
    timeBeginPeriod 1

End Sub
