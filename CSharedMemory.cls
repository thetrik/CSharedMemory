VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSharedMemory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // CSharedMemory.cls
' // Class provides dynamic memory allocation in shared memory
' // By The Trick 2020 - 2022
' //

Option Explicit

' // Heap structure

' // 0x00 : Lock
' // 0x04 : MaxSize not including header
' // 0x08 : Head Offset relative to begin of section
' // 0x0C : Blocks

' // Block structure (describes a free memory block)

' // 0x00 : Block size not including header
' // 0x04 : Next offset relative to begin of section
' // 0x08 : Prev offset relative to begin of section
' // 0x0C : Data

Private Const OFFSETOF_LOCK           As Long = 0
Private Const OFFSETOF_MAXSIZE        As Long = 4
Private Const OFFSETOF_HEAD           As Long = 8
Private Const OFFSETOF_FIRSTBLOCK     As Long = 12
Private Const HEAP_HDR_SIZE           As Long = 12
Private Const OFFSETOF_SIZE           As Long = 0
Private Const OFFSETOF_PREV           As Long = 4
Private Const OFFSETOF_NEXT           As Long = 8
Private Const OFFSETOF_DATA           As Long = 12
Private Const BLOCK_HDR_SIZE          As Long = 12

Private Const INVALID_HANDLE_VALUE    As Long = -1
Private Const FILE_MAP_WRITE          As Long = 2
Private Const FILE_MAP_READ           As Long = 4
Private Const PAGE_READWRITE          As Long = 4&
Private Const SectionBasicInformation As Long = 0

Private Type tBlock
    lSize               As Long
    lPrev               As Long
    lNext               As Long
End Type

Private Type LARGE_INTEGER
    lowPart             As Long
    highPart            As Long
End Type

Private Type SECTION_BASIC_INFORMATION
    lUnknown            As Long
    lSectionAttributes  As Long
    szSectionSize       As LARGE_INTEGER
End Type

Private Declare Function NtQuerySection Lib "ntdll" ( _
                         ByVal hSection As Long, _
                         ByVal lInformationClass As Long, _
                         ByRef pInformationBuffer As Any, _
                         ByVal lInformationBufferSize As Long, _
                         ByRef lReturnLength As Long) As Long
Private Declare Function InterlockedExchange Lib "kernel32" ( _
                         ByRef Target As Any, _
                         ByVal Value As Long) As Long
Private Declare Function InterlockedCompareExchange Lib "kernel32" ( _
                         ByRef Destination As Any, _
                         ByVal Exchange As Long, _
                         ByVal Comperand As Long) As Long
Private Declare Function Sleep Lib "kernel32" ( _
                         ByVal dwMilliseconds As Long) As Long
Private Declare Function GetMem4 Lib "msvbvm60" ( _
                         ByRef pSrc As Any, _
                         ByRef pDst As Any) As Long
Private Declare Function GetMem8 Lib "msvbvm60" ( _
                         ByRef pSrc As Any, _
                         ByRef pDst As Any) As Long
Private Declare Function PutMem8 Lib "msvbvm60" ( _
                         ByRef pDst As Any, _
                         ByVal cVal As Currency) As Long
Private Declare Function PutMem4 Lib "msvbvm60" ( _
                         ByRef pDst As Any, _
                         ByVal lVal As Long) As Long
Private Declare Function CreateFileMapping Lib "kernel32" _
                         Alias "CreateFileMappingW" ( _
                         ByVal hFile As Long, _
                         ByRef lpFileMappingAttributes As Any, _
                         ByVal flProtect As Long, _
                         ByVal dwMaximumSizeHigh As Long, _
                         ByVal dwMaximumSizeLow As Long, _
                         ByVal lpName As Long) As Long
Private Declare Function MapViewOfFile Lib "kernel32" ( _
                         ByVal hFileMappingObject As Long, _
                         ByVal dwDesiredAccess As Long, _
                         ByVal dwFileOffsetHigh As Long, _
                         ByVal dwFileOffsetLow As Long, _
                         ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" ( _
                         ByVal hObject As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32" ( _
                         ByVal lpBaseAddress As Long) As Long
Private Declare Sub memcpy Lib "kernel32" _
                    Alias "RtlMoveMemory" ( _
                    ByRef pDestination As Any, _
                    ByRef pSource As Any, _
                    ByVal lLength As Long)
                    
Private m_hMap      As OLE_HANDLE
Private m_pRawData  As Long

' // Section handle
Public Property Get Handle() As OLE_HANDLE
    Handle = m_hMap
End Property

' // Address of heap
Public Property Get Address() As Long
    Address = m_pRawData
End Property

' // Allocate memory block
Public Function Alloc( _
                ByVal lSize As Long) As Long
    Dim lError  As Long
    
    If m_pRawData = 0 Or lSize <= 0 Then
        Err.Raise 5
    End If
    
    On Error GoTo unlock_access
    
    LockAccess
    
    Alloc = AllocInternal(lSize) + m_pRawData
    
unlock_access:
    
    lError = Err.Number
    
    UnlockAccess
    
    If lError Then
        Err.Raise lError
    End If
    
End Function

' // Free memory block
Public Sub Free( _
           ByVal pData As Long)
    Dim lError      As Long
    Dim lMaxSize    As Long
    
    If m_pRawData = 0 Then
        Err.Raise 5
    Else
        
        GetMem4 ByVal m_pRawData + OFFSETOF_MAXSIZE, lMaxSize
    
        If pData < m_pRawData + HEAP_HDR_SIZE + BLOCK_HDR_SIZE Or pData >= m_pRawData + lMaxSize + HEAP_HDR_SIZE Then
            Err.Raise 9
        End If
        
    End If
    
    On Error GoTo unlock_access
    
    LockAccess
    
    FreeInternal pData - m_pRawData
    
unlock_access:
    
    lError = Err.Number
    
    UnlockAccess
    
    If lError Then
        Err.Raise lError
    End If
        
End Sub

' // Initialize new memory region
Public Sub Initialize( _
           ByVal lMaxSize As Long, _
           Optional ByRef sName As String)
    Dim hMap    As OLE_HANDLE
    Dim pData   As Long
    
    If lMaxSize <= BLOCK_HDR_SIZE Then
        Err.Raise 5
    End If
    
    hMap = CreateFileMapping(INVALID_HANDLE_VALUE, ByVal 0&, PAGE_READWRITE, 0, lMaxSize + HEAP_HDR_SIZE, StrPtr(sName))
    
    If hMap = 0 Then
        Err.Raise 7
    End If
    
    pData = MapViewOfFile(hMap, FILE_MAP_READ Or FILE_MAP_WRITE, 0, 0, lMaxSize + HEAP_HDR_SIZE)
    
    If pData = 0 Then
        CloseHandle hMap
        Err.Raise 7
    Else
        If m_pRawData Then
            UnmapViewOfFile m_pRawData
            CloseHandle m_hMap
        End If
    End If
    
    m_hMap = hMap
    m_pRawData = pData
    
    ' // Update max size
    PutMem4 ByVal m_pRawData + OFFSETOF_MAXSIZE, lMaxSize
    
    ' // Create free block
    PutMem4 ByVal m_pRawData + OFFSETOF_FIRSTBLOCK + OFFSETOF_SIZE, lMaxSize - BLOCK_HDR_SIZE
    BlockAdd OFFSETOF_FIRSTBLOCK
    
End Sub

' // Open existing region
Public Sub OpenFromHandle( _
           ByVal hMap As OLE_HANDLE, _
           Optional ByVal bInitialize As Boolean)
    Dim pData       As Long
    Dim lStatus     As Long
    Dim tSecInfo    As SECTION_BASIC_INFORMATION
    
    lStatus = NtQuerySection(hMap, SectionBasicInformation, tSecInfo, Len(tSecInfo), 0)
    If lStatus <> 0 Then
        Err.Raise 5
    End If
    
    pData = MapViewOfFile(hMap, FILE_MAP_READ Or FILE_MAP_WRITE, 0, 0, tSecInfo.szSectionSize.lowPart)
    If pData = 0 Then
        Err.Raise 7
    End If
    
    If m_hMap Then
        UnmapViewOfFile m_pRawData
        CloseHandle m_hMap
    End If
    
    m_hMap = hMap
    m_pRawData = pData
    
    If bInitialize Then
    
        LockAccess
        
        ' // Update max size
        PutMem4 ByVal m_pRawData + OFFSETOF_MAXSIZE, tSecInfo.szSectionSize.lowPart - HEAP_HDR_SIZE
        
        ' // Create free block
        PutMem4 ByVal m_pRawData + OFFSETOF_FIRSTBLOCK + OFFSETOF_SIZE, tSecInfo.szSectionSize.lowPart - HEAP_HDR_SIZE - BLOCK_HDR_SIZE
        BlockAdd OFFSETOF_FIRSTBLOCK
        
        UnlockAccess
        
    End If
    
End Sub

' // Unlock access
Private Sub UnlockAccess()
    InterlockedExchange ByVal m_pRawData + OFFSETOF_LOCK, 0
End Sub

' // Lock access
Private Sub LockAccess()
    Dim lSpinCount  As Long
    Dim lIndex      As Long
    
    lSpinCount = 10000
    
    Do While InterlockedCompareExchange(ByVal m_pRawData + OFFSETOF_LOCK, 1, 0)
        
        For lIndex = 0 To lSpinCount - 1
            If InterlockedCompareExchange(ByVal m_pRawData + OFFSETOF_LOCK, 1, 0) = 0 Then
                Exit Do
            End If
        Next
        
        Sleep 10
        
    Loop

End Sub

' // Free memory block
Private Sub FreeInternal( _
            ByVal lOffset As Long)
    BlockAdd lOffset - BLOCK_HDR_SIZE
    BlockScanFreeMerge
End Sub

' // Alloc memory. Returns offset
Private Function AllocInternal( _
                 ByVal lSize As Long) As Long
    Dim lBlock      As Long
    Dim lBlockSize  As Long
    Dim lNewBlock   As Long
    Dim lHeadOffset As Long
    
    GetMem4 ByVal m_pRawData + OFFSETOF_HEAD, lHeadOffset
    
    If lHeadOffset = 0 Then
        ' // No memory
        Err.Raise 7
    End If
    
    ' // Scan for free blocks
    lBlock = lHeadOffset
    
    Do While lBlock
        
        GetMem4 ByVal m_pRawData + lBlock + OFFSETOF_SIZE, lBlockSize
        
        If lBlockSize >= lSize And lBlockSize <= lSize + BLOCK_HDR_SIZE Then
            
            ' // Allocate full block
            BlockRemove lBlock
            AllocInternal = lBlock + BLOCK_HDR_SIZE
            Exit Function
            
        ElseIf lBlockSize > lSize + BLOCK_HDR_SIZE Then
            
            ' // Divide block
            BlockRemove lBlock
            lNewBlock = BlockSplit(lBlock, lSize)
            BlockAdd lNewBlock
            AllocInternal = lBlock + BLOCK_HDR_SIZE
            Exit Function
                
        Else
            GetMem4 ByVal m_pRawData + lBlock + OFFSETOF_NEXT, lBlock
        End If
        
    Loop
                    
    Err.Raise 7
    
End Function

' // Merge adjacent blocks
Private Sub BlockScanFreeMerge()
    Dim lHeadOffset As Long
    Dim lCurBlock   As Long
    Dim lNextBlock  As Long
    Dim lCurSize    As Long
    Dim lNextSize   As Long
    
    ' // Get head offset
    GetMem4 ByVal m_pRawData + OFFSETOF_HEAD, lHeadOffset
    
    If lHeadOffset = 0 Then
        ' // No free blocks
        Exit Sub
    End If
    
    lCurBlock = lHeadOffset
    
    Do
        
        ' // Get next block
        GetMem4 ByVal m_pRawData + lCurBlock + OFFSETOF_NEXT, lNextBlock
 
        If lNextBlock = 0 Then
            Exit Do
        End If
        
        ' // Get current block size
        GetMem4 ByVal m_pRawData + lCurBlock + OFFSETOF_SIZE, lCurSize
        
        If lCurBlock + lCurSize + BLOCK_HDR_SIZE = lNextBlock Then
            
            ' // Next block is adjacent. Merge both block to a single block
            
            ' // Get next block size
            GetMem4 ByVal m_pRawData + lNextBlock + OFFSETOF_SIZE, lNextSize
            
            ' // Update current block size
            PutMem4 ByVal m_pRawData + lCurBlock + OFFSETOF_SIZE, lNextSize + lCurSize + BLOCK_HDR_SIZE

            ' // Get next block of adjacent block
            GetMem4 ByVal m_pRawData + lNextBlock + OFFSETOF_NEXT, lNextBlock

            ' // Update next block of current block
            PutMem4 ByVal m_pRawData + lCurBlock + OFFSETOF_NEXT, lNextBlock

            If lNextBlock Then
                ' // Update prev block of next block of adjacent block
                PutMem4 ByVal m_pRawData + lNextBlock + OFFSETOF_PREV, lCurBlock
            Else
                Exit Do
            End If
            
        Else
            lCurBlock = lNextBlock
        End If
        
    Loop
    
End Sub

' // Split free memory block. Returns new block offset
Private Function BlockSplit( _
                 ByVal lBlockOffset As Long, _
                 ByVal lSize As Long) As Long
    Dim lNewBlock   As Long
    Dim lBlockSize  As Long
    
    ' // Get block size
    GetMem4 ByVal m_pRawData + lBlockOffset + OFFSETOF_SIZE, lBlockSize
    
    ' // Get new block offset
    lNewBlock = lBlockOffset + OFFSETOF_DATA + lSize
    
    ' // Update new block size
    PutMem4 ByVal m_pRawData + lNewBlock + OFFSETOF_SIZE, lBlockSize - (lSize + BLOCK_HDR_SIZE)
    
    ' // Update old block size
    PutMem4 ByVal m_pRawData + lBlockOffset + OFFSETOF_SIZE, lSize

    BlockSplit = lNewBlock
    
End Function

' // Add free memory block
Private Sub BlockAdd( _
            ByVal lBlockOffset As Long)
    Dim lCurBlock   As Long
    Dim lNextBlock  As Long
    Dim lHeadOffset As Long
    
    ' // Get head offset
    GetMem4 ByVal m_pRawData + OFFSETOF_HEAD, lHeadOffset
    
    ' // Zero next/prev
    PutMem8 ByVal m_pRawData + lBlockOffset + OFFSETOF_PREV, 0@
    
    If lHeadOffset = 0 Or lHeadOffset > lBlockOffset Then
        
        If lHeadOffset Then
            ' // Update prev block
            PutMem4 ByVal m_pRawData + lHeadOffset + OFFSETOF_PREV, lBlockOffset
        End If
        
        ' // Update next block
        PutMem4 ByVal m_pRawData + lBlockOffset + OFFSETOF_NEXT, lHeadOffset
        
        ' // Update head
        PutMem4 ByVal m_pRawData + OFFSETOF_HEAD, lBlockOffset

    Else
        
        ' // Search for block
        
        lCurBlock = lHeadOffset
        
        Do
        
            GetMem4 ByVal m_pRawData + lCurBlock + OFFSETOF_NEXT, lNextBlock
            
            If lNextBlock = 0 Or lNextBlock >= lBlockOffset Then
                Exit Do
            End If
            
            lCurBlock = lNextBlock
            
        Loop
        
        ' // Update blocks
        PutMem4 ByVal m_pRawData + lBlockOffset + OFFSETOF_NEXT, lNextBlock
        PutMem4 ByVal m_pRawData + lCurBlock + OFFSETOF_NEXT, lBlockOffset
        PutMem4 ByVal m_pRawData + lBlockOffset + OFFSETOF_PREV, lCurBlock

        If lNextBlock Then
            PutMem4 ByVal m_pRawData + lNextBlock + OFFSETOF_PREV, lBlockOffset
        End If
        
    End If

End Sub

' // Remove free memory block
Private Sub BlockRemove( _
            ByVal lBlockOffset As Long)
    Dim tBlock  As tBlock
    
    GetMem8 ByVal m_pRawData + lBlockOffset + OFFSETOF_PREV, tBlock.lPrev
     
    If tBlock.lPrev Then
        ' // Update prev block
        PutMem4 ByVal m_pRawData + tBlock.lPrev + OFFSETOF_NEXT, tBlock.lNext
    Else
        ' // Update head
        PutMem4 ByVal m_pRawData + OFFSETOF_HEAD, tBlock.lNext
    End If
    
    If tBlock.lNext Then
        ' // Update next block
        PutMem4 ByVal m_pRawData + tBlock.lNext + OFFSETOF_PREV, tBlock.lPrev
    End If

End Sub

Private Sub Class_Terminate()

    If m_hMap Then
        
        UnmapViewOfFile m_pRawData
        CloseHandle m_hMap
        
    End If
    
End Sub
