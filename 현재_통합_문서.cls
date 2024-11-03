VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "현재_통합_문서"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    Call ListAuthorsAndWorks
End Sub

Sub ListAuthorsAndWorks()
    Application.ScreenUpdating = False  ' 화면 업데이트 일시 중지

    Dim FileSystem As Object
    Dim MainFolder As Object
    Dim AuthorFolder As Object
    Dim WorkFolder As Object
    Dim MainFolderPath As String
    Dim ExistingAuthors As Object
    Dim ExistingWorks As Object
    Dim CurrentAuthors As Object
    Dim CurrentWorks As Object
    Dim AuthorName As String
    Dim WorkName As String
    Dim CurrentAuthorName As String
    Dim CurrentAuthorRow As Long
    Dim RowNum As Variant  ' Variant로 변경
    Dim Key As Variant
    Dim OtherKey As Variant
    Dim ExistingKey As Variant
    Dim RowItem As Variant
    Dim DeletedRow As Variant
    Dim i As Long
    Dim LastRow As Long
    Dim StartRow As Long
    Dim EndRow As Long
    Dim UniqueRows As New Collection
    Dim SortedRows As New Collection
    Dim MaxValue As Long

    MainFolderPath = "각자 폴더 경로 작성하기"  ' 명작 폴더 경로 설정

    ' 경로 확인
    If Not FolderExists(MainFolderPath) Then
        MsgBox "경로를 찾을 수 없습니다: " & MainFolderPath
        Exit Sub
    End If

    Set FileSystem = CreateObject("Scripting.FileSystemObject")
    Set MainFolder = FileSystem.GetFolder(MainFolderPath)

    ' *** 헤더 추가 시작 ***
    With ThisWorkbook.ActiveSheet
        ' 헤더가 이미 있는지 확인
        If .Cells(1, 1).Value <> "작가명" Or .Cells(1, 2).Value <> "수정일" Or .Cells(1, 3).Value <> "작품명" Then
            ' 헤더 작성 (기존 데이터를 밀어내지 않음)
            .Cells(1, 1).Value = "작가명"
            .Cells(1, 2).Value = "수정일"
            .Cells(1, 3).Value = "작품명"
            
            ' 헤더 서식 지정 (선택 사항)
            With .Range("A1:C1")
                .Font.Bold = True
                .Interior.Color = RGB(200, 200, 200)
                .HorizontalAlignment = xlCenter
            End With
        End If
    End With
    ' *** 헤더 추가 끝 ***

    ' 기존 데이터 시트에서 기존 목록 확인
    Set ExistingAuthors = CreateObject("Scripting.Dictionary")
    Set ExistingWorks = CreateObject("Scripting.Dictionary")

    ' 그룹화 해제 (나중에 다시 적용)
    Cells.EntireRow.ClearOutline

    ' 기존 목록을 딕셔너리에 추가하여 추적
    ' 딕셔너리의 키: AuthorName 또는 AuthorName & "\" & WorkName
    ' 값: 행 번호
    CurrentAuthorName = ""
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        If Cells(i, 1).Value <> "" Then
            CurrentAuthorName = Cells(i, 1).Value
            CurrentAuthorRow = i
            ExistingAuthors(CurrentAuthorName) = i
        End If
        If Cells(i, 3).Value <> "" Then
            ExistingWorks(CurrentAuthorName & "\" & Cells(i, 3).Value) = i
        End If
    Next i

    ' 현재 폴더 구조에서 목록 가져오기
    Set CurrentAuthors = CreateObject("Scripting.Dictionary")
    Set CurrentWorks = CreateObject("Scripting.Dictionary")

    ' 현재 폴더 구조를 딕셔너리에 저장
    For Each AuthorFolder In MainFolder.SubFolders
        AuthorName = AuthorFolder.Name
        CurrentAuthors(AuthorName) = True
        For Each WorkFolder In AuthorFolder.SubFolders
            WorkName = WorkFolder.Name
            CurrentWorks(AuthorName & "\" & WorkName) = True
        Next WorkFolder
    Next AuthorFolder

    ' 삭제된 작품 찾기
    Dim RowsToDelete As New Collection
    For Each Key In ExistingWorks.Keys
        If Not CurrentWorks.Exists(Key) Then
            ' 삭제된 작품의 행 번호 가져오기
            RowNum = ExistingWorks(Key)
            RowsToDelete.Add RowNum
        End If
    Next Key

    ' 삭제된 작가 찾기 (작품보다 먼저 삭제해야 함)
    For Each Key In ExistingAuthors.Keys
        If Not CurrentAuthors.Exists(Key) Then
            ' 삭제된 작가의 행 번호 가져오기
            RowNum = ExistingAuthors(Key)
            ' 작가와 그 아래의 작품 행을 모두 삭제해야 함
            StartRow = RowNum

            ' 다음 작가의 행 번호 찾기
            EndRow = Cells(Rows.Count, 1).End(xlUp).Row
            For Each OtherKey In ExistingAuthors.Keys
                If ExistingAuthors(OtherKey) > StartRow Then
                    If ExistingAuthors(OtherKey) - 1 < EndRow Then
                        EndRow = ExistingAuthors(OtherKey) - 1
                    End If
                End If
            Next OtherKey

            ' 삭제할 행 번호를 모두 추가
            For i = StartRow To EndRow
                RowsToDelete.Add i
            Next i
        End If
    Next Key

    ' 중복 제거 및 내림차순 정렬 후 행 삭제
    Dim RowExistsFlag As Boolean
    For i = 1 To RowsToDelete.Count
        RowExistsFlag = False
        For Each RowItem In UniqueRows
            If RowItem = RowsToDelete(i) Then
                RowExistsFlag = True
                Exit For
            End If
        Next RowItem
        If Not RowExistsFlag Then
            UniqueRows.Add RowsToDelete(i)
        End If
    Next i

    ' 내림차순 정렬
    While UniqueRows.Count > 0
        MaxValue = 0
        For Each RowItem In UniqueRows
            If RowItem > MaxValue Then
                MaxValue = RowItem
            End If
        Next RowItem
        SortedRows.Add MaxValue
        ' Remove MaxValue from UniqueRows
        For i = 1 To UniqueRows.Count
            If UniqueRows(i) = MaxValue Then
                UniqueRows.Remove i
                Exit For
            End If
        Next i
    Wend

    ' 행 삭제
    For Each RowNum In SortedRows
        Rows(RowNum).Delete
    Next RowNum

    ' 행 삭제 후 ExistingAuthors와 ExistingWorks의 행 번호 업데이트
    For Each DeletedRow In SortedRows
        For Each Key In ExistingAuthors.Keys
            If ExistingAuthors(Key) > DeletedRow Then
                ExistingAuthors(Key) = ExistingAuthors(Key) - 1
            End If
        Next Key
        For Each Key In ExistingWorks.Keys
            If ExistingWorks(Key) > DeletedRow Then
                ExistingWorks(Key) = ExistingWorks(Key) - 1
            End If
        Next Key
    Next DeletedRow

    ' 추가된 작가 및 작품 추가
    ' 작가 추가
    For Each Key In CurrentAuthors.Keys
        If Not ExistingAuthors.Exists(Key) Then
            ' 마지막 행에 추가
            LastRow = Cells(Rows.Count, 1).End(xlUp).Row + 1
            Cells(LastRow, 1).Value = Key
            Cells(LastRow, 2).Value = Format(FileSystem.GetFolder(MainFolderPath & "\" & Key).DateLastModified, "yyyy-mm-dd hh:mm:ss")
            ' 서식 적용
            Cells(LastRow, 1).Font.Bold = True
            Cells(LastRow, 1).Interior.Color = RGB(220, 240, 255)
            Cells(LastRow, 1).HorizontalAlignment = xlCenter
            Cells(LastRow, 2).HorizontalAlignment = xlCenter
            ExistingAuthors(Key) = LastRow
        End If
    Next Key

    ' 작품 추가
    For Each Key In CurrentWorks.Keys
        If Not ExistingWorks.Exists(Key) Then
            ' Key 형식은 "AuthorName\WorkName"
            AuthorName = Split(Key, "\")(0)
            WorkName = Split(Key, "\")(1)
            ' 작가의 행 번호 찾기
            If ExistingAuthors.Exists(AuthorName) Then
                RowNum = ExistingAuthors(AuthorName)
                ' 작품을 작가 아래에 추가
                Dim InsertRow As Long
                InsertRow = RowNum + 1
                ' 기존 작품 수만큼 행 번호 증가
                Do While Cells(InsertRow, 3).Value <> ""
                    InsertRow = InsertRow + 1
                Loop
                ' 행 삽입
                Rows(InsertRow).Insert Shift:=xlDown
                Cells(InsertRow, 3).Value = WorkName
                Cells(InsertRow, 3).Interior.Color = RGB(245, 245, 245)
                Cells(InsertRow, 3).Font.Italic = True
                Cells(InsertRow, 3).HorizontalAlignment = xlLeft
                ExistingWorks(Key) = InsertRow
                ' 행 삽입 후 ExistingAuthors와 ExistingWorks의 행 번호 업데이트
                For Each ExistingKey In ExistingAuthors.Keys
                    If ExistingAuthors(ExistingKey) >= InsertRow And ExistingKey <> AuthorName Then
                        ExistingAuthors(ExistingKey) = ExistingAuthors(ExistingKey) + 1
                    End If
                Next ExistingKey
                For Each ExistingKey In ExistingWorks.Keys
                    If ExistingWorks(ExistingKey) >= InsertRow And ExistingKey <> Key Then
                        ExistingWorks(ExistingKey) = ExistingWorks(ExistingKey) + 1
                    End If
                Next ExistingKey
            End If
        End If
    Next Key

    ' 그룹화 적용
    ' 모든 작가에 대해 그룹화
    For Each Key In ExistingAuthors.Keys
        RowNum = ExistingAuthors(Key)
        ' 작품의 끝 행 번호 찾기
        StartRow = RowNum + 1
        EndRow = StartRow
        Do While Cells(EndRow, 3).Value <> ""
            EndRow = EndRow + 1
            If EndRow > Rows.Count Then Exit Do  ' 안전장치
        Loop
        EndRow = EndRow - 1
        If EndRow >= StartRow Then
            Rows(StartRow & ":" & EndRow).Rows.Group
        End If
    Next Key

    ' 그룹화 기본 설정: 축소된 상태로 표시
    ActiveSheet.Outline.ShowLevels RowLevels:=1

    ' 열 너비 자동 조정 및 수정일 열 너비 최적화
    Columns("A").AutoFit
    Columns("B").ColumnWidth = 20
    Columns("C").ColumnWidth = 120

    Application.ScreenUpdating = True  ' 화면 업데이트 재개

    MsgBox "폴더 구조의 변경 사항이 업데이트되었습니다."
End Sub

Function FolderExists(FolderPath As String) As Boolean
    FolderExists = (Dir(FolderPath, vbDirectory) <> "") And (GetAttr(FolderPath) And vbDirectory) = vbDirectory
End Function

