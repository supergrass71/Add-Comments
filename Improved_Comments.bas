Attribute VB_Name = "Improved_Comments"
Option Explicit

Sub AddComment()
Attribute AddComment.VB_ProcData.VB_Invoke_Func = "t\n14"
'
' AddComment just shows the cells I have changed
' updated to offer optional common comment text
' updated to use commentthreaded ("modern comments")20240312

Dim cmt As CommentThreaded
Dim commentCells As Range, cell As Range, firstItemInRange As Range
Dim commentTime As String, userName As String
Dim commentText As String
Dim answer As Integer

userName = GetUserFullName
commentTime = Format(Now(), "dd/mm/yy hh:mm AM/PM")

Set commentCells = Selection
Set firstItemInRange = Cells(commentCells.Row, commentCells.Column)
'MsgBox firstItemInRange.Address
'determine if cell has existing comment
Set cmt = firstItemInRange.CommentThreaded
If Not cmt Is Nothing Then
    answer = MsgBox(Prompt:="Overwrite comment?", Title:="Comment Writer", Buttons:=vbYesNo)
    If answer = vbYes Then
        commentText = Application.InputBox(Prompt:="Add new text to all comment cells", Title:="Add Comments to Selected Cells", Type:=2)
        'prevent error from blank inputbox
        If commentText = "False" Then
            commentText = "" 'result of clicking Cancel on InputBox
            Exit Sub
        End If
        For Each cell In commentCells
            Set cmt = cell.CommentThreaded
            If cmt Is Nothing Then
                cell.AddCommentThreaded Text:=commentText
            Else
                cell.CommentThreaded.Delete
                cell.AddCommentThreaded Text:=commentText
            End If
            Set cmt = Nothing
        Next cell
    Else
        commentText = cmt.Text
        For Each cell In commentCells
            Set cmt = cell.CommentThreaded
            If cmt Is Nothing Then
                cell.AddCommentThreaded Text:=commentText
            Else
                cell.CommentThreaded.Delete
                cell.AddCommentThreaded Text:=commentText
            End If
            Set cmt = Nothing
        Next cell
    End If
    Exit Sub 'end of non-blank comment case
End If

'case of no comment in cells
commentText = Application.InputBox(Prompt:="Add new text to all comment cells", Title:="Add Comments to Selected Cells", Type:=2)
If commentText = "False" Then
    commentText = "" 'result of clicking Cancel on InputBox
    Exit Sub
End If
For Each cell In commentCells
    Set cmt = cell.CommentThreaded
    If cmt Is Nothing Then
        cell.AddCommentThreaded Text:=commentText
    Else
        cell.CommentThreaded.Delete
        cell.AddCommentThreaded Text:=commentText
    End If
    Set cmt = Nothing
Next cell

End Sub

Sub testCommentstatements()
Dim cmt As CommentThreaded

With Selection
Set cmt = .CommentThreaded

    If cmt Is Nothing Then
        MsgBox "no comment!!"
    Else
        MsgBox cmt.Text
        Set cmt = Nothing
    End If
End With
End Sub

Sub reset_box_size()
'https://stackoverflow.com/questions/45515769/resize-excel-comments-to-fit-text-with-specific-width
Dim pComment As Comment
Dim lArea As Double
For Each pComment In Application.ActiveSheet.Comments
    With pComment.Shape

        .TextFrame.AutoSize = True

        lArea = .Width * .Height
        
        'only resize the autosize if width is above 300
        If .Width > 300 Then .Height = (lArea / .Width)       ' used .width so that it is less work to change final width

        
        .TextFrame.AutoMargins = False
        .TextFrame.MarginBottom = 0      ' margins need to be tweaked
        .TextFrame.MarginTop = 0
        .TextFrame.MarginLeft = 0
        .TextFrame.MarginRight = 0
        End With
Next

End Sub

Sub reset_Comment_size(pComment As Comment)
Dim lArea As Double
With pComment.Shape

    .TextFrame.AutoSize = True

    lArea = .Width * .Height
    
    'only resize the autosize if width is above 300
    If .Width > 300 Then .Height = (lArea / .Width)       ' used .width so that it is less work to change final width
    .TextFrame.AutoMargins = False
    .TextFrame.MarginBottom = 0      ' margins need to be tweaked
    .TextFrame.MarginTop = 0
    .TextFrame.MarginLeft = 0
    .TextFrame.MarginRight = 0
End With

End Sub
