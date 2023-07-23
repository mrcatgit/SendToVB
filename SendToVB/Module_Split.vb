Module Module_Split
    Public Function Split(ByVal expression As String, ByVal delimiter As String, ByVal qualifier As String, ByVal ignoreCase As Boolean) As String()
        ' Based on the work of LSteinle
        ' http://www.codeproject.com/KB/dotnet/TextQualifyingSplit.aspx?fid=336054&select=1797240&fr=1#xx0xx

        Dim _QualifierState As Boolean = False
        Dim _StartIndex As Integer = 0
        Dim _Values As New System.Collections.ArrayList

        For _CharIndex As Integer = 0 To expression.Length - 1
            If Not qualifier Is Nothing AndAlso String.Compare(expression.Substring(_CharIndex, qualifier.Length), qualifier, ignoreCase) = 0 Then
                _QualifierState = Not _QualifierState
            ElseIf Not _QualifierState AndAlso Not delimiter Is Nothing AndAlso String.Compare(expression.Substring(_CharIndex, delimiter.Length), delimiter, ignoreCase) = 0 Then
                _Values.Add(expression.Substring(_StartIndex, _CharIndex - _StartIndex))
                _StartIndex = _CharIndex + 1
            End If
        Next

        If _StartIndex < expression.Length Then _Values.Add(expression.Substring(_StartIndex, expression.Length - _StartIndex))

        Dim _returnValues(_Values.Count - 1) As String
        _Values.CopyTo(_returnValues)
        Return _returnValues
    End Function
End Module
