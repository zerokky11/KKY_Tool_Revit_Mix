Imports Autodesk.Revit.DB
Imports System.Runtime.CompilerServices

Namespace Infrastructure

    ''' <summary>
    ''' ElementId 관련 버전 호환 헬퍼
    ''' - 2019~2023: IntegerValue / New ElementId(Integer)
    ''' - 2025: Value / New ElementId(Long)
    ''' </summary>
    Public Module ElementIdCompat

        ''' <summary>
        ''' ElementId를 Int64로 꺼내는 공용 함수
        ''' </summary>
        <Extension>
        Public Function GetLongId(id As ElementId) As Long
            If id Is Nothing Then Return -1
#If REVIT2025 Then
            Return id.Value
#Else
            Return CLng(id.IntegerValue)
#End If
        End Function

        ''' <summary>
        ''' ElementId를 Int32로 꺼내는 공용 함수
        ''' </summary>
        <Extension>
        Public Function GetIntId(id As ElementId) As Integer
            Return CInt(GetLongId(id))
        End Function

        ''' <summary>
        ''' (기존 호환) ElementId를 Int32로 꺼내는 공용 함수
        ''' </summary>
        <Extension>
        Public Function IntValue(id As ElementId) As Integer
            Return GetIntId(id)
        End Function

        ''' <summary>
        ''' Int32에서 ElementId 생성 (버전별 생성자 호환)
        ''' </summary>
        Public Function FromInt(id As Integer) As ElementId
#If REVIT2025 Then
            Return New ElementId(CLng(id))
#Else
            Return New ElementId(id)
#End If
        End Function

        ''' <summary>
        ''' Int64에서 ElementId 생성 (버전별 생성자 호환)
        ''' </summary>
        Public Function FromLong(id As Long) As ElementId
#If REVIT2025 Then
            Return New ElementId(id)
#Else
            Return New ElementId(CInt(id))
#End If
        End Function

    End Module

End Namespace
