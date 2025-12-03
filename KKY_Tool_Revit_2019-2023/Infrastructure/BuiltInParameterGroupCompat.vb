Imports Autodesk.Revit.DB
Imports System.Runtime.CompilerServices

Namespace Global.Autodesk.Revit.DB

    ''' <summary>
    ''' Revit 2019~2023용 BuiltInParameterGroup 호환 헬퍼.
    ''' 정의에서 그룹을 조회/비교하는 공통 API를 제공한다.
    ''' </summary>
    Public Module BuiltInParameterGroupCompat

        Public Function GetGroup(def As Definition) As BuiltInParameterGroup
            If def Is Nothing Then
                Return BuiltInParameterGroup.PG_DATA
            End If

            Return def.ParameterGroup
        End Function

        <Extension>
        Public Function IsInGroup(def As Definition, group As BuiltInParameterGroup) As Boolean
            If def Is Nothing Then Return False
            Return GetGroup(def) = group
        End Function

    End Module

End Namespace
