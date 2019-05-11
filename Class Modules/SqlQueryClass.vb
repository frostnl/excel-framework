Option Explicit
Private pSql As String
Private pBaseSql As String
Private pJoinClauses As Collection
Private pWhereClauses As Collection
Private pGroupByClause As String
Private pFromClause As String

Private Sub Class_Initialize()
    

End Sub

Public Function BaseQuery(sql As String)
    pBaseSql = sql
End Function

Public Function FromTable(sql As String)
    pFromClase = sql
End Function

Public Function AddJoin(sql As String)
    pJoinClauses.Add sql
End Function

Public Function AddWhere(sql As String)
    pWhereClauses.Add sql
End Function

Public Function GroupByClause(sql As String)
    pGroupByClase = sql
End Function

Public Function GetQuery()
    Dim sql As String
    Dim fromClause As String
    Dim var As Variant
        
    Dim andClause As String
    
    sql = pBaseSql
    
    'from clause
    fromClause = " " & pFromClause
    
    'joins
    For Each var In pJoinClauses
        fromClause = "(" & fromClause & " AND " & CStr(var) & ")"
    End If
    

    
    
End Function
