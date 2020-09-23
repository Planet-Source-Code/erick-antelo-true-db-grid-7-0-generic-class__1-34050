Attribute VB_Name = "mdlFiltro"
'Author:  Erick Antelo
'Date:    04/22/2002
'email:   erickantelo@hotmail.com

Public Function getFilter(Col As TrueOleDBGrid70.Column, cols As TrueOleDBGrid70.Columns, rs As ADODB.Recordset) As String
    'Creates the SQL statement in adodc1.recordset.filter
    'and only filters text currently. It must be modified to
    'filter other data types.

    Dim tmp As String
    Dim n As Integer
    Dim mes As String, anno As String
    Dim inimes As Date, finmes As Date
    
    For Each Col In cols
        
        If Trim(Col.FilterText) <> "" And (Col.DataField <> "") Then
            n = n + 1
            If n > 1 Then
                tmp = tmp & " AND "
            End If
            
            Select Case rs.Fields(Col.DataField).Type
            Case adDate, adDBDate, adDBTime, adDBTimeStamp
                If Col.FilterText = "!" Then
                  
                        tmp = tmp & "(((" & Col.DataField & ") is null) or (" & Col.DataField & ") = '')"
                Else
                    If IsDate(Col.FilterText) Then
                        tmp = tmp & Col.DataField & " = #" & CDate(Col.FilterText) & "#"
                    Else
                        If Left(Col.FilterText, 1) = "#" Then
                            mes = Mid(Col.FilterText, 2, InStr(1, Col.FilterText, "/") - 2)
                            anno = Right(Col.FilterText, -InStr(1, Col.FilterText, "/") + Len(Col.FilterText))
                            inimes = DateSerial(anno, mes, 1)
                            finmes = DateSerial(anno, mes, UltimoDiaMes(inimes))
                           tmp = tmp & "[" & Col.DataField & "] >= #" & inimes & "# and [" & Col.DataField & "] <= #" & finmes & "#"
                        Else
                            tmp = tmp & Col.DataField & " " & Col.FilterText
                        End If
                            
                    End If
                End If
            Case adDouble, adInteger, adNumeric
                If IsNumeric(Col.FilterText) Then
                    tmp = tmp & "[" & Col.DataField & "] = " & Col.FilterText
                Else
                    tmp = tmp & "[" & Col.DataField & "] = " & Col.FilterText
                End If
            Case adChar, adVarChar, adVarWChar
                If Col.FilterText = "!" Then
                        tmp = tmp & "((" & Col.DataField & ") = null or (" & Col.DataField & ") = '')"
                Else
                    'If Right(Col.FilterText, 1) <> "*" Then
                        'If Left(Col.FilterText, 1) = "!" Then
                        '    tmp = tmp & Col.DataField & " NOT LIKE '" & Mid(Col.FilterText, 2) & "*'"
                        'Else
                    '        tmp = tmp & Col.DataField & " LIKE '" & Col.FilterText & "*'"
                        'End If
                    'Else
                        'If Left(Col.FilterText, 1) = "!" Then
                         '   tmp = tmp & Col.DataField & " NOT LIKE '" & Col.FilterText & "'"
                        'Else
                            tmp = tmp & Col.DataField & " LIKE '" & Col.FilterText & "'"
                        'End If
                    'End If
                End If
            Case adBoolean
                If IsNumeric(Col.FilterText) Then
                    tmp = tmp & Col.DataField & " = " & Col.FilterText
                End If
            End Select
        End If
    Next Col

    getFilter = tmp

End Function

Function UltimoDiaMes(ByVal Fecha As Date) As Long
Dim SiguienteMes As Long
Dim res As Date
    
    mes = Month(Fecha)
    
    If mes = 12 Then
        SiguienteMes = 1
    Else
        SiguienteMes = mes + 1
    End If
    
    res = DateSerial(Year(Fecha), SiguienteMes, 1) - 1
    
    UltimoDiaMes = Day(res)
    
End Function
