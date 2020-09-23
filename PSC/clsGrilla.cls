VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author:  Erick Antelo
'Date:    04/22/2002
'email:   erickantelo@hotmail.com

Dim WithEvents mGrilla As TDBGrid
Attribute mGrilla.VB_VarHelpID = -1
Public rs As ADODB.Recordset

Dim LastOrdered As Long
Dim LastOrder As String
Dim ColumnaAOrdenar As Long
Dim Col As TrueOleDBGrid70.Column
Dim cols As TrueOleDBGrid70.Columns

Private mrsClone As ADODB.Recordset

'Public Property Set rs(adoRecordSet As ADODB.Recordset)
'    Set mRS = adoRecordSet
'    Set mrsClone = mRS.Clone
'End Property

Public Property Set Grilla(g As TDBGrid)
    Set mGrilla = g
    mGrilla.MultiSelect = dbgMultiSelectExtended
    mGrilla.ScrollTips = True
    mGrilla.ScrollTrack = True
End Property


Private Sub mGrilla_DblClick()
On Error GoTo err_h1:
Dim strOrden As String
Dim strDataField As String
    
    strDataField = mGrilla.Columns(ColumnaAOrdenar).DataField
    
    If LastOrdered = mGrilla.SelStartCol Then
        If LastOrder = "ASC" Then
            LastOrder = "DESC"
        Else
            LastOrder = "ASC"
        End If
        
        strOrden = strDataField & " " & LastOrder
    Else
        LastOrdered = mGrilla.SelStartCol
        strOrden = strDataField & " Asc"
    End If
    
    rs.Sort = strOrden
    
    Exit Sub
err_h1:
    MsgBox Err.Description

End Sub

Private Sub mGrilla_FetchScrollTips(ByVal SplitIndex As Integer, ByVal ColIndex As Integer, Bookmark As Variant, ByVal ScrollBar As TrueOleDBGrid70.ScrollBarsConstants, ScrollTip As String, ByVal TipStyle As TrueOleDBGrid70.StyleDisp)
' Gets called when scrolling occurs.
' Set the ScrollTip depending on which
' scroll bar was moved

    Select Case ScrollBar
    Case dbgVertical:
        'rs.Bookmark = Bookmark
       ' ScrollTip = "Record: " & CStr(rs.AbsolutePosition) & " de " & CStr(rs.RecordCount)
        ScrollTip = CStr(rs.RecordCount)
    Case dbgHorizontal:
        ScrollTip = mGrilla.Columns(ColIndex).Caption
    End Select

    TipStyle.ForeColor = vbBlue

End Sub

Private Sub mGrilla_GroupColMove(ByVal Position As Integer, ByVal ColIndex As Integer, Cancel As Integer)
Dim strSort As String
Dim Col As TrueOleDBGrid70.Column
' Loop through GroupColumns collection and construct
' the sort string for the Sort property of the Recordset

    For Each Col In mGrilla.GroupColumns
        If strSort <> vbNullString Then
            strSort = strSort & ", "
        End If
        strSort = strSort & "[" & Col.DataField & "]"
    Next Col

    mGrilla.HoldFields
    rs.Sort = strSort


End Sub

Private Sub mGrilla_HeadClick(ByVal ColIndex As Integer)
    ColumnaAOrdenar = ColIndex
End Sub

Private Sub mGrilla_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And mGrilla.FilterActive Then
            FilterChange
    End If

    If KeyCode = vbKeyEscape Then
    'If the user hit the  key and this column hasn't been changed,
    'then discard all changed to this record.
        If Not mGrilla.Columns(mGrilla.Col).DataChanged Then
            mGrilla.DataChanged = False      ' Undo entire record.
            KeyCode = 0                       ' Throw away the  key.

            ' In case there is an error updating value to database
            ' the recordset is left in edit mode - so cancel it
            If rs.EditMode <> dbEditNone Then
                rs.CancelUpdate
            End If
        End If
    End If
    
   If Shift = 2 Then
        Select Case KeyCode
        Case 67  'C
           If mGrilla.SelRange Then
            Copiar   'Copia lo seleccionado
           Else
            CopiarColumnas  'Copia las columnas seleccionadas
           End If
        Case 70  'F
            Find
        Case 66 'B muestra u oculta el filterbar
            mGrilla.FilterBar = Not mGrilla.FilterBar
        Case 71  'G
            If mGrilla.DataView = 0 Then  'No puede tomar el valor de 1 ya que es Hierarchical y no permite adicionar
                mGrilla.DataView = 2
            ElseIf mGrilla.DataView = 4 Then
                mGrilla.DataView = 0
            Else
                mGrilla.DataView = mGrilla.DataView + 1
            End If
        End Select
   End If
End Sub

Private Sub FilterChange()
'Gets called when an action is performed on the filter bar
On Error GoTo errHandler

Set cols = mGrilla.Columns
Dim c As Integer

c = mGrilla.Col
mGrilla.HoldFields
rs.Filter = getFilter(Col, cols, rs)

mGrilla.Col = c
mGrilla.EditActive = True

Exit Sub

errHandler:
    MsgBox Err.Source & ":" & vbCrLf & Err.Description

End Sub

Public Sub Copiar()
Dim adoRs As ADODB.Recordset
Dim strTemp As String  'string to be copied onto clipboard
Dim Col As Integer, row As Variant

If mGrilla.SelRange Then

'You must format the string so it can be pasted directly
'into Excel (tab delimited)
    Set adoRs = rs.Clone
        adoRs.Sort = rs.Sort
        adoRs.Filter = rs.Filter
        
    For Each row In mGrilla.SelBookmarks
        adoRs.Bookmark = row
        For Col = mGrilla.SelStartCol To mGrilla.SelEndCol
            strTemp = strTemp & adoRs(Col).Value & vbTab
        Next Col
        strTemp = strTemp & vbCrLf
    Next row

    Clipboard.Clear
    Clipboard.SetText strTemp, vbCFText
    Set adoRs = Nothing
    'MsgBox "Range of " & CStr(Abs(TDBGrid1.SelEndCol - TDBGrid1.SelStartCol) + 1) _
        & " x " & CStr(TDBGrid1.SelBookmarks.Count) & _
        " cells have been copied to the clipboard in TAB delimited format"
Else
    MsgBox "Por favor seleccione un rango de celdas"
End If

End Sub


Private Sub mGrilla_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If mGrilla.RowContaining(Y) = -1 And _
       mGrilla.ColContaining(X) = -1 Then
        mGrilla.PostMsg 1
    End If
    
End Sub

'Permite utilizar la tecla shift para seleccionar

'Private Sub mGrilla_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   'If Shift = vbShiftMask And mGrilla.SelBookmarks.Count > 0 Then'

      'Dim RSClone As Recordset
 '     Dim StartRow As Long
  '     Dim EndRow As Long
 '     Dim tmpBM As Variant
        
      ' Create clone of recordset
 '     Set RSClone = rs.Clone

      ' Position clone to shift-click record
      'tmpBM = mGrilla.RowBookmark(mGrilla.RowContaining(Y))
'      If tmpBM <> "" Then
'      RSClone.Bookmark = tmpBM
 '      End If

      ' Set EndRow equal to row number of shift-click record
 '     EndRow = RSClone.AbsolutePosition

      ' Position clone to start record
      'RSClone.Bookmark = mGrilla.SelBookmarks(0)

      ' Set StartRow equal to row number of start record
'      StartRow = RSClone.AbsolutePosition

      ' Clear currently selected records
'      For X = 0 To mGrilla.SelBookmarks.Count - 1
'         mGrilla.SelBookmarks.Remove (0)
'      Next X
        
'      If StartRow < EndRow Then

            ' StartRow is before the EndRow, so loop forward.
        ' Loop through clone, selecting all records between
            ' the StartRow and EndRow, inclusive
'            While CStr(RSClone.Bookmark) <> CStr(tmpBM)

               ' Add clone bookmark to SelBookmarks collection
'               mGrilla.SelBookmarks.Add RSClone.Bookmark
'               RSClone.MoveNext

'            Wend

 '     ElseIf StartRow > EndRow Then
 
         ' StartRow is after the EndRow, so loop backward.
         ' Loop through clone, selecting all records between
         ' the StartRow and EndRow, inclusive
'         While CStr(RSClone.Bookmark) <> CStr(tmpBM)

            ' Add clone bookmark to SelBookmarks collection
''            mGrilla.SelBookmarks.Add RSClone.Bookmark
'            RSClone.MovePrevious

 '        Wend

  '    End If

      ' Add final record's bookmark to SelBookmarks collection
 '     mGrilla.SelBookmarks.Add tmpBM

''   End If

''End Sub


'Private Sub mGrilla_SelChange(Cancel As Integer)

   ' Make sure that grid has focus when selecting records
 '  mGrilla.SetFocus

'End Sub

Private Sub mGrilla_PostEvent(ByVal MsgId As Integer)
    
    If MsgId = 1 Then
       If mGrilla.DataMode = 0 Then
           While Not rs.EOF
               mGrilla.SelBookmarks.Add (rs.Bookmark)
               rs.MoveNext
           Wend
       End If
     End If
        
End Sub

Private Sub CopiarColumnas()
   Dim tmpStr As String
    Dim GC As Columns
    Dim rsClone As Recordset, X As Integer
    
    Set rsClone = rs.Clone
        rsClone.Sort = rs.Sort
        rsClone.Filter = rs.Filter
        
    Set GC = mGrilla.Columns
    
    rsClone.MoveFirst
    tmpStr = ""
    If Not (mGrilla.SelStartCol < 0) And Not (mGrilla.SelEndCol < 0) Then
    While Not rsClone.EOF
        For X = mGrilla.SelStartCol To mGrilla.SelEndCol
            If X = mGrilla.SelStartCol Then
                tmpStr = tmpStr & rsClone.Fields(GC(X).DataField)
            Else
                tmpStr = tmpStr & vbTab & rsClone.Fields(GC(X).DataField)
            End If
        Next X
        tmpStr = tmpStr & Chr$(13)
        rsClone.MoveNext
    Wend
    End If
    Clipboard.Clear
    Clipboard.SetText tmpStr

End Sub

Public Sub SetExactWidth(frm As Form, Col As Integer)

    Dim rsClone As New ADODB.Recordset
    Dim Length As Integer, RsLen As Integer, RsCol As String
    Dim Bkmk As Variant
    Dim Value As String
    
    Set rsClone = rs.Clone

    rsClone.MoveFirst
    Value = mGrilla.Columns(Col).Caption
    Length = Len(Value)
    RsCol = mGrilla.Columns(Col).DataField
    While Not rsClone.EOF
        RsLen = IIf(IsNull(rsClone(RsCol).Value), 0, Len(rsClone(RsCol).Value))
        If RsLen > Length Then
            Length = RsLen
            Value = rsClone(RsCol).Value
            'Bkmk = rsClone.Bookmark
        End If
        rsClone.MoveNext
        DoEvents
    Wend
    'rsClone.Bookmark = Bkmk

    'mGrilla.Columns(Col).Width = frm.TextWidth(rsClone(RsCol).Value & Space(2))
    mGrilla.Columns(Col).Width = frm.TextWidth(Value & Space(2))
    
    Set rsClone = Nothing
End Sub

Public Sub Find()
Dim rsClone As ADODB.Recordset
    
    strTexto = InputBox("Ingrese texto a buscar :")
    If strTexto <> "" Then
        Set rsClone = rs.Clone
        rsClone.Filter = rs.Filter
        rsClone.Sort = rs.Sort
        rsClone.Bookmark = rs.Bookmark
        
        If rsClone.RecordCount > 0 And rsClone.EOF Then
            rsClone.MoveFirst
        End If
        
        If Not rsClone.EOF Then
            rsClone.Find mGrilla.Columns(mGrilla.Col).DataField & "='" & strTexto & "'", adSearchForward
            If Not rsClone.EOF Then
                rs.Bookmark = rsClone.Bookmark
            End If
        End If
    End If
    
End Sub
