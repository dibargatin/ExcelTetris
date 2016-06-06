'' ExcelTetris is a classical tetris game.
'' 
'' Copyright (c) 2016 Dmitriy Bargatin
'' 
'' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated 
'' documentation files (the "Software"), to deal in the Software without restriction, including without limitation 
'' the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and 
'' to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'' 
'' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'' 
'' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO 
'' THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS 
'' OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR 
'' OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Option Explicit

Public g_row As Integer
Public g_row_delta As Integer
Public g_col As Integer
Public g_col_delta As Integer
Public g_isDraw As Boolean
Public g_isStoped As Boolean
Public g_isPaused As Boolean

Private m_gameover As Boolean
Private m_field(FIELD_WIDTH, FIELD_HEIGHT) As Long
Private m_patterns As FigurePatternCollection

' Current figure params
Dim m_current_X As Integer
Dim m_current_Y As Integer
Dim m_current_direction As Integer
Dim m_current_fig_type As Integer
Dim m_current_fig_color As Long
Dim m_current_copy_to_field As Boolean

' Next figure params
Dim m_next_fig_direction As Integer
Dim m_next_fig_type As Integer
Dim m_next_fig_color As Long

Public Sub ExcelTetris()
  m_gameover = False
  Call init
  
  Dim collision As Boolean
  Dim i, nextDir, xCollision, yCollision As Integer
  
  For i = 1 To FIELD_HEIGHT + 1
      g_isDraw = False
      
      If DELAY > 0 Then _
        Application.Wait (DateAdd("s", DELAY, Now))
        
      DoEvents
      g_isDraw = True
      
      If g_isPaused = True Then
        i = i - 1
      Else
        nextDir = nextDirection(g_row_delta)
        collision = checkCollision(g_col, i, nextDir, xCollision, yCollision)
              
        If collision = False Then
          m_current_X = g_col
          m_current_Y = i
          m_current_direction = nextDir
        ElseIf yCollision <> -1 Then
          m_current_copy_to_field = True
        End If
        
        Call draw
        
        If m_current_copy_to_field Then
          Call processField
          
          m_current_fig_type = m_next_fig_type
          m_current_direction = m_next_fig_direction
          m_current_fig_color = m_next_fig_color
          
          Call nextFigure(m_next_fig_type, m_next_fig_direction, m_next_fig_color)
          
          m_current_X = Int(Round(FIELD_WIDTH * 0.5))
          i = 1
          m_current_Y = i
          m_current_copy_to_field = False
        End If
        
        If m_gameover = True Or g_isStoped = True Then GoTo GameOver
      End If
      
      g_col = m_current_X
      g_row = m_current_Y
      g_row_delta = 0
      g_col_delta = 0
  Next i

GameOver:
  If g_isStoped = False Then _
    MsgBox "Game over", vbInformation + vbOKOnly, "ExcelTetris"
End Sub

Private Sub processField()
  Dim fullRow(FIELD_HEIGHT) As Boolean
  Dim x, y As Integer
  
  For y = 1 To FIELD_HEIGHT
    fullRow(y) = True
  Next y
  
  For x = 1 To FIELD_WIDTH
      If m_field(x, 1) <> xlAutomatic Then
        m_gameover = True
        x = FIELD_WIDTH
      End If
      
      If m_gameover <> True Then
        For y = 1 To FIELD_HEIGHT
          If m_field(x, y) = xlAutomatic Then
            fullRow(y) = False
          End If
        Next y
      End If
  Next x
  
  If m_gameover <> True Then
    Dim count As Integer
    Dim row As Integer
    
    count = 0
    
    For row = 1 To FIELD_HEIGHT
      If fullRow(row) Then
        For x = 1 To FIELD_WIDTH
          For y = row To y = 1 Step -1
            m_field(x, y) = xlAutomatic
            count = count + 1
            
            If y - 1 >= 1 Then
              m_field(x, y) = m_field(x, y - 1)
            End If
          Next y
        Next x
      End If
    Next row
    
    Cells(INFO_PANEL_RESULT_Y, INFO_PANEL_RESULT_X).Select
    If Len(ActiveCell.FormulaR1C1) = 0 Then
      ActiveCell.FormulaR1C1 = count
    Else
      ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1 + count
    End If
    Cells(m_current_Y, m_current_X).Select
  End If
End Sub

Private Function checkCollision(ByRef x, y, dir, xResult, yResult As Integer) As Boolean
  Dim fig() As Integer
  Dim cX, cY, w, h, tx, ty As Integer
  
  Call m_patterns.getItem(m_current_fig_type).getFigurePattern(Int(dir), fig)
  cX = m_patterns.getItem(m_current_fig_type).getCenterX(Int(dir))
  cY = m_patterns.getItem(m_current_fig_type).getCenterY(Int(dir))
  w = UBound(fig, 1)
  h = UBound(fig, 2)
  
  xResult = -1
  yResult = -1
  checkCollision = False
  
  Dim i, k As Integer
  For i = 0 To w
    For k = 0 To h
      If fig(i, k) <> 0 Then
        tx = x + (i - cX)
        ty = y + (k - cY)
        
        If tx < 1 Then
          checkCollision = True
          xResult = max(i, xResult)
        ElseIf tx > FIELD_WIDTH Then
          checkCollision = True
          xResult = max(i, xResult)
        End If

        If ty >= FIELD_HEIGHT + 1 Then
          checkCollision = True
          yResult = max(k, yResult)
        End If

        If xResult = -1 Then
          If ty <= FIELD_HEIGHT Then
            If m_field(tx, ty) <> xlAutomatic Then
              checkCollision = True
              xResult = max(i, xResult)
              yResult = max(k, yResult)
            End If
          End If
        End If
      End If
    Next k
  Next i
End Function

Public Sub ClearField()
  Dim x, y As Integer
    For x = 1 To FIELD_WIDTH + INFO_PANEL_WIDTH
      For y = 1 To FIELD_HEIGHT + 3
        Cells(y, x).Select
        With Selection.Interior
          .Pattern = xlSolid
          .PatternColorIndex = xlAutomatic
          .color = xlAutomatic
          .TintAndShade = 0
          .PatternTintAndShade = 0
        End With
      Next y
    Next x
End Sub

Private Sub init()
  Call Randomize(Now)
  If m_patterns Is Nothing Then Set m_patterns = New FigurePatternCollection
  
  Call nextFigure(m_current_fig_type, m_current_direction, m_current_fig_color)
  Call nextFigure(m_next_fig_type, m_next_fig_direction, m_next_fig_color)
  
  m_current_X = Int(Round(FIELD_WIDTH * 0.5))
  m_current_Y = 1
  m_current_copy_to_field = False
  
  g_col = m_current_X
  g_row = Int(Round(FIELD_HEIGHT * 0.5))
  
  g_col_delta = 0
  g_row_delta = 0
  
  g_isStoped = False
  g_isPaused = False
  
  Dim x, y As Integer
  For x = 0 To FIELD_WIDTH
    For y = 0 To FIELD_HEIGHT
      m_field(x, y) = xlAutomatic
    Next y
  Next x
  
  Call ClearField
  Cells(m_current_Y, m_current_X).Select
End Sub

Private Function nextDirection(delta As Integer) As Integer
  If delta <> 0 Then
    nextDirection = Abs((m_current_direction + delta) Mod 4)
  Else
    nextDirection = m_current_direction
  End If
End Function

Private Sub nextFigure(ByRef figType, dir As Integer, ByRef color As Long)
  figType = Int((FIGURE_COUNT + 1) * Rnd)
  dir = Int((DIRECTION_COUNT + 1) * Rnd)
  
  Select Case Int(10 * Rnd)
    Case 0
      color = 192
    Case 1
      color = 255
    Case 2
      color = 49407
    Case 3
      color = 65535
    Case 4
      color = 5287936
    Case 5
      color = 5296274
    Case 6
      color = 15773696
    Case 7
      color = 12611584
    Case 8
      color = 6299648
    Case 9
      color = 10498160
  End Select
End Sub

Private Sub draw()
    Dim x, y As Integer
    For x = 1 To FIELD_WIDTH
      For y = 1 To FIELD_HEIGHT
        Cells(y, x).Select
        With Selection.Interior
          .Pattern = xlSolid
          .PatternColorIndex = xlAutomatic
          .color = m_field(x, y)
          .TintAndShade = 0
          .PatternTintAndShade = 0
        End With
      Next y
    Next x
            
    For x = FIELD_WIDTH To FIELD_WIDTH + INFO_PANEL_WIDTH
      For y = 1 To INFO_PANEL_RESULT_Y
        Cells(y, x).Select
        With Selection.Interior
          .Pattern = xlSolid
          .PatternColorIndex = xlAutomatic
          .color = xlAutomatic
          .TintAndShade = 0
          .PatternTintAndShade = 0
        End With
      Next y
    Next x
    
    Call drawFigure(m_current_X, m_current_Y, m_current_fig_type, m_current_direction, m_current_fig_color, True, m_current_copy_to_field)
    Call drawFigure(FIELD_WIDTH + 2, 2, m_next_fig_type, m_next_fig_direction, m_next_fig_color, False, False, True)
    
    Cells(m_current_Y, m_current_X).Select
End Sub

Private Sub drawFigure(ByVal x, y, figureType, direction As Integer, color As Long, Optional useCenterInfo As Boolean = True, Optional copyToField As Boolean = False, Optional isPanelFig As Boolean = False)
  Dim fig() As Integer
  Dim cX, cY, w, h As Integer
  
  Call m_patterns.getItem(figureType).getFigurePattern(direction, fig)
  
  w = UBound(fig, 1)
  h = UBound(fig, 2)
  
  If useCenterInfo = True Then
    cX = m_patterns.getItem(figureType).getCenterX(direction)
    cY = m_patterns.getItem(figureType).getCenterY(direction)
  Else
    cX = 0
    cY = 0
  End If
  
  Dim i, k, tx, ty As Integer
  For i = 0 To w
    For k = 0 To h
      tx = x + (i - cX)
      ty = y + (k - cY)
      
      If (tx <= FIELD_WIDTH And ty <= FIELD_HEIGHT) Or isPanelFig = True Then
        If tx >= 1 And ty >= 1 And fig(i, k) <> 0 Then
          Cells(ty, tx).Select
          With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .color = color
            .TintAndShade = 0
            .PatternTintAndShade = 0
          End With
          
          If copyToField Then
            m_field(tx, ty) = color
          End If
        End If
      End If
    Next k
  Next i
End Sub
