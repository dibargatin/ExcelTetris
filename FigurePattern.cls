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

Private m_centerX(DIRECTION_COUNT) As Integer
Private m_centerY(DIRECTION_COUNT) As Integer

Private m_figureLeft() As Integer
Private m_figureTop() As Integer
Private m_figureRight() As Integer
Private m_figureBottom() As Integer

Public Sub init(figure() As Integer, centerX As Integer, centerY As Integer)
  m_figureLeft = figure
  
  Call tran(m_figureLeft, m_figureBottom)
  Call tran(m_figureBottom, m_figureRight)
  Call tran(m_figureRight, m_figureTop)
  
  m_centerX(DIRECTION_LEFT) = centerX
  m_centerY(DIRECTION_LEFT) = centerY
  
  m_centerX(DIRECTION_TOP) = UBound(m_figureLeft, 2) - m_centerY(DIRECTION_LEFT)
  m_centerY(DIRECTION_TOP) = m_centerX(DIRECTION_LEFT)
  
  m_centerX(DIRECTION_RIGHT) = UBound(m_figureLeft, 1) - m_centerX(DIRECTION_LEFT)
  m_centerY(DIRECTION_RIGHT) = UBound(m_figureLeft, 2) - m_centerY(DIRECTION_LEFT)
  
  m_centerX(DIRECTION_BOTTOM) = UBound(m_figureTop, 1) - m_centerX(DIRECTION_TOP)
  m_centerY(DIRECTION_BOTTOM) = UBound(m_figureTop, 2) - m_centerY(DIRECTION_TOP)
End Sub

Public Sub getFigurePattern(direction As Integer, ByRef res() As Integer)
  Select Case direction
    Case DIRECTION_LEFT
      res = m_figureLeft
    Case DIRECTION_TOP
      res = m_figureTop
    Case DIRECTION_RIGHT
      res = m_figureRight
    Case DIRECTION_BOTTOM
      res = m_figureBottom
  End Select
End Sub

Public Function getCenterX(direction As Integer) As Integer
  getCenterX = m_centerX(direction)
End Function

Public Function getCenterY(direction As Integer) As Integer
  getCenterY = m_centerY(direction)
End Function

Private Sub tran(fig() As Integer, ByRef res() As Integer)
  Dim x, y As Integer
  Dim w, h As Integer
      
  w = UBound(fig, 2)
  h = UBound(fig, 1)
    
  ReDim res(w, h)
    
  For x = 0 To w
    For y = 0 To h
      res(x, y) = fig(h - y, x)
    Next y
  Next x
End Sub
