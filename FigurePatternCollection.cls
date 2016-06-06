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

Private Items(FIGURE_COUNT) As FigurePattern

Private Sub Class_Initialize()
  Dim fig() As Integer
  
  ReDim fig(0, 0)
  fig(0, 0) = 1
  Set Items(FIGURE_DOT) = New FigurePattern
  Call Items(FIGURE_DOT).init(fig, 0, 0)
  
  ReDim fig(0, 1)
  fig(0, 0) = 1
  fig(0, 1) = 1
  Set Items(FIGURE_I_2) = New FigurePattern
  Call Items(FIGURE_I_2).init(fig, 0, 0)
    
  ReDim fig(1, 1)
  fig(0, 0) = 1
  fig(0, 1) = 1
  fig(1, 0) = 1
  fig(1, 1) = 0
  Set Items(FIGURE_L_3) = New FigurePattern
  Call Items(FIGURE_L_3).init(fig, 0, 0)
  
  ReDim fig(0, 2)
  fig(0, 0) = 1
  fig(0, 1) = 1
  fig(0, 2) = 1
  Set Items(FIGURE_I_3) = New FigurePattern
  Call Items(FIGURE_I_3).init(fig, 0, 1)
  
  ReDim fig(1, 1)
  fig(0, 0) = 1
  fig(0, 1) = 1
  fig(1, 0) = 1
  fig(1, 1) = 1
  Set Items(FIGURE_DOT_4) = New FigurePattern
  Call Items(FIGURE_DOT_4).init(fig, 0, 0)
  
  ReDim fig(0, 3)
  fig(0, 0) = 1
  fig(0, 1) = 1
  fig(0, 2) = 1
  fig(0, 3) = 1
  Set Items(FIGURE_I_4) = New FigurePattern
  Call Items(FIGURE_I_4).init(fig, 0, 1)
  
  ReDim fig(1, 2)
  fig(0, 0) = 1
  fig(0, 1) = 0
  fig(0, 2) = 0
  fig(1, 0) = 1
  fig(1, 1) = 1
  fig(1, 2) = 1
  Set Items(FIGURE_L_4) = New FigurePattern
  Call Items(FIGURE_L_4).init(fig, 0, 1)
  
  ReDim fig(1, 2)
  fig(0, 0) = 1
  fig(0, 1) = 1
  fig(0, 2) = 1
  fig(1, 0) = 1
  fig(1, 1) = 0
  fig(1, 2) = 0
  Set Items(FIGURE_RL_4) = New FigurePattern
  Call Items(FIGURE_RL_4).init(fig, 1, 1)
  
  ReDim fig(1, 2)
  fig(0, 0) = 1
  fig(0, 1) = 1
  fig(0, 2) = 1
  fig(1, 0) = 0
  fig(1, 1) = 1
  fig(1, 2) = 0
  Set Items(FIGURE_T_4) = New FigurePattern
  Call Items(FIGURE_T_4).init(fig, 1, 1)
  
  ReDim fig(1, 2)
  fig(0, 0) = 0
  fig(0, 1) = 1
  fig(0, 2) = 1
  fig(1, 0) = 1
  fig(1, 1) = 1
  fig(1, 2) = 0
  Set Items(FIGURE_Z_4) = New FigurePattern
  Call Items(FIGURE_Z_4).init(fig, 1, 1)
  
  ReDim fig(1, 2)
  fig(0, 0) = 1
  fig(0, 1) = 1
  fig(0, 2) = 0
  fig(1, 0) = 0
  fig(1, 1) = 1
  fig(1, 2) = 1
  Set Items(FIGURE_S_4) = New FigurePattern
  Call Items(FIGURE_S_4).init(fig, 1, 1)
End Sub

Public Function getItem(ByVal figType As Integer) As FigurePattern
  Set getItem = Items(figType)
End Function
