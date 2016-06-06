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

Public Const DELAY As Integer = 0

Public Const FIELD_WIDTH As Integer = 10
Public Const FIELD_HEIGHT As Integer = 20

Public Const INFO_PANEL_WIDTH As Integer = 6
Public Const INFO_PANEL_RESULT_X As Integer = FIELD_WIDTH + 3
Public Const INFO_PANEL_RESULT_Y As Integer = 7

Public Const DIRECTION_COUNT As Integer = 3
Public Const DIRECTION_LEFT As Integer = 0
Public Const DIRECTION_TOP As Integer = 1
Public Const DIRECTION_RIGHT As Integer = 2
Public Const DIRECTION_BOTTOM As Integer = 3

Public Const FIGURE_COUNT As Integer = 10
Public Const FIGURE_DOT As Integer = 0
Public Const FIGURE_I_2 As Integer = 1
Public Const FIGURE_L_3 As Integer = 2
Public Const FIGURE_I_3 As Integer = 3
Public Const FIGURE_DOT_4 As Integer = 4
Public Const FIGURE_I_4 As Integer = 5
Public Const FIGURE_L_4 As Integer = 6
Public Const FIGURE_RL_4 As Integer = 7
Public Const FIGURE_T_4 As Integer = 8
Public Const FIGURE_Z_4 As Integer = 9
Public Const FIGURE_S_4 As Integer = 10

Public Function max(num01, num02 As Variant) As Variant
  max = IIf(num01 > num02, num01, num02)
End Function

Public Function min(num01, num02 As Variant) As Variant
  max = IIf(num01 < num02, num01, num02)
End Function
