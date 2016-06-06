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

Sub Play()
  If g_isPaused = True Then
    Call Pause
  Else
    Call ExcelTetris
  End If
End Sub

Sub Clear()
  Call ClearField
  Cells(INFO_PANEL_RESULT_Y, INFO_PANEL_RESULT_X).Select
  ActiveCell.FormulaR1C1 = 0
End Sub

Sub StopGame()
  g_isStoped = True
  
  If g_isPaused = True Then
    Call Pause
  End If
End Sub

Sub Pause()
  g_isPaused = Not g_isPaused
  
  Cells(FIELD_HEIGHT + 2, 1).Select
  
  If g_isPaused = True Then
    ActiveCell.FormulaR1C1 = "PAUSED"
  Else
    ActiveCell.FormulaR1C1 = ""
  End If
End Sub


