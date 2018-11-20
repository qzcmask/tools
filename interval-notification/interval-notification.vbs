' This program is to build a interval notification based on Windows System.
' It originally aimed to notify me to drink water in a interval.
' @author Mask
' @createdAt 2018/11/20

Dim interval
Dim ret

' set task's interval
Do
  interval = InputBox("Please input the interval:", "Interval(seconds)", 30 * 60 * 1000)

  If interval = vbEmpty Then
    ' click cancel
    MsgBox "You input empty, interval is set default half an hour."
    interval = 1800000
    Exit Do
  End If

  If Not IsNumeric(interval) Then
    ' input what is't a numeric
     MsgBox "Please input number"
   Else
     Exit Do
  End If
Loop

' execute interval task
Do
  ret = MsgBox("Drink some water", vbOKCancel, "Time up")
  If ret = 1 Then
    Wscript.sleep interval
  Else
    ' clear interval task
    MsgBox "Close notification", vbOkOnly, "Close notification"
    Exit Do
  End If
Loop