VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PPrintProject.Examples"
Option Explicit

Public Property Get CurrentTime() As String
    CurrentTime = Strings.Format(DateTime.Now, "\[ddmmyy\] hh:nn")
End Property

Public Function Repr__() As String
    Repr__ = Information.TypeName(Me) & "(CurrentTime=" & Me.CurrentTime & ")"
End Function
