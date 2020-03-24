VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "sample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public x As String
'Public y As Object
Private m_x As String
Private m_y As Object
Private m_abc As Long

Private Property Let x(ByVal x_ As String)
  m_x = x_
End Property

Private Property Set y(y_ As Object)
  Set m_y = y_
End Property

Private Property Get x() As String
  x = m_x
End Property

Private Property Get y() As Object
  Set y = m_y
End Property

Private Property Set abc(ByVal abc_ As Long)
  Set m_abc = abc_
End Property

Private Property Let abc(ByVal abc_ As Long)
  m_abc = abc_
End Property
