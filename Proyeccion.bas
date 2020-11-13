Attribute VB_Name = "Proyeccion"

Option Explicit

' Declaración de variables
Dim miPX As Double
Dim miPY As Double
Dim miDistancia As Double
Dim miEX As Double
Dim miEY As Double
Dim miEZ As Double

' Calcula la distancia del punto al centro
Public Function CalculaDistancia(ByVal pX As Double, ByVal pY As Double) As Double
  CalculaDistancia = Sqr((pX ^ 2) + (pY ^ 2))
End Function

' Calcula la componente X en la esfera
Public Function CalculaEX(ByVal pX As Double, ByVal pD As Double) As Double
  CalculaEX = (2 * pX) / (1 + (pD ^ 2))
End Function

' Calcula la componente Y en la esfera
Public Function CalculaEY(ByVal pY As Double, ByVal pD As Double) As Double
  CalculaEY = (2 * pY) / (1 + (pD ^ 2))
End Function

' Calcula la componente Z en la esfera
Public Function CalculaEZ(ByVal pD As Double) As Double
  CalculaEZ = ((pD ^ 2) - 1) / (1 + (pD ^ 2))
End Function

