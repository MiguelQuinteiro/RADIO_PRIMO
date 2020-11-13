VERSION 5.00
Begin VB.Form frmRadioPrimo 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000008&
   Caption         =   "Estudio: Frecuencia y distribución de los números primos"
   ClientHeight    =   10755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18960
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10755
   ScaleWidth      =   18960
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAviso 
      Height          =   495
      Left            =   16200
      TabIndex        =   27
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000006&
      Caption         =   "Descripción "
      Height          =   3735
      Left            =   14280
      TabIndex        =   11
      Top             =   5400
      Width           =   4575
      Begin VB.TextBox Text1 
         Height          =   3135
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   12
         Text            =   "frmRadioPrimo.frx":0000
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000006&
      Caption         =   "Controles "
      Height          =   5055
      Left            =   14280
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.CommandButton Command3 
         Caption         =   "PG"
         Height          =   495
         Left            =   3120
         TabIndex        =   26
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtPrimosGemelos 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   3480
         TabIndex        =   25
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdEstadisticaOrbita 
         Caption         =   "Estadística Órbitas"
         Height          =   495
         Left            =   3120
         TabIndex        =   24
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "C"
         Height          =   495
         Left            =   3120
         TabIndex        =   23
         Top             =   4320
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "P"
         Height          =   495
         Left            =   1680
         TabIndex        =   22
         Top             =   4320
         Width           =   375
      End
      Begin VB.CommandButton cmdPrimoAnterior 
         Caption         =   "Primo Anterior"
         Height          =   495
         Left            =   1680
         TabIndex        =   21
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdListaOrbitas 
         Caption         =   "Lista Órbitas"
         Height          =   495
         Left            =   1680
         TabIndex        =   20
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox txtNumeroOrbita 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   240
         TabIndex        =   19
         Text            =   "2"
         Top             =   4320
         Width           =   1215
      End
      Begin VB.TextBox txtOrbitaCompuesto 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   3480
         TabIndex        =   18
         Top             =   4320
         Width           =   855
      End
      Begin VB.CommandButton cmdInformacionOrbita 
         Caption         =   "Información Órbita"
         Height          =   495
         Left            =   240
         TabIndex        =   17
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox txtOrbitaPrimo 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   2040
         TabIndex        =   16
         Top             =   4320
         Width           =   855
      End
      Begin VB.CommandButton cmdPrimoSiguiente 
         Caption         =   "Primo Siguiente"
         Height          =   495
         Left            =   1680
         TabIndex        =   15
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton cmdUne 
         Caption         =   "Líneas Gemelos"
         Height          =   495
         Left            =   3120
         TabIndex        =   14
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton cmdEjes 
         Caption         =   "Ejes"
         Height          =   495
         Left            =   3120
         TabIndex        =   13
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdMitad 
         Caption         =   "Mitad"
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdDoble 
         Caption         =   "Doble"
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtN 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   1680
         TabIndex        =   8
         Text            =   "2"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdGrafica 
         Caption         =   "Mostrar"
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdLineasCompuestos 
         Caption         =   "Líneas Compuestos"
         Height          =   495
         Left            =   1680
         TabIndex        =   6
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton cmdLineas 
         Caption         =   "Líneas Primos"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton cmdCompuestos 
         Caption         =   "Compuestos"
         Height          =   495
         Left            =   1680
         TabIndex        =   4
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdEtiqueta 
         Caption         =   "Etiqueta"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdCerca 
         Caption         =   "Cerca"
         Height          =   495
         Left            =   3120
         TabIndex        =   2
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton cmdLejos 
         Caption         =   "Lejos"
         Height          =   495
         Left            =   3120
         TabIndex        =   1
         Top             =   1080
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmRadioPrimo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************
'* PROYECTO      : RADIO PRIMOS
'* CONTENIDO     : CALCULAR NÚMEROS PRIMOS, MOSTRARLOS POR NIVELES EN CIRCUNFERENCIAS
'* VERSION       : 1.1
'* AUTORES       : MIGUEL QUINTEIRO PIÑERO / MIGUEL QUINTEIRO FERNANDEZ
'* INICIO        : 16 DE MAYO DE 2017
'* ACTUALIZACION : 16 DE MAYO DE 2017
'****************************************************************************************
Option Explicit

' Declaración de variables
Dim miFactorCircular As Double
Dim miPi As Double
Dim X1 As Double
Dim Y1 As Double
Dim X2 As Double
Dim Y2 As Double

Dim X3 As Double
Dim Y3 As Double
Dim X5 As Double
Dim Y5 As Double


Dim miRadio As Long
Dim miN As Long
Dim miCuentaPrimos As Long
Dim miCuentaSuperior As Long
Dim miCuentaInferior As Long
Dim miMiniRadio As Long
Dim r As Integer
Dim miZoom As Long
Dim miEtiqueta As Boolean
Dim miCompuestos As Boolean
Dim miLineasP As Boolean
Dim miLineasC As Boolean
Dim miEjes As Boolean
Dim miUne As Boolean
Dim miOrbitaMaxima As Long

' Declaración de arreglos
Dim miOrbitaP() As Long
Dim miOrbitaPG() As Boolean

Dim miOrbitaC() As Long


' Declaración de variables
Dim CentroX As Double
Dim CentroY As Double
Dim CentroZ As Double



' AL CARGAR EL FORMULARIO
Private Sub Form_Load()
' Inicialización de variable
  miPi = 3.1415926535
  miRadio = 4000
  miFactorCircular = 1.15
  miN = 2
  miMiniRadio = 1
  miZoom = 8
  miOrbitaMaxima = 200
  miEtiqueta = False
  miCompuestos = False
  miLineasP = False
  miLineasC = False
  miEjes = False
  miUne = False

  ReDim miOrbitaP(miOrbitaMaxima)
  ReDim miOrbitaPG(miOrbitaMaxima)
  ReDim miOrbitaC(miOrbitaMaxima)

End Sub

' AL DARLE DOBLE CLICK
Private Sub Form_DblClick()
' Coordenadas del Centro
  CentroX = 0
  CentroY = frmEspacio3D.Width / 2
  CentroZ = frmEspacio3D.Height / 2

  ' Dibuja Circulo
  miN = InputBox("Ingrese el número N (Entre 1 y 30000)")

  Call Grafica
End Sub

' Muestra los ejes de coordenadas
Private Sub cmdEjes_Click()
  If miEjes = True Then
    miEjes = False
  Else
    miEjes = True
  End If
  Call Grafica
End Sub

' Muestra los números compuestos
Private Sub cmdCompuestos_Click()
  If miCompuestos = True Then
    miCompuestos = False
  Else
    miCompuestos = True
  End If
  Call Grafica
End Sub

' Muestra las etiquetas de los números
Private Sub cmdEtiqueta_Click()
  If miEtiqueta = True Then
    miEtiqueta = False
  Else
    miEtiqueta = True
  End If
  Call Grafica
End Sub

' Muestra la línea de los primos
Private Sub cmdLineas_Click()
  If miLineasP = True Then
    miLineasP = False
  Else
    miLineasP = True
  End If
  Call Grafica
End Sub

' Muestra las líneas de los compuestos
Private Sub cmdLineasCompuestos_Click()
  If miLineasC = True Then
    miLineasC = False
  Else
    miLineasC = True
  End If
  Call Grafica
End Sub

' Muestra las líneas de los primos gemelos
Private Sub cmdUne_Click()
  If miUne = True Then
    miUne = False
  Else
    miUne = True
  End If
  Call Grafica
End Sub

' Pinta los puntos en la circunferencia
Private Sub cmdGrafica_Click()
  miN = txtN
  Call Grafica
End Sub

' Aleja la imagen de los puntos
Private Sub cmdLejos_Click()
  miZoom = miZoom * 2
  Call Grafica
End Sub

' Acerca la imagen de los puntos
Private Sub cmdCerca_Click()
  miZoom = miZoom / 2
  Call Grafica
End Sub

' Reduce a la mitad la cantidad de puntos mostrados
Private Sub cmdMitad_Click()
  txtN.Text = Val(txtN.Text) / 2
  miN = txtN
  Call Grafica
End Sub

' Duplica la cantidad de puntos mostrados
Private Sub cmdDoble_Click()
  txtN.Text = Val(txtN.Text) * 2
  miN = txtN
  Call Grafica
End Sub

' Muestra información de la orbita
Private Sub cmdInformacionOrbita_Click()
  Dim i As Long
  If Val(txtNumeroOrbita.Text) <> 0 Then
    txtOrbitaPrimo.Text = miOrbitaP(Val(txtNumeroOrbita.Text))
    txtOrbitaCompuesto.Text = miOrbitaC(Val(txtNumeroOrbita.Text))
    ' Marca la orbita con líneas
    Call Grafica
    If ((miRadio * (Val(txtNumeroOrbita.Text) / miZoom) * miFactorCircular) - 30) > 0 Then
      Circle (4750, 4750), (miRadio * (Val(txtNumeroOrbita.Text) / miZoom) * miFactorCircular) - 30, vbWhite
      Circle (4750, 4750), (miRadio * (Val(txtNumeroOrbita.Text) / miZoom) * miFactorCircular) + 30, vbWhite
    End If

    txtAviso.Text = miOrbitaPG(Val(txtNumeroOrbita.Text))
  End If


End Sub

' Información sobre las órbitas
Private Sub cmdListaOrbitas_Click()
  Dim miMensaje As String
  Dim o As Integer
  miMensaje = ""
  miMensaje = miMensaje + "****   Listado Órbitas   ****" + vbCrLf + vbCrLf
  miMensaje = miMensaje + "   #      P       C" + vbCrLf
  For o = 0 To miOrbitaMaxima
    miMensaje = miMensaje + Tabulado(Trim(Str(o)), 7) + _
                Tabulado(Trim(Str(miOrbitaP(o))), 7) + _
                Tabulado(Trim(Str(miOrbitaPG(o))), 7) + _
                Tabulado(Trim(Str(miOrbitaC(o))), 7) + vbCrLf
    If miOrbitaP(o) = 0 And miOrbitaC(o) = 0 Then
      o = miOrbitaMaxima + 1
    End If
  Next o
  MsgBox miMensaje, , "Información Orbitas"
End Sub

' Busca primo anterior
Private Sub cmdPrimoAnterior_Click()
  Dim miDato As Long
  miDato = Val(txtN.Text)
  miN = txtN
  If miDato > 2 Then
    miDato = miDato - 1
    While Not Primo(miDato)
      miDato = miDato - 1
    Wend
    Cls
    txtN.Text = miDato
    Call Grafica
  End If
  DoEvents
End Sub

' Busca primo siguiente
Private Sub cmdPrimoSiguiente_Click()
  Dim miDato As Long
  miDato = Val(txtN.Text)
  miN = txtN
  If miDato > 2 Then
    miDato = miDato + 1
    While Not Primo(miDato)
      miDato = miDato + 1
    Wend
    Cls
    txtN.Text = miDato
    Call Grafica
  End If
  DoEvents
End Sub

' DIBUJA UN CIRCULO
Public Sub DibujaCirculo(ByVal pX As Long, ByVal pY As Long, ByVal pRadio As Long, ByVal pColor As Integer)
'Circle (pX, pY), pRadio, QBColor(pColor)
'Circle (pY, pX), pRadio, QBColor(pColor)
End Sub

' Dibuja los puntos de la imagen
Public Sub Grafica()
  If miN <= 900000 Then
    ' Borra la pantalla
    Cls
    ' Marco
    'Line (100, 100)-(9500, 9500), , B

    ' Ejes de Coordenadas
    Line (4750, 0)-(4750, 9500)
    Line (0, 4750)-(9500, 4750)
    Line (0, 0)-(9500, 9500)
    Line (0, 9500)-(9500, 0)

    ' Borra el área de la circunferencia
    Dim r As Long
    For r = 1 To miRadio * miFactorCircular
      Circle (4750, 4750), r, frmRadioPrimo.BackColor
    Next r

    ' Inicializa los contadores de orbitas
    ReDim miOrbitaP(miOrbitaMaxima)
    ReDim miOrbitaPG(miOrbitaMaxima)
    ReDim miOrbitaC(miOrbitaMaxima)

    ' Recorre toda las circunferencia
    Dim i As Long
    For i = 1 To miN
      ' Puntos iniciales
      If i = 1 Then
        miMiniRadio = 0.5
      End If
      If i = 2 Then
        miMiniRadio = 1
      End If
      If i = 3 Then
        miMiniRadio = 2
      End If

      '            ' Puntos colocados arbitrariamente
      '            ' Puntos iniciales
      '            If i = 1 Then
      '                miMiniRadio = 0.5
      '            End If
      '            If i = 2 Then
      '                miMiniRadio = 0.5
      '            End If
      '            If i = 3 Then
      '                miMiniRadio = 1
      '            End If

      ' Cálculo de las coordenadas X, Y
      X1 = 4750 + ((miRadio * (miMiniRadio / miZoom)) * Cos((360 / miN) * (miPi / 180) * i) * miFactorCircular)
      Y1 = 4750 + ((miRadio * (miMiniRadio / miZoom)) * -Sin((360 / miN) * (miPi / 180) * i) * miFactorCircular)

      ' Calcula las coordenadas en la esfera
      Dim mDistancia As Double
      Dim mEX As Double
      Dim mEY As Double
      Dim mEZ As Double

      '''            mDistancia = CalculaDistancia(X1, Y1)
      '''            mEX = CalculaEX(X1, mDistancia)
      '''            mEY = CalculaEY(Y1, mDistancia)
      '''            mEZ = CalculaEZ(mDistancia)

      ' Calcula si es primo
      If Primo(i) = True Then
        miCuentaPrimos = miCuentaPrimos + 1
        ' Calcula cantidad Superior e inferior
        If Y1 <= 4750 Then
          miCuentaSuperior = miCuentaSuperior + 1
        Else
          miCuentaInferior = miCuentaInferior + 1
        End If

        ' Muestra los primos
        If miLineasP = True Then
          ' Línea Prima
          'Line (X1, Y1)-(4750, 4750), vbYellow
          Line (Y1 - X1 + 4750, X1)-(4750, 4750), vbYellow
        End If

        ' Punto Primo
        For r = 0 To 10
          'Circle (X1, Y1), r, QBColor(12)
          Circle (Y1 - X1 + 4750, X1), r, QBColor(12)
        Next r

        If miEtiqueta = True Then
          frmRadioPrimo.ForeColor = vbRed
          Print i
        End If

        ' Une Puntos
        If i <> 2 Then
          If miUne = True Then

            If i = 3 Then
              X3 = X1
              Y3 = Y1
            End If
            If i = 5 Then
              X5 = X1
              Y5 = Y1
              'Line (X3, Y3)-(X5, Y5), vbGreen
              Line (Y3 - X3 + 4750, X3)-(Y5 - X5 + 4750, X5), vbGreen
            End If

            ' Almacena el actual
            If Primo(i + 2) Then
              X2 = X1
              Y2 = Y1
            End If
            ' Los une en el momento oportuno
            If (i - 2) > 0 Then
              If Primo(i - 2) Then
                'Line (X1, Y1)-(X2, Y2), vbGreen
                Line (Y1 - X1 + 4750, X1)-(Y2 - X2 + 4750, X2), vbGreen


              End If
            End If
          End If
        End If

        ' Aumenta un primo en la orbita
        miOrbitaP(miMiniRadio) = miOrbitaP(miMiniRadio) + 1
        If Primo(i + 2) Then
          miOrbitaPG(miMiniRadio) = True
        End If

        '**************************************************
        '**************************************************
        mEX = X1
        mEY = Y1
        'mEZ = ((X1 ^ 2) + (Y1 ^ 2)) / ((4750 * 0.01) ^ 2)
        mEZ = (miMiniRadio ^ 2) * 2

        ' calcula las coordenadas del Plot3d
        Call Coordenadas3D(mEX, mEY, mEZ)
        ' Grafica en 3d
        'Circle (miSX, miSY), 10, QBColor(13)

        '**************************************************
        Call PP3D(mEX, mEY, mEZ, 15, vbBlue)
        '**************************************************



        miMiniRadio = 1






      Else
        ' Muestra los compuestos
        If miCompuestos = True Then
          If miLineasC = True Then
            ' Línea Prima
            'Line (X1, Y1)-(4750, 4750), vbBlue
            Line (Y1 - X1 + 4750, X1)-(4750, 4750), vbBlue
          End If
          For r = 0 To 10
            'Circle (X1, Y1), r, QBColor(0)
            Circle (Y1 - X1 + 4750, X1), r, QBColor(0)
          Next r
          If miEtiqueta = True Then
            frmRadioPrimo.ForeColor = vbBlack
            Print i
          End If
        End If
        ' Aumenta un compuesto en la orbita
        miOrbitaC(miMiniRadio) = miOrbitaC(miMiniRadio) + 1

        miMiniRadio = miMiniRadio + 1

        ' Controla a la orbita maxima
        If miMiniRadio > miOrbitaMaxima Then
          miOrbitaMaxima = miOrbitaMaxima + 1
          ReDim Preserve miOrbitaP(miOrbitaMaxima)
          ReDim Preserve miOrbitaC(miOrbitaMaxima)
        End If
      End If
    Next i

    'Restablece color de fuentes y dibuja los ejes
    frmRadioPrimo.ForeColor = vbBlack
    If miEjes = True Then
      ' Ejes de Coordenadas
      Line (4750, 0)-(4750, 9500)
      Line (0, 4750)-(9500, 4750)
      'Line (0, 0)-(9500, 9500)
      'Line (0, 9500)-(9500, 0)
    End If

    ' Indica los Primos Gemelos
    If (miOrbitaP(2) - 1) > 0 Then
      txtPrimosGemelos.Text = miOrbitaP(2) - 1
    End If
  End If
End Sub

Private Sub cmdEstadisticaOrbita_Click()
' Abre archivo para escritura
  Open "EstadisticaOrbita.txt" For Output As #1

  Dim i As Integer
  Dim acumP As Long
  Dim acumC As Long

  Dim relacion As Double
  Dim PorcentajePrimo As Double
  Dim PorcentajeCompuesto As Double

  Print #1, "Estadística para .--- "; miN
  Print #1, ""
  Print #1, "    O.     P.     C.                 %P.     %C.     C/P."
  Print #1, "---------------------------------------------------------"

  acumP = 0
  acumC = 0
  For i = 0 To miOrbitaMaxima
    acumP = acumP + miOrbitaP(i)
    acumC = acumC + miOrbitaC(i)

    If (miOrbitaP(i) + miOrbitaC(i)) > 0 Then
      PorcentajePrimo = (miOrbitaP(i) * 100) / (miOrbitaP(i) + miOrbitaC(i))
      PorcentajeCompuesto = (miOrbitaC(i) * 100) / (miOrbitaP(i) + miOrbitaC(i))
    Else
      PorcentajePrimo = 0
      PorcentajeCompuesto = 0
    End If

    Print #1, Tabulado(Trim(Str(i)), 5); "  "; _
              Tabulado(Trim(Str(miOrbitaP(i))), 5); "  "; _
              Tabulado(Trim(Str(miOrbitaC(i))), 5); "  "; _
              "   ***    "; "  "; _
              Tabulado(Trim(Format(PorcentajePrimo, "##,##0.00")), 6); "  "; _
              Tabulado(Trim(Format(PorcentajeCompuesto, "##,##0.00")), 6); "  ";
    If miOrbitaP(i) > 0 Then
      relacion = miOrbitaC(i) / miOrbitaP(i)
      Print #1, Tabulado(Trim(Format(relacion, "##,##0.00")), 8); "   ";
      Print #1, Tabulado(Trim(Str(miOrbitaPG(i))), 5)
    Else
      Print #1, "--------"; "   ";
      Print #1, Tabulado(Trim(Str(miOrbitaPG(i))), 5)
    End If

    If (miOrbitaP(i) = 0) And (miOrbitaC(i) = 0) Then
      Print #1, "---------------------------------------------------------"
      Print #1, "T.  "; Tabulado(Trim(Str(acumP)), 8); " "; Tabulado(Trim(Str(acumC)), 8); "            ";

      PorcentajePrimo = (acumP * 100) / miN
      PorcentajeCompuesto = (acumC * 100) / miN
      Print #1, Tabulado(Trim(Format(PorcentajePrimo, "##,##0.00")), 6); "  "; _
                Tabulado(Trim(Format(PorcentajeCompuesto, "##,##0.00")), 6); "  ";
      relacion = acumC / acumP
      Print #1, Tabulado(Trim(Format(relacion, "##,##0.00")), 8)

      i = miOrbitaMaxima + 1
    End If

  Next i

  ' Cierra archivo
  Close #1

End Sub


' FUNCION PARA CALCULAR SI EL NUMERO ES PRIMO
Public Function Primo(ByVal pN As Long) As Boolean
  Dim i As Long
  Primo = True
  If pN = 1 Then
    Primo = False
  Else
    For i = 2 To Sqr(pN)
      If (pN / i) = Int(pN / i) Then
        Primo = False
      End If
    Next i
  End If
End Function

' FUNCION PARA CALCULAR SI EL NUMERO ES PRIMO
Public Function Tabulado(ByVal pT As String, ByVal pA As Integer) As String
  Dim i As Integer
  Dim miAncho As Integer
  miAncho = Len(Trim(pT))

  For i = 1 To (pA - miAncho)
    'pT = pT + " "
    pT = " " + pT
  Next i
  Tabulado = pT
End Function

Public Sub Coordenadas3D(ByVal pX As Double, ByVal pY As Double, ByVal pZ As Double)
' Cálculo de las coordenadas X, Y
  miSX = (4750 / 2) + ((-(pX * Cos(-45 * 3.1415 / 180))) + pY)
  miSY = (4750 / 2) + ((-(pX * Sin(-90 * 3.1415 / 180))) + (-pZ))
End Sub

' Grafica en espacio vectorial de tres dimensiones 3D
Public Sub PP3D(ByVal pX As Double, ByVal pY As Double, ByVal pZ As Double, pTamaño As Integer, ByVal pColor As Long)
' Declaración de variables
  Dim ppX As Double
  Dim ppY As Double
  Dim Ang As Double
  Dim Radio As Double
  ' Ajuste del Angulo
  Ang = 15 * (3.141592 / 180)
  ' Coordenadas de Pantalla del Punto
  ppX = CentroY + (-pX * Cos(Ang)) + (pY) + (0) + 5000
  ppY = CentroZ + (pX * Sin(Ang)) + (0) + (-pZ) + 8000
  ' Mostrar el punto con control del Tamaño y Color
  If pTamaño <= 0 Then
    PSet (ppX, ppY), pColor
  Else
    For Radio = 1 To pTamaño
      Circle (ppX, ppY), Radio, pColor
    Next
  End If
End Sub

