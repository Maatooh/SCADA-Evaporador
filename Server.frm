VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Server 
   BackColor       =   &H00FCFCFC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maltexco - Evaporador"
   ClientHeight    =   11085
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   20400
   Icon            =   "Server.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   739
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1360
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BAlarm 
      BackColor       =   &H008080FF&
      Caption         =   "OFF"
      Height          =   255
      Left            =   15360
      Style           =   1  'Graphical
      TabIndex        =   218
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton EvAM 
      BackColor       =   &H008080FF&
      Caption         =   "Man"
      Height          =   255
      Left            =   15960
      Style           =   1  'Graphical
      TabIndex        =   216
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton EvAgua 
      BackColor       =   &H008080FF&
      Caption         =   "OFF"
      Height          =   255
      Left            =   15360
      Style           =   1  'Graphical
      TabIndex        =   215
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton RefreshDev 
      BackColor       =   &H008080FF&
      Caption         =   "Refresh"
      Height          =   375
      Left            =   18240
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   213
      Top             =   0
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock WServer 
      Index           =   0
      Left            =   3000
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LED G"
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   212
      Top             =   10920
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LED R"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   211
      Top             =   10920
      Width           =   735
   End
   Begin VB.Frame DriveFrame 
      BackColor       =   &H00FCFCFC&
      Caption         =   "Controles"
      Height          =   735
      Left            =   14280
      TabIndex        =   48
      Top             =   960
      Width           =   6015
      Begin VB.CommandButton CommandTab 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Caption         =   "Circuito Evaporación"
         Height          =   375
         Index           =   0
         Left            =   120
         MaskColor       =   &H0080FFFF&
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton CommandTab 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Caption         =   "Evaporador Quintuple"
         Height          =   375
         Index           =   1
         Left            =   1800
         MaskColor       =   &H0080FFFF&
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton CommandTab 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Caption         =   "Evaporador PU."
         Height          =   375
         Index           =   2
         Left            =   3480
         MaskColor       =   &H0080FFFF&
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CommandTab 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Caption         =   "Circuitos CIP"
         Height          =   375
         Index           =   3
         Left            =   4800
         MaskColor       =   &H0080FFFF&
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton SDevices 
      BackColor       =   &H0080C0FF&
      Caption         =   "Devices"
      Height          =   375
      Left            =   19320
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   1095
   End
   Begin VB.Frame FrameWork 
      BackColor       =   &H00FCFCFC&
      Caption         =   "EVAPORADOR CONTRERAS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9135
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   20175
      Begin VB.Frame SensorFrame 
         BackColor       =   &H00FCFCFC&
         Caption         =   "Sensores"
         Height          =   4095
         Index           =   1
         Left            =   14160
         TabIndex        =   183
         Top             =   5040
         Visible         =   0   'False
         Width           =   5895
         Begin VB.CommandButton SFrameNB 
            Caption         =   "<"
            Height          =   1215
            Index           =   1
            Left            =   120
            TabIndex        =   184
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label LSENSOR 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "PIC203 PRES. TERMO 1°EF"
            Height          =   495
            Index           =   15
            Left            =   600
            TabIndex        =   208
            Top             =   3165
            Width           =   1335
         End
         Begin VB.Label Sensor 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   15
            Left            =   720
            TabIndex        =   207
            Top             =   3660
            Width           =   1095
         End
         Begin VB.Label Sensor 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   19
            Left            =   2400
            TabIndex        =   206
            Top             =   3660
            Width           =   1095
         End
         Begin VB.Label Sensor 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   23
            Left            =   4080
            TabIndex        =   205
            Top             =   3660
            Width           =   1095
         End
         Begin VB.Label LSENSOR 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "FIC202 CAUDAL DE SALIDA 3°EF Q"
            Height          =   615
            Index           =   19
            Left            =   2160
            TabIndex        =   204
            Top             =   3165
            Width           =   1575
         End
         Begin VB.Label LSENSOR 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "FIC204 CAUDAL DE SALIDA PU"
            Height          =   615
            Index           =   23
            Left            =   3840
            TabIndex        =   203
            Top             =   3165
            Width           =   1575
         End
         Begin VB.Label LSENSOR 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TIA211 T° AGUA DESDE PU"
            Height          =   615
            Index           =   22
            Left            =   3960
            TabIndex        =   202
            Top             =   2235
            Width           =   1335
         End
         Begin VB.Label LSENSOR 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TIA203 T° AGUA DE TORRE "
            Height          =   615
            Index           =   18
            Left            =   2280
            TabIndex        =   201
            Top             =   2235
            Width           =   1335
         End
         Begin VB.Label Sensor 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   22
            Left            =   4080
            TabIndex        =   200
            Top             =   2715
            Width           =   1095
         End
         Begin VB.Label Sensor 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   18
            Left            =   2400
            TabIndex        =   199
            Top             =   2715
            Width           =   1095
         End
         Begin VB.Label Sensor 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   14
            Left            =   720
            TabIndex        =   198
            Top             =   2715
            Width           =   1095
         End
         Begin VB.Label LSENSOR 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TIA210 TEMP. VAHOS 5° EF Q"
            Height          =   615
            Index           =   14
            Left            =   600
            TabIndex        =   197
            Top             =   2235
            Width           =   1335
         End
         Begin VB.Label LSENSOR 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TIA208 TEMP. VAHOS 3° EF Q"
            Height          =   615
            Index           =   13
            Left            =   600
            TabIndex        =   196
            Top             =   1275
            Width           =   1335
         End
         Begin VB.Label Sensor 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   13
            Left            =   720
            TabIndex        =   195
            Top             =   1755
            Width           =   1095
         End
         Begin VB.Label Sensor 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   17
            Left            =   2400
            TabIndex        =   194
            Top             =   1755
            Width           =   1095
         End
         Begin VB.Label Sensor 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   21
            Left            =   4080
            TabIndex        =   193
            Top             =   1755
            Width           =   1095
         End
         Begin VB.Label LSENSOR 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "PIC205 PRES. CALANDRIA PU"
            Height          =   615
            Index           =   17
            Left            =   2280
            TabIndex        =   192
            Top             =   1275
            Width           =   1335
         End
         Begin VB.Label LSENSOR 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TIA204 T° AGUA DESDE Q"
            Height          =   615
            Index           =   21
            Left            =   3960
            TabIndex        =   191
            Top             =   1275
            Width           =   1335
         End
         Begin VB.Label LSENSOR 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "T° MALTOSA"
            Height          =   375
            Index           =   20
            Left            =   3960
            TabIndex        =   190
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label LSENSOR 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "PIA206 VACIO EVAPORADOR PU"
            Height          =   615
            Index           =   16
            Left            =   2280
            TabIndex        =   189
            Top             =   195
            Width           =   1335
         End
         Begin VB.Label Sensor 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   20
            Left            =   4080
            TabIndex        =   188
            Top             =   795
            Width           =   1095
         End
         Begin VB.Label Sensor 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   16
            Left            =   2400
            TabIndex        =   187
            Top             =   795
            Width           =   1095
         End
         Begin VB.Label Sensor 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   12
            Left            =   720
            TabIndex        =   186
            Top             =   795
            Width           =   1095
         End
         Begin VB.Label LSENSOR 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TIA205 TEMP.CALANDRIA 1°EF Q"
            Height          =   615
            Index           =   12
            Left            =   480
            TabIndex        =   185
            Top             =   195
            Width           =   1575
         End
      End
      Begin VB.Frame SensorFrame 
         BackColor       =   &H00FCFCFC&
         Caption         =   "Sensores"
         Height          =   4095
         Index           =   0
         Left            =   14160
         TabIndex        =   157
         Top             =   5040
         Width           =   5895
         Begin VB.CommandButton SFrameNB 
            Caption         =   ">"
            Height          =   1215
            Index           =   0
            Left            =   5280
            TabIndex        =   164
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label LSENSOR 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "----------------"
            Height          =   375
            Index           =   3
            Left            =   600
            TabIndex        =   182
            Top             =   3165
            Width           =   1335
         End
         Begin VB.Label Sensor 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   720
            TabIndex        =   181
            Top             =   3660
            Width           =   1095
         End
         Begin VB.Label Sensor 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   7
            Left            =   2400
            TabIndex        =   180
            Top             =   3660
            Width           =   1095
         End
         Begin VB.Label Sensor 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   11
            Left            =   4080
            TabIndex        =   179
            Top             =   3660
            Width           =   1095
         End
         Begin VB.Label LSENSOR 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "FIC201 CAUDAL ALIM."
            Height          =   615
            Index           =   7
            Left            =   2280
            TabIndex        =   178
            Top             =   3160
            Width           =   1335
         End
         Begin VB.Label LSENSOR 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TEMPERATURA PURATO CONTROL"
            Height          =   615
            Index           =   11
            Left            =   3840
            TabIndex        =   177
            Top             =   3165
            Width           =   1575
         End
         Begin VB.Label LSENSOR 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TIC213 T° VAHOS PU"
            Height          =   615
            Index           =   10
            Left            =   4080
            TabIndex        =   176
            Top             =   2235
            Width           =   1095
         End
         Begin VB.Label LSENSOR 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TIA209 TEMP. VAHOS 4° EF Q"
            Height          =   615
            Index           =   6
            Left            =   2280
            TabIndex        =   175
            Top             =   2235
            Width           =   1335
         End
         Begin VB.Label Sensor 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   10
            Left            =   4080
            TabIndex        =   174
            Top             =   2715
            Width           =   1095
         End
         Begin VB.Label Sensor 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   2400
            TabIndex        =   173
            Top             =   2715
            Width           =   1095
         End
         Begin VB.Label Sensor 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   720
            TabIndex        =   172
            Top             =   2715
            Width           =   1095
         End
         Begin VB.Label LSENSOR 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TIA202 TEMP. PRODUCTO TQ2"
            Height          =   615
            Index           =   2
            Left            =   600
            TabIndex        =   171
            Top             =   2235
            Width           =   1335
         End
         Begin VB.Label LSENSOR 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TIA201 TEMP. PRODUCTO TQ1"
            Height          =   375
            Index           =   1
            Left            =   600
            TabIndex        =   170
            Top             =   1275
            Width           =   1335
         End
         Begin VB.Label Sensor 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   720
            TabIndex        =   169
            Top             =   1755
            Width           =   1095
         End
         Begin VB.Label Sensor 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   2400
            TabIndex        =   168
            Top             =   1755
            Width           =   1095
         End
         Begin VB.Label Sensor 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   9
            Left            =   4080
            TabIndex        =   167
            Top             =   1755
            Width           =   1095
         End
         Begin VB.Label LSENSOR 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TIA206 TEMP. VAHOS 1° EF Q"
            Height          =   615
            Index           =   5
            Left            =   2280
            TabIndex        =   166
            Top             =   1275
            Width           =   1335
         End
         Begin VB.Label LSENSOR 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TIA207 TEMP. VAHOS 2° EF Q"
            Height          =   615
            Index           =   9
            Left            =   3960
            TabIndex        =   165
            Top             =   1275
            Width           =   1335
         End
         Begin VB.Label LSENSOR 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "T° Y CONTROL CALANDRIA "
            Height          =   615
            Index           =   8
            Left            =   3960
            TabIndex        =   163
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label LSENSOR 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "PIA202 VACIO EVAPORADOR Q"
            Height          =   615
            Index           =   4
            Left            =   2280
            TabIndex        =   162
            Top             =   195
            Width           =   1335
         End
         Begin VB.Label Sensor 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   8
            Left            =   4080
            TabIndex        =   161
            Top             =   795
            Width           =   1095
         End
         Begin VB.Label Sensor 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   2400
            TabIndex        =   160
            Top             =   795
            Width           =   1095
         End
         Begin VB.Label Sensor 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   720
            TabIndex        =   159
            Top             =   795
            Width           =   1095
         End
         Begin VB.Label LSENSOR 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "PIA201 PRES. GENERAL DE VAPOR"
            Height          =   615
            Index           =   0
            Left            =   600
            TabIndex        =   158
            Top             =   195
            Width           =   1335
         End
      End
      Begin VB.Frame ControlFrame 
         BackColor       =   &H00FCFCFC&
         Caption         =   "Control Manual"
         Height          =   5655
         Index           =   3
         Left            =   14160
         TabIndex        =   117
         Top             =   150
         Visible         =   0   'False
         Width           =   5895
         Begin VB.CommandButton Switch3 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   16
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   136
            Top             =   480
            Width           =   615
         End
         Begin VB.CommandButton Switch3 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   17
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   135
            Top             =   480
            Width           =   615
         End
         Begin VB.CommandButton Switch3 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   18
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   134
            Top             =   480
            Width           =   615
         End
         Begin VB.CommandButton Switch3 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   3
            Left            =   4800
            Style           =   1  'Graphical
            TabIndex        =   133
            Top             =   1380
            Width           =   615
         End
         Begin VB.CommandButton Switch3 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   2
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   132
            Top             =   1380
            Width           =   615
         End
         Begin VB.CommandButton Switch3 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   1
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   131
            Top             =   1380
            Width           =   615
         End
         Begin VB.CommandButton Switch3 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   0
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   130
            Top             =   1380
            Width           =   615
         End
         Begin VB.CommandButton Switch3 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   4
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   129
            Top             =   2340
            Width           =   615
         End
         Begin VB.CommandButton Switch3 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   5
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   128
            Top             =   2340
            Width           =   615
         End
         Begin VB.CommandButton Switch3 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   6
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   127
            Top             =   2340
            Width           =   615
         End
         Begin VB.CommandButton Switch3 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   7
            Left            =   4800
            Style           =   1  'Graphical
            TabIndex        =   126
            Top             =   2340
            Width           =   615
         End
         Begin VB.CommandButton Switch3 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   11
            Left            =   4800
            Style           =   1  'Graphical
            TabIndex        =   125
            Top             =   3240
            Width           =   615
         End
         Begin VB.CommandButton Switch3 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   10
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   124
            Top             =   3240
            Width           =   615
         End
         Begin VB.CommandButton Switch3 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   9
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   123
            Top             =   3240
            Width           =   615
         End
         Begin VB.CommandButton Switch3 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   8
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   122
            Top             =   3240
            Width           =   615
         End
         Begin VB.CommandButton Switch3 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   15
            Left            =   4800
            Style           =   1  'Graphical
            TabIndex        =   121
            Top             =   4200
            Width           =   615
         End
         Begin VB.CommandButton Switch3 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   14
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   120
            Top             =   4200
            Width           =   615
         End
         Begin VB.CommandButton Switch3 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   13
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   119
            Top             =   4200
            Width           =   615
         End
         Begin VB.CommandButton Switch3 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   12
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   118
            Top             =   4200
            Width           =   615
         End
         Begin VB.Label ControlTitle 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "SECTOR VALVULAS DE CIRCUITOS CIP"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   156
            Top             =   5160
            Width           =   5895
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "PRUEBA LAMPARAS"
            Height          =   375
            Index           =   69
            Left            =   240
            TabIndex        =   155
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TERMO COMPRESOR"
            Height          =   375
            Index           =   68
            Left            =   1680
            TabIndex        =   154
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "SILENCIO DE ALARMA"
            Height          =   375
            Index           =   67
            Left            =   3000
            TabIndex        =   153
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VMC220 VALV. HABILITA CIP"
            Height          =   375
            Index           =   65
            Left            =   4440
            TabIndex        =   152
            Top             =   1740
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VMC219 VALV. REGULA CIP"
            Height          =   375
            Index           =   64
            Left            =   3000
            TabIndex        =   151
            Top             =   1740
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VMC218 VALV. CIP TQ2"
            Height          =   375
            Index           =   63
            Left            =   1560
            TabIndex        =   150
            Top             =   1740
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VMC216 VALV. CIP TQ1"
            Height          =   375
            Index           =   62
            Left            =   120
            TabIndex        =   149
            Top             =   1740
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VMC221 VALV. CIP SUP. 1°EF Q"
            Height          =   375
            Index           =   61
            Left            =   120
            TabIndex        =   148
            Top             =   2700
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VMC222 VALV. CIP INF. 1°EF Q"
            Height          =   375
            Index           =   60
            Left            =   1560
            TabIndex        =   147
            Top             =   2700
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VMC223 VALV. CIP SUP. 2°EF Q"
            Height          =   375
            Index           =   58
            Left            =   3000
            TabIndex        =   146
            Top             =   2700
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VMC224 VALV. CIP INF. 2°EF Q"
            Height          =   375
            Index           =   56
            Left            =   4440
            TabIndex        =   145
            Top             =   2700
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VMC228 VALV. CIP INF. 4°EF Q"
            Height          =   375
            Index           =   55
            Left            =   4440
            TabIndex        =   144
            Top             =   3600
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VMC227 VALV. CIP SUP. 4°EF Q"
            Height          =   375
            Index           =   53
            Left            =   3000
            TabIndex        =   143
            Top             =   3600
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VMC226 VALV. CIP INF. 3°EF Q"
            Height          =   375
            Index           =   50
            Left            =   1560
            TabIndex        =   142
            Top             =   3600
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VMC225 VALV. CIP SUP. 3°EF Q"
            Height          =   375
            Index           =   48
            Left            =   120
            TabIndex        =   141
            Top             =   3600
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VMC232 VALV. CIP INF. PU"
            Height          =   375
            Index           =   47
            Left            =   4440
            TabIndex        =   140
            Top             =   4560
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VMC231 VALV. CIP SUP. PU"
            Height          =   375
            Index           =   45
            Left            =   3000
            TabIndex        =   139
            Top             =   4560
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VMC230 VALV. CIP INF. 5°EF Q"
            Height          =   375
            Index           =   43
            Left            =   1560
            TabIndex        =   138
            Top             =   4560
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VMC229 VALV. CIP SUP. 5°EF Q"
            Height          =   375
            Index           =   41
            Left            =   120
            TabIndex        =   137
            Top             =   4560
            Width           =   1335
         End
      End
      Begin VB.Frame ControlFrame 
         BackColor       =   &H00FCFCFC&
         Caption         =   "Control Manual"
         Height          =   5655
         Index           =   2
         Left            =   14160
         TabIndex        =   95
         Top             =   150
         Visible         =   0   'False
         Width           =   5895
         Begin VB.CommandButton Switch2 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   0
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   105
            Top             =   480
            Width           =   615
         End
         Begin VB.CommandButton Switch2 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   1
            Left            =   4080
            Style           =   1  'Graphical
            TabIndex        =   104
            Top             =   480
            Width           =   615
         End
         Begin VB.CommandButton Switch2 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   3
            Left            =   4080
            Style           =   1  'Graphical
            TabIndex        =   103
            Top             =   1380
            Width           =   615
         End
         Begin VB.CommandButton Switch2 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   2
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   102
            Top             =   1380
            Width           =   615
         End
         Begin VB.CommandButton Switch2 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   4
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   101
            Top             =   2340
            Width           =   615
         End
         Begin VB.CommandButton Switch2 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   5
            Left            =   4080
            Style           =   1  'Graphical
            TabIndex        =   100
            Top             =   2340
            Width           =   615
         End
         Begin VB.CommandButton Switch2 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   7
            Left            =   4080
            Style           =   1  'Graphical
            TabIndex        =   99
            Top             =   3240
            Width           =   615
         End
         Begin VB.CommandButton Switch2 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   6
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   98
            Top             =   3240
            Width           =   615
         End
         Begin VB.CommandButton Switch2 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   9
            Left            =   4080
            Style           =   1  'Graphical
            TabIndex        =   97
            Top             =   4200
            Width           =   615
         End
         Begin VB.CommandButton Switch2 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   8
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   96
            Top             =   4200
            Width           =   615
         End
         Begin VB.Label ControlTitle 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "SECTOR EVAPORADOR PU."
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   116
            Top             =   5160
            Width           =   5895
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VM209 VALV. TORRE A PU"
            Height          =   375
            Index           =   59
            Left            =   960
            TabIndex        =   115
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VM210 VALV. PU A TORRE"
            Height          =   375
            Index           =   57
            Left            =   3720
            TabIndex        =   114
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "EV205 VALV. EYECTORES PU"
            Height          =   375
            Index           =   54
            Left            =   3720
            TabIndex        =   113
            Top             =   1740
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "BBA203 BOMBA TORRE A PU"
            Height          =   375
            Index           =   52
            Left            =   840
            TabIndex        =   112
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "EV204 VALV. VAPOR PU"
            Height          =   375
            Index           =   51
            Left            =   720
            TabIndex        =   111
            Top             =   2700
            Width           =   1575
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "BBA214 BOMBA EXTRACCION PU"
            Height          =   375
            Index           =   49
            Left            =   3720
            TabIndex        =   110
            Top             =   2700
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VMC207 VALV. BYPASS CIP IN"
            Height          =   375
            Index           =   46
            Left            =   3720
            TabIndex        =   109
            Top             =   3600
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "BBA216 BOMBA CONDENSADO 3"
            Height          =   375
            Index           =   44
            Left            =   840
            TabIndex        =   108
            Top             =   3600
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "BBA215 BOMBA BYPASS EN CIP"
            Height          =   375
            Index           =   42
            Left            =   3720
            TabIndex        =   107
            Top             =   4560
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VMC208 VALV. BYPASS CIP OUT"
            Height          =   375
            Index           =   40
            Left            =   840
            TabIndex        =   106
            Top             =   4560
            Width           =   1335
         End
      End
      Begin VB.Frame ControlFrame 
         BackColor       =   &H00FCFCFC&
         Caption         =   "Control Manual"
         Height          =   5655
         Index           =   1
         Left            =   14160
         TabIndex        =   53
         Top             =   150
         Visible         =   0   'False
         Width           =   5895
         Begin VB.CommandButton Switch1 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   16
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   73
            Top             =   4200
            Width           =   615
         End
         Begin VB.CommandButton Switch1 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   17
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   72
            Top             =   4200
            Width           =   615
         End
         Begin VB.CommandButton Switch1 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   18
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   71
            Top             =   4200
            Width           =   615
         End
         Begin VB.CommandButton Switch1 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   19
            Left            =   4800
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   4200
            Width           =   615
         End
         Begin VB.CommandButton Switch1 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   12
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   69
            Top             =   3240
            Width           =   615
         End
         Begin VB.CommandButton Switch1 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   13
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   68
            Top             =   3240
            Width           =   615
         End
         Begin VB.CommandButton Switch1 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   14
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   67
            Top             =   3240
            Width           =   615
         End
         Begin VB.CommandButton Switch1 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   15
            Left            =   4800
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   3240
            Width           =   615
         End
         Begin VB.CommandButton Switch1 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   11
            Left            =   4800
            Style           =   1  'Graphical
            TabIndex        =   65
            Top             =   2340
            Width           =   615
         End
         Begin VB.CommandButton Switch1 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   10
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   2340
            Width           =   615
         End
         Begin VB.CommandButton Switch1 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   9
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   63
            Top             =   2340
            Width           =   615
         End
         Begin VB.CommandButton Switch1 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   8
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   2340
            Width           =   615
         End
         Begin VB.CommandButton Switch1 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   4
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   1380
            Width           =   615
         End
         Begin VB.CommandButton Switch1 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   5
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   1380
            Width           =   615
         End
         Begin VB.CommandButton Switch1 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   6
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   1380
            Width           =   615
         End
         Begin VB.CommandButton Switch1 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   7
            Left            =   4800
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   1380
            Width           =   615
         End
         Begin VB.CommandButton Switch1 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   3
            Left            =   4800
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   480
            Width           =   615
         End
         Begin VB.CommandButton Switch1 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   2
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   480
            Width           =   615
         End
         Begin VB.CommandButton Switch1 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   1
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   480
            Width           =   615
         End
         Begin VB.CommandButton Switch1 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   0
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   480
            Width           =   615
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "BBA210 BOMBA RECIRCULADO 1"
            Height          =   375
            Index           =   20
            Left            =   120
            TabIndex        =   94
            Top             =   4560
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "BBA211 BOMBA TRASVASE 3"
            Height          =   375
            Index           =   21
            Left            =   1560
            TabIndex        =   93
            Top             =   4560
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "BBA212 BOMBA RECIRCULADO 2"
            Height          =   375
            Index           =   22
            Left            =   3000
            TabIndex        =   92
            Top             =   4560
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "BBA213 BOMBA EXTRACCION 2"
            Height          =   375
            Index           =   23
            Left            =   4440
            TabIndex        =   91
            Top             =   4560
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "BBA206 BOMBA TRASVASE 1"
            Height          =   375
            Index           =   24
            Left            =   120
            TabIndex        =   90
            Top             =   3600
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "BBA207 BOMBA TRASVASE 2"
            Height          =   375
            Index           =   25
            Left            =   1560
            TabIndex        =   89
            Top             =   3600
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "BBA208 BOMBA EXTRACCION 1"
            Height          =   375
            Index           =   26
            Left            =   3000
            TabIndex        =   88
            Top             =   3600
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "BBA209 BOMBA CONDENSADO 2"
            Height          =   375
            Index           =   27
            Left            =   4440
            TabIndex        =   87
            Top             =   3600
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "BBA205 BOMBA CONDENSADO 1"
            Height          =   375
            Index           =   28
            Left            =   4440
            TabIndex        =   86
            Top             =   2700
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "BBA204 BOMBA ALIMENTACION"
            Height          =   375
            Index           =   29
            Left            =   3000
            TabIndex        =   85
            Top             =   2700
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "EV203 VALV. VAPOR DOBLE"
            Height          =   375
            Index           =   30
            Left            =   1560
            TabIndex        =   84
            Top             =   2700
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "EV202 VALV VAPOR TERMO"
            Height          =   375
            Index           =   31
            Left            =   120
            TabIndex        =   83
            Top             =   2700
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VEN201 VENT. TORRE"
            Height          =   375
            Index           =   32
            Left            =   240
            TabIndex        =   82
            Top             =   1740
            Width           =   1095
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VM204 VALV. TORRE A Q"
            Height          =   375
            Index           =   33
            Left            =   1560
            TabIndex        =   81
            Top             =   1740
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VM205 VALV. Q A TORRE"
            Height          =   375
            Index           =   34
            Left            =   3000
            TabIndex        =   80
            Top             =   1740
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "BBA202 BOMBA TORRE A Q"
            Height          =   375
            Index           =   35
            Left            =   4440
            TabIndex        =   79
            Top             =   1740
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "EV201 VALV. EYECTORES Q"
            Height          =   375
            Index           =   36
            Left            =   4440
            TabIndex        =   78
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "BBA235 BOMBA VACIADO TQS"
            Height          =   375
            Index           =   37
            Left            =   3000
            TabIndex        =   77
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VM217 VALV. AGUA A TQ2"
            Height          =   375
            Index           =   38
            Left            =   1680
            TabIndex        =   76
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VM215 VALV. AGUA A TQ1"
            Height          =   375
            Index           =   39
            Left            =   240
            TabIndex        =   75
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label ControlTitle 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "SECTOR EVAPORADOR QUINTUPLE"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   74
            Top             =   5160
            Width           =   5895
         End
      End
      Begin VB.Frame ControlFrame 
         BackColor       =   &H00FCFCFC&
         Caption         =   "Control Manual"
         Height          =   5655
         Index           =   0
         Left            =   14160
         TabIndex        =   4
         Top             =   150
         Visible         =   0   'False
         Width           =   5895
         Begin VB.CommandButton Switch0 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   16
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   4200
            Width           =   615
         End
         Begin VB.CommandButton Switch0 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   17
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   4200
            Width           =   615
         End
         Begin VB.CommandButton Switch0 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   18
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   4200
            Width           =   615
         End
         Begin VB.CommandButton Switch0 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   19
            Left            =   4800
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   4200
            Width           =   615
         End
         Begin VB.CommandButton Switch0 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   12
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   3240
            Width           =   615
         End
         Begin VB.CommandButton Switch0 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   13
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   3240
            Width           =   615
         End
         Begin VB.CommandButton Switch0 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   14
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   3240
            Width           =   615
         End
         Begin VB.CommandButton Switch0 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   15
            Left            =   4800
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   3240
            Width           =   615
         End
         Begin VB.CommandButton Switch0 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   11
            Left            =   4800
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   2340
            Width           =   615
         End
         Begin VB.CommandButton Switch0 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   10
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   2340
            Width           =   615
         End
         Begin VB.CommandButton Switch0 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   9
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   2340
            Width           =   615
         End
         Begin VB.CommandButton Switch0 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   8
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   2340
            Width           =   615
         End
         Begin VB.CommandButton Switch0 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   4
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   1380
            Width           =   615
         End
         Begin VB.CommandButton Switch0 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   5
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   1380
            Width           =   615
         End
         Begin VB.CommandButton Switch0 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   6
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   1380
            Width           =   615
         End
         Begin VB.CommandButton Switch0 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   7
            Left            =   4800
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   1380
            Width           =   615
         End
         Begin VB.CommandButton Switch0 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   3
            Left            =   4800
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   480
            Width           =   615
         End
         Begin VB.CommandButton Switch0 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   2
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   480
            Width           =   615
         End
         Begin VB.CommandButton Switch0 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   1
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   480
            Width           =   615
         End
         Begin VB.CommandButton Switch0 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            Height          =   255
            Index           =   0
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   480
            Width           =   615
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "V3V204 VALV. DRENAJE"
            Height          =   375
            Index           =   19
            Left            =   240
            TabIndex        =   47
            Top             =   4560
            Width           =   1095
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "V3V205 VALV. RETORNO CIP"
            Height          =   375
            Index           =   18
            Left            =   1680
            TabIndex        =   46
            Top             =   4560
            Width           =   1095
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "V3V206 VALV. RECUPERACION"
            Height          =   375
            Index           =   17
            Left            =   3000
            TabIndex        =   45
            Top             =   4560
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "V3V207 VALV. RECIRCULACION"
            Height          =   375
            Index           =   16
            Left            =   4440
            TabIndex        =   44
            Top             =   4560
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VM211 VALV. ALIM. TQ 80%"
            Height          =   375
            Index           =   15
            Left            =   240
            TabIndex        =   39
            Top             =   3600
            Width           =   1095
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VM212 VALV. RETORNO PU"
            Height          =   375
            Index           =   14
            Left            =   1680
            TabIndex        =   38
            Top             =   3600
            Width           =   1095
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VM213 VALV. RETORNO COND"
            Height          =   375
            Index           =   13
            Left            =   3000
            TabIndex        =   37
            Top             =   3600
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VM214 VALV. BLOQUEO COND"
            Height          =   375
            Index           =   12
            Left            =   4440
            TabIndex        =   36
            Top             =   3600
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VM206 VALV. ALIM. TQ SPRAY"
            Height          =   375
            Index           =   11
            Left            =   4440
            TabIndex        =   30
            Top             =   2700
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "V3V202 VALV PU DIRECTO"
            Height          =   375
            Index           =   10
            Left            =   3000
            TabIndex        =   29
            Top             =   2700
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "V3V201 VALV. Q O PU"
            Height          =   375
            Index           =   9
            Left            =   1680
            TabIndex        =   28
            Top             =   2700
            Width           =   1095
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VM202 VALV. SALIDA DE TQ2"
            Height          =   375
            Index           =   8
            Left            =   240
            TabIndex        =   27
            Top             =   2700
            Width           =   1215
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VM235 VALV. VACIADO TQ2"
            Height          =   375
            Index           =   7
            Left            =   240
            TabIndex        =   22
            Top             =   1740
            Width           =   1095
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VM236 VALV. AGUA BARRIDO"
            Height          =   375
            Index           =   6
            Left            =   1560
            TabIndex        =   21
            Top             =   1740
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VM203 VALV. AGUA DE ALIM."
            Height          =   375
            Index           =   5
            Left            =   3000
            TabIndex        =   20
            Top             =   1740
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VM201 VALV. SALDA DE TQ1"
            Height          =   375
            Index           =   4
            Left            =   4440
            TabIndex        =   19
            Top             =   1740
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VM234 VALV. VACIADO TQ1"
            Height          =   375
            Index           =   3
            Left            =   4440
            TabIndex        =   14
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "EV211 VALV. AGUA A SELLOS"
            Height          =   375
            Index           =   2
            Left            =   3000
            TabIndex        =   12
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "CIP START / STOP"
            Height          =   375
            Index           =   1
            Left            =   1680
            TabIndex        =   10
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label LSwitch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Producción START/STOP"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   8
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label ControlTitle 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "SECTOR VALVULAS DE CIRCUITOS DE EVAPORACION"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   6
            Top             =   5160
            Width           =   5895
         End
      End
      Begin VB.PictureBox EBackground 
         Height          =   8225
         Left            =   60
         Picture         =   "Server.frx":0ECA
         ScaleHeight     =   544
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   935
         TabIndex        =   3
         Top             =   240
         Width           =   14085
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   75
            Left            =   10755
            Shape           =   3  'Circle
            Top             =   6900
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   74
            Left            =   1920
            Shape           =   3  'Circle
            Top             =   7440
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   73
            Left            =   1515
            Shape           =   3  'Circle
            Top             =   7440
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   72
            Left            =   1920
            Shape           =   3  'Circle
            Top             =   6915
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   71
            Left            =   1515
            Shape           =   3  'Circle
            Top             =   6915
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   70
            Left            =   1920
            Shape           =   3  'Circle
            Top             =   6390
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   69
            Left            =   1515
            Shape           =   3  'Circle
            Top             =   6390
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   68
            Left            =   1920
            Shape           =   3  'Circle
            Top             =   5880
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H0000FF00&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   67
            Left            =   1500
            Shape           =   3  'Circle
            Top             =   5880
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   66
            Left            =   13800
            Shape           =   3  'Circle
            Top             =   7080
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   65
            Left            =   12480
            Shape           =   3  'Circle
            Top             =   7560
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   64
            Left            =   11670
            Shape           =   3  'Circle
            Top             =   7755
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   63
            Left            =   9870
            Shape           =   3  'Circle
            Top             =   7755
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   62
            Left            =   9075
            Shape           =   3  'Circle
            Top             =   7560
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   61
            Left            =   12465
            Shape           =   3  'Circle
            Top             =   6585
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   60
            Left            =   9060
            Shape           =   3  'Circle
            Top             =   6585
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   59
            Left            =   7350
            Shape           =   3  'Circle
            Top             =   6870
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   58
            Left            =   6885
            Shape           =   3  'Circle
            Top             =   6600
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   57
            Left            =   6990
            Shape           =   3  'Circle
            Top             =   6870
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   56
            Left            =   6480
            Shape           =   3  'Circle
            Top             =   7050
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   55
            Left            =   5835
            Shape           =   3  'Circle
            Top             =   7440
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   54
            Left            =   5565
            Shape           =   3  'Circle
            Top             =   7440
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   53
            Left            =   5790
            Shape           =   3  'Circle
            Top             =   6870
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   52
            Left            =   5670
            Shape           =   3  'Circle
            Top             =   6600
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   51
            Left            =   13665
            Shape           =   3  'Circle
            Top             =   5760
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   50
            Left            =   11580
            Shape           =   3  'Circle
            Top             =   5880
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   49
            Left            =   10890
            Shape           =   3  'Circle
            Top             =   5760
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   48
            Left            =   10320
            Shape           =   3  'Circle
            Top             =   5760
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   47
            Left            =   12675
            Shape           =   3  'Circle
            Top             =   5760
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   46
            Left            =   9510
            Shape           =   3  'Circle
            Top             =   5760
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   45
            Left            =   8895
            Shape           =   3  'Circle
            Top             =   5760
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   44
            Left            =   7440
            Shape           =   3  'Circle
            Top             =   5760
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   43
            Left            =   6120
            Shape           =   3  'Circle
            Top             =   5760
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   42
            Left            =   5520
            Shape           =   3  'Circle
            Top             =   5760
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   41
            Left            =   4665
            Shape           =   3  'Circle
            Top             =   5760
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   40
            Left            =   4215
            Shape           =   3  'Circle
            Top             =   5760
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   39
            Left            =   3390
            Shape           =   3  'Circle
            Top             =   5760
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   38
            Left            =   2445
            Shape           =   3  'Circle
            Top             =   5760
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   37
            Left            =   11490
            Shape           =   3  'Circle
            Top             =   5355
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   36
            Left            =   5100
            Shape           =   3  'Circle
            Top             =   5340
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   35
            Left            =   12885
            Shape           =   3  'Circle
            Top             =   4740
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   34
            Left            =   12510
            Shape           =   3  'Circle
            Top             =   4740
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   33
            Left            =   11340
            Shape           =   3  'Circle
            Top             =   4740
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   32
            Left            =   10935
            Shape           =   3  'Circle
            Top             =   4740
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   31
            Left            =   3660
            Shape           =   3  'Circle
            Top             =   4710
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   30
            Left            =   3255
            Shape           =   3  'Circle
            Top             =   4710
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   29
            Left            =   3465
            Shape           =   3  'Circle
            Top             =   4320
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   28
            Left            =   12315
            Shape           =   3  'Circle
            Top             =   4065
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   27
            Left            =   11775
            Shape           =   3  'Circle
            Top             =   4065
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   26
            Left            =   10830
            Shape           =   3  'Circle
            Top             =   4065
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   25
            Left            =   9960
            Shape           =   3  'Circle
            Top             =   4050
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   24
            Left            =   9450
            Shape           =   3  'Circle
            Top             =   4050
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   23
            Left            =   8895
            Shape           =   3  'Circle
            Top             =   4050
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   22
            Left            =   7170
            Shape           =   3  'Circle
            Top             =   4050
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   21
            Left            =   6195
            Shape           =   3  'Circle
            Top             =   4050
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   20
            Left            =   5610
            Shape           =   3  'Circle
            Top             =   4050
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   19
            Left            =   5025
            Shape           =   3  'Circle
            Top             =   4050
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   18
            Left            =   4290
            Shape           =   3  'Circle
            Top             =   4050
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   17
            Left            =   3720
            Shape           =   3  'Circle
            Top             =   4050
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   16
            Left            =   3060
            Shape           =   3  'Circle
            Top             =   4050
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   15
            Left            =   3510
            Shape           =   3  'Circle
            Top             =   2730
            Width           =   120
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   36
            Left            =   735
            Shape           =   3  'Circle
            Top             =   660
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   14
            Left            =   1995
            Shape           =   3  'Circle
            Top             =   4590
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   13
            Left            =   1425
            Shape           =   3  'Circle
            Top             =   4590
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   12
            Left            =   2265
            Shape           =   3  'Circle
            Top             =   4200
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   11
            Left            =   1170
            Shape           =   3  'Circle
            Top             =   4200
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   10
            Left            =   13815
            Shape           =   3  'Circle
            Top             =   1320
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   9
            Left            =   12405
            Shape           =   3  'Circle
            Top             =   1305
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   8
            Left            =   10980
            Shape           =   3  'Circle
            Top             =   1305
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   7
            Left            =   7365
            Shape           =   3  'Circle
            Top             =   1290
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   6
            Left            =   3495
            Shape           =   3  'Circle
            Top             =   1260
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   5
            Left            =   2580
            Shape           =   3  'Circle
            Top             =   1995
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   4
            Left            =   2295
            Shape           =   3  'Circle
            Top             =   1995
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   3
            Left            =   1140
            Shape           =   3  'Circle
            Top             =   1995
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   2
            Left            =   840
            Shape           =   3  'Circle
            Top             =   1995
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   1
            Left            =   375
            Shape           =   3  'Circle
            Top             =   1995
            Width           =   120
         End
         Begin VB.Shape LEDG 
            BackColor       =   &H00F4FFE5&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   0
            Left            =   210
            Shape           =   3  'Circle
            Top             =   1995
            Width           =   120
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   35
            Left            =   13800
            Shape           =   3  'Circle
            Top             =   6720
            Width           =   120
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   34
            Left            =   10770
            Shape           =   3  'Circle
            Top             =   7530
            Width           =   120
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   33
            Left            =   10770
            Shape           =   3  'Circle
            Top             =   7350
            Width           =   120
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   32
            Left            =   11400
            Shape           =   3  'Circle
            Top             =   6600
            Width           =   120
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   31
            Left            =   10170
            Shape           =   3  'Circle
            Top             =   6600
            Width           =   120
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   30
            Left            =   8100
            Shape           =   3  'Circle
            Top             =   6855
            Width           =   120
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   29
            Left            =   4905
            Shape           =   3  'Circle
            Top             =   6870
            Width           =   120
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   28
            Left            =   11940
            Shape           =   3  'Circle
            Top             =   5370
            Width           =   120
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   27
            Left            =   6735
            Shape           =   3  'Circle
            Top             =   5370
            Width           =   120
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   26
            Left            =   5595
            Shape           =   3  'Circle
            Top             =   5370
            Width           =   120
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   25
            Left            =   3450
            Shape           =   3  'Circle
            Top             =   5370
            Width           =   120
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   24
            Left            =   13260
            Shape           =   3  'Circle
            Top             =   1860
            Width           =   120
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   23
            Left            =   12525
            Shape           =   3  'Circle
            Top             =   2940
            Width           =   120
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   22
            Left            =   12060
            Shape           =   3  'Circle
            Top             =   1860
            Width           =   120
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   21
            Left            =   10515
            Shape           =   3  'Circle
            Top             =   1860
            Width           =   120
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   20
            Left            =   9720
            Shape           =   3  'Circle
            Top             =   2910
            Width           =   120
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   19
            Left            =   8670
            Shape           =   3  'Circle
            Top             =   2910
            Width           =   120
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   18
            Left            =   7815
            Shape           =   3  'Circle
            Top             =   1860
            Width           =   120
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   17
            Left            =   6960
            Shape           =   3  'Circle
            Top             =   2910
            Width           =   120
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   16
            Left            =   5820
            Shape           =   3  'Circle
            Top             =   2910
            Width           =   120
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   15
            Left            =   4515
            Shape           =   3  'Circle
            Top             =   2910
            Width           =   120
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   14
            Left            =   4035
            Shape           =   3  'Circle
            Top             =   1800
            Width           =   120
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   13
            Left            =   2430
            Shape           =   3  'Circle
            Top             =   3750
            Width           =   120
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   12
            Left            =   2430
            Shape           =   3  'Circle
            Top             =   3210
            Width           =   120
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   11
            Left            =   2880
            Shape           =   3  'Circle
            Top             =   3210
            Width           =   120
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   10
            Left            =   2880
            Shape           =   3  'Circle
            Top             =   2400
            Width           =   120
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   9
            Left            =   990
            Shape           =   3  'Circle
            Top             =   3750
            Width           =   120
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   8
            Left            =   1005
            Shape           =   3  'Circle
            Top             =   3210
            Width           =   120
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   7
            Left            =   540
            Shape           =   3  'Circle
            Top             =   3210
            Width           =   120
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   120
            Index           =   6
            Left            =   540
            Shape           =   3  'Circle
            Top             =   2400
            Width           =   120
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   150
            Index           =   5
            Left            =   12120
            Shape           =   3  'Circle
            Top             =   465
            Width           =   150
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   150
            Index           =   4
            Left            =   10110
            Shape           =   3  'Circle
            Top             =   450
            Width           =   150
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   150
            Index           =   3
            Left            =   8115
            Shape           =   3  'Circle
            Top             =   450
            Width           =   150
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   150
            Index           =   2
            Left            =   6105
            Shape           =   3  'Circle
            Top             =   450
            Width           =   150
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   150
            Index           =   1
            Left            =   4080
            Shape           =   3  'Circle
            Top             =   450
            Width           =   150
         End
         Begin VB.Shape LEDR 
            BackColor       =   &H00DBDBFE&
            BackStyle       =   1  'Opaque
            Height          =   150
            Index           =   0
            Left            =   2040
            Shape           =   3  'Circle
            Top             =   450
            Width           =   150
         End
         Begin VB.Label Alarm 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   210
            Top             =   -120
            Visible         =   0   'False
            Width           =   14055
         End
      End
   End
   Begin VB.PictureBox Logon 
      BorderStyle     =   0  'None
      Height          =   1400
      Left            =   120
      Picture         =   "Server.frx":25CC2
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   0
      Top             =   120
      Width           =   2880
      Begin VB.Timer SSInt 
         Interval        =   1000
         Left            =   2760
         Top             =   600
      End
      Begin VB.Timer LedControl 
         Interval        =   500
         Left            =   2760
         Top             =   1080
      End
      Begin VB.Timer UserCV 
         Interval        =   3500
         Left            =   2760
         Top             =   120
      End
      Begin VB.Timer HKeys 
         Interval        =   500
         Left            =   2760
         Top             =   -240
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Alarma:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14280
      TabIndex        =   217
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label EAgua 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ev. Agua:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14280
      TabIndex        =   214
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label PDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   17640
      TabIndex        =   209
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label PanelTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SECTOR VALVULAS DE CIRCUITOS DE EVAPORACIÓN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   31
      Top             =   1080
      Width           =   11175
   End
   Begin VB.Label Mark 
      BackStyle       =   0  'Transparent
      Caption         =   "Developed by Maatooh-Software © 2023"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15840
      TabIndex        =   2
      Top             =   10800
      Width           =   3735
   End
End
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AlState As Long
Dim k As Long
Dim j As Long

Private Sub BAlarm_Click()
'--Switch Function
If (BAlarm.Caption = "ON") Then
BAlarm.Caption = "OFF"
BAlarm.BackColor = &H8080FF
Call ControlFx.SendValvStatus("[CMD V2]:" & ControlFx.SwitchState(2), "vDev2")
GoTo ECase
End If
If (BAlarm.Caption = "OFF") Then
BAlarm.Caption = "ON"
BAlarm.BackColor = &H80FF80
Call ControlFx.SendValvStatus("[CMD V2]:" & ControlFx.SwitchState(2), "vDev2")
GoTo ECase
End If
'-----------------
ECase:
'---Save----------
SaveFx.SaveSwitch
'-----------------
'Debug.Print Index
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
For X = 0 To LEDR.UBound
LEDR(X).BackColor = &HDBDBFE
Next X
LEDR(k).BackColor = &HFF&
k = k + 1
If k > LEDR.UBound Then k = 0
Case 1
For X = 0 To LEDG.UBound
LEDG(X).BackColor = &HF4FFE5
Next X
LEDG(j).BackColor = &HFF00&
j = j + 1
If j > LEDG.UBound Then j = 0
End Select
End Sub

Private Sub CommandTab_Click(Index As Integer)
Select Case Index
Case 0
PanelTitle = "SECTOR VALVULAS DE CIRCUITOS DE EVAPORACIÓN"
ControlFrame(0).Visible = True
ControlFrame(1).Visible = False
ControlFrame(2).Visible = False
ControlFrame(3).Visible = False
Case 1
PanelTitle = "SECTOR EVAPORADOR QUINTUPLE"
ControlFrame(0).Visible = False
ControlFrame(1).Visible = True
ControlFrame(2).Visible = False
ControlFrame(3).Visible = False
Case 2
PanelTitle = "SECTOR EVAPORADOR PU."
ControlFrame(0).Visible = False
ControlFrame(1).Visible = False
ControlFrame(2).Visible = True
ControlFrame(3).Visible = False
Case 3
PanelTitle = "SECTOR VALVULAS DE CIRCUITOS CIP"
ControlFrame(0).Visible = False
ControlFrame(1).Visible = False
ControlFrame(2).Visible = False
ControlFrame(3).Visible = True
End Select
End Sub

Private Sub EvAgua_Click()
'--Switch Function
If (EvAgua.Caption = "ON") Then
EvAgua.Caption = "OFF"
EvAgua.BackColor = &H8080FF
Call ControlFx.SendValvStatus("[CMD V2]:" & ControlFx.SwitchState(2), "vDev2")
GoTo ECase
End If
If (EvAgua.Caption = "OFF") Then
EvAgua.Caption = "ON"
EvAgua.BackColor = &H80FF80
Call ControlFx.SendValvStatus("[CMD V2]:" & ControlFx.SwitchState(2), "vDev2")
GoTo ECase
End If
'-----------------
ECase:
'---Save----------
SaveFx.SaveSwitch
'-----------------
'Debug.Print Index
End Sub

Private Sub EvAM_Click()
'--Switch Function
If (EvAM.Caption = "Auto") Then
EvAM.Caption = "Man"
EvAM.BackColor = &H8080FF
EvAgua.Enabled = True
GoTo ECase
End If
If (EvAM.Caption = "Man") Then
EvAM.Caption = "Auto"
EvAM.BackColor = &H80FF80
EvAgua.Enabled = False
GoTo ECase
End If
'-----------------
ECase:
'---Save----------
SaveFx.SaveSwitch
'-----------------
'Debug.Print Index
End Sub

Private Sub Form_Load()
ScreenFx.ScreenFix
ScreenFx.Mark
ServerLog.Hide
SaveFx.LoadSwitch
ControlFrame(0).Visible = True
ReDim Preserve DevicesFx.DevName(DevicesFx.TUsers)
ReDim Preserve DevicesFx.DevAct(DevicesFx.TUsers)
'-------------
With WServer(TUsers)
.Close
.LocalPort = 5225
.Listen
End With
'-------------
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub HKeys_Timer()
'--Escape
If Not GetAsyncKeyState(vbKeyEscape) = 0 Then
Select Case ScreenFx.WinState
Case True
Call ScreenFx.SStyle.FullContentFx(Server.hWnd, True)
ScreenFx.WinState = False
Server.WindowState = vbMaximized
Case False
Call ScreenFx.SStyle.FullContentFx(Server.hWnd, False)
ScreenFx.WinState = True
Server.WindowState = vbMaximized
End Select
End If
'--Ctrl + F4
If Not GetAsyncKeyState(vbKeyControl) = 0 And Not GetAsyncKeyState(vbKeyF4) = 0 Then
End
End If
'--Ctrl + Q
If Not GetAsyncKeyState(vbKeyControl) = 0 And Not GetAsyncKeyState(vbKeyQ) = 0 Then
ServerLog.Hide
End If
'--Ctrl + W
If Not GetAsyncKeyState(vbKeyControl) = 0 And Not GetAsyncKeyState(vbKeyW) = 0 Then
ServerLog.Show
End If
PDate = Time & "    " & Format(Date, "dd/mm/yy")
End Sub

Private Sub LedControl_Timer()
'Led Test
If Server.Switch3(16).Caption = "ON" Then
LedFx.AllOn
Exit Sub
End If
If Server.Switch3(16).Caption = "OFF" Then LedFx.AllOff
'Select Switch0
For l = 0 To Server.Switch0.UBound
LedFx.LedFunctionsS0 (l)
Next l
'Select Switch1
For l = 0 To Server.Switch1.UBound
LedFx.LedFunctionsS1 (l)
Next l
'Select Switch2
For l = 0 To Server.Switch2.UBound
LedFx.LedFunctionsS2 (l)
Next l
'Select Switch3
For l = 0 To Server.Switch3.UBound
LedFx.LedFunctionsS3 (l)
Next l
'Alarm Trigger
Dim Atr As Long
For r = 0 To 5
If Server.LEDR(r).BackColor = &HFF& Then Atr = Atr + 1
Next r
'-Insert----
'-Desactiva
    If Atr <= 1 Then
    AlState = Atr
    End If
'-Activa
If Atr > 1 Then
    If (BAlarm.Caption = "OFF") And Not AlState = Atr Then
    BAlarm.Caption = "ON"
    BAlarm.BackColor = &H80FF80
    Call ControlFx.SendValvStatus("[CMD V2]:" & ControlFx.SwitchState(2), "vDev2")
    End If
'---------------------------
    If Atr > AlState Then
    AlState = Atr
    End If
    If Atr < AlState Then
    AlState = Atr
    End If
'---------------------------
End If
'-----------
End Sub

Private Sub RefreshDev_Click()
ControlFx.RefreshState
End Sub

Private Sub SDevices_Click()
Devices.Show
End Sub


Private Sub Sensor_DblClick(Index As Integer)
If Sensor(Index).BackColor = &H80FF& Then
ViewParams.LoadParam (Index)
End If
End Sub

Private Sub SFrameNB_Click(Index As Integer)
Select Case Index
Case 0
SensorFrame(0).Visible = False
SensorFrame(1).Visible = True
Case 1
SensorFrame(0).Visible = True
SensorFrame(1).Visible = False
End Select
End Sub

Private Sub SSInt_Timer()
'----Switch Register
If Not SaveSensor.DRCont = Minute(Time) Then
'-------
Select Case Minute(Time)
Case 0, 2, 4, 6, 8
Debug.Print "Nuevo registro " & Minute(Time)
Call SaveSensor.SaveRegister
SaveSensor.DRCont = Minute(Time)
Case 10, 12, 14, 16, 18
Debug.Print "Nuevo registro " & Minute(Time)
Call SaveSensor.SaveRegister
SaveSensor.DRCont = Minute(Time)
Case 20, 22, 24, 26, 28
Debug.Print "Nuevo registro " & Minute(Time)
Call SaveSensor.SaveRegister
SaveSensor.DRCont = Minute(Time)
Case 30, 32, 34, 36, 38
Debug.Print "Nuevo registro " & Minute(Time)
Call SaveSensor.SaveRegister
SaveSensor.DRCont = Minute(Time)
Case 40, 42, 44, 46, 48
Debug.Print "Nuevo registro " & Minute(Time)
Call SaveSensor.SaveRegister
SaveSensor.DRCont = Minute(Time)
Case 50, 52, 54, 56, 58
Debug.Print "Nuevo registro " & Minute(Time)
Call SaveSensor.SaveRegister
SaveSensor.DRCont = Minute(Time)
End Select
'-------
End If

End Sub

Private Sub Switch0_Click(Index As Integer)
'--Switch Function
If (Switch0(Index).Caption = "ON") Then
Switch0(Index).Caption = "OFF"
Switch0(Index).BackColor = &H8080FF
Call ControlFx.SendValvStatus("[CMD V0]:" & ControlFx.SwitchState(0), "vDev0")
GoTo ECase
End If
If (Switch0(Index).Caption = "OFF") Then
Switch0(Index).Caption = "ON"
Switch0(Index).BackColor = &H80FF80
Call ControlFx.SendValvStatus("[CMD V0]:" & ControlFx.SwitchState(0), "vDev0")
GoTo ECase
End If
'-----------------
ECase:
'---Save----------
SaveFx.SaveSwitch
'-----------------
'Debug.Print ControlFx.SwitchState(0)
'Debug.Print Index
End Sub

Private Sub Switch1_Click(Index As Integer)
'--Switch Function
If (Switch1(Index).Caption = "ON") Then
Switch1(Index).Caption = "OFF"
Switch1(Index).BackColor = &H8080FF
Call ControlFx.SendValvStatus("[CMD V1]:" & ControlFx.SwitchState(1), "vDev1")
GoTo ECase
End If
If (Switch1(Index).Caption = "OFF") Then
Switch1(Index).Caption = "ON"
Switch1(Index).BackColor = &H80FF80
Call ControlFx.SendValvStatus("[CMD V1]:" & ControlFx.SwitchState(1), "vDev1")
GoTo ECase
End If
'-----------------
ECase:
'---Save----------
SaveFx.SaveSwitch
'-----------------
'Debug.Print Index
End Sub

Private Sub Switch2_Click(Index As Integer)
'--Switch Function
If (Switch2(Index).Caption = "ON") Then
Switch2(Index).Caption = "OFF"
Switch2(Index).BackColor = &H8080FF
Call ControlFx.SendValvStatus("[CMD V2]:" & ControlFx.SwitchState(2), "vDev2")
GoTo ECase
End If
If (Switch2(Index).Caption = "OFF") Then
Switch2(Index).Caption = "ON"
Switch2(Index).BackColor = &H80FF80
Call ControlFx.SendValvStatus("[CMD V2]:" & ControlFx.SwitchState(2), "vDev2")
GoTo ECase
End If
'-----------------
ECase:
'---Save----------
SaveFx.SaveSwitch
'-----------------
'Debug.Print Index
End Sub

Private Sub Switch3_Click(Index As Integer)
'--Switch Function
If (Switch3(Index).Caption = "ON") Then
Switch3(Index).Caption = "OFF"
Switch3(Index).BackColor = &H8080FF
Call ControlFx.SendValvStatus("[CMD V3]:" & ControlFx.SwitchState(3), "vDev3")
GoTo ECase
End If
If (Switch3(Index).Caption = "OFF") Then
Switch3(Index).Caption = "ON"
Switch3(Index).BackColor = &H80FF80
Call ControlFx.SendValvStatus("[CMD V3]:" & ControlFx.SwitchState(3), "vDev3")
GoTo ECase
End If
'-----------------
ECase:
'---Save----------
SaveFx.SaveSwitch
'-----------------
'Debug.Print Index
End Sub

Private Sub UserCV_Timer()
Devices.DeviceList.Clear
On Error GoTo QH
For w = 0 To WServer.UBound
If WServer(w).State = sckConnected And Not DevicesFx.DevName(w) = vbNullString And DevicesFx.DevAct(w) = True Then
Devices.DeviceList.AddItem (DevicesFx.DevName(w))
GoTo RS
End If
If WServer(w).State = sckConnected And Not DevicesFx.DevName(w) = vbNullString And DevicesFx.DevAct(w) = False Then
WServer(w).Close
GoTo RS
End If
RS:
Next w
'---SetBack
For w = 0 To WServer.UBound
DevicesFx.DevAct(w) = False
Next w
QH:
End Sub

Private Sub WServer_Close(Index As Integer)
'WServer(Index).Close
End Sub

Private Sub WServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
DevicesFx.TUsers = DevicesFx.TUsers + 1
Load WServer(DevicesFx.TUsers)
WServer(DevicesFx.TUsers).Accept requestID
Debug.Print "--Connect New User--"
Debug.Print WServer(DevicesFx.TUsers).RemoteHostIP
ServerLog.Log = ServerLog.Log & Time & "    " & Format(Date, "dd/mm/yy") & vbCrLf
ServerLog.Log = ServerLog.Log & "--Connect New User--" & vbCrLf
ServerLog.Log = ServerLog.Log & WServer(DevicesFx.TUsers).RemoteHostIP & vbCrLf
ServerLog.Log.SelStart = Len(ServerLog.Log)
DevicesFx.delay (1.5)
ControlFx.RefreshState
DoEvents
End Sub

Private Sub WServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error GoTo QH
Dim xData As String
WServer(Index).GetData xData
'---Connected
If Not InStr(xData, "- connected") = 0 Then Call DevicesFx.RegDevice(xData, CLng(Index))
'---Ping
If Not InStr(xData, Chr(1)) = 0 Then xData = DevicesFx.PingDx(xData, CLng(Index))
'---Functions
If Not InStr(xData, "[CMD S") = 0 Then Call ControlFx.ClassParamsSensor(xData)
If Not InStr(xData, "[CMD L") = 0 Then Call ControlFx.SetLeds(xData)

QH:
End Sub

Private Sub WServer_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'WServer(Index).Close
End Sub
