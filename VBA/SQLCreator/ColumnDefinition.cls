VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public DataType As String
Public category As String
Public Precision As Long
Public HasPrecision As Boolean
Public ScaleDigits As Long
Public HasScale As Boolean
Public IsPrimaryKey As Boolean
Public IsNotNull As Boolean

