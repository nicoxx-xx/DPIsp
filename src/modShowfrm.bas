Attribute VB_Name = "modShowfrm"
' ==============================
' Copyright (C) 2026 Domenico Longo
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, version 3.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <https://www.gnu.org/licenses/>.
' ==============================

Option Explicit

Public Const MOD_INPUT_FORM_VERSION As String = "v2.0.0"

Public Sub MostraGestioneDPI()
    frmDPI.Show
End Sub

' =====================================================
'   VERSIONING
' =====================================================
Public Function GetInputFormPanelVersion() As String
    GetInputFormPanelVersion = "CRUD panel " & MOD_INPUT_FORM_VERSION & ";"
End Function
