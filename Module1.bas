Attribute VB_Name = "Module1"
Option Explicit

Sub Simular_Fluxo_Vagoes_Consolidado()

    Dim ws As Worksheet
    Dim i As Integer, j As Integer, k As Integer, indexL4 As Integer
    Dim vagoesL1 As Variant, vagoesL2 As Variant, vagoesL3 As Variant, vagoesL4 As Variant
    Dim azul As Long, branco As Long, cinza As Long
    Dim alvoL1 As Double, alvoL2 As Double, alvoAB16 As Double
    Dim topoInicialCaboL1 As Double, topoInicialCaboL2 As Double

    ' >>> AJUSTE <<< posições corretas de descarga
    Dim posDescargaL1 As Double
    Dim posDescargaL2 As Double

    Set ws = ActiveSheet

    azul = RGB(0, 112, 192)
    branco = RGB(255, 255, 255)
    cinza = RGB(192, 192, 192)

    ' ================= CONFIGURAÇÃO DOS ARRAYS =================
    vagoesL1 = Array("V_L1_16", "V_L1_15", "V_L1_14", "V_L1_13", "V_L1_12", "V_L1_11", "V_L1_10", "V_L1_09", _
                     "V_L1_08", "V_L1_07", "V_L1_06", "V_L1_05", "V_L1_04", "V_L1_03", "V_L1_02", "V_L1_01")

    vagoesL2 = Array("V_L2_32", "V_L2_31", "V_L2_30", "V_L2_29", "V_L2_28", "V_L2_27", "V_L2_26", "V_L2_25", _
                     "V_L2_24", "V_L2_23", "V_L2_22", "V_L2_21", "V_L2_20", "V_L2_19", "V_L2_18", "V_L2_17")

    vagoesL3 = Array("V_L3_48", "V_L3_47", "V_L3_46", "V_L3_45", "V_L3_44", "V_L3_43", "V_L3_42", "V_L3_41", _
                     "V_L3_40", "V_L3_39", "V_L3_38", "V_L3_37", "V_L3_36", "V_L3_35", "V_L3_34", "V_L3_33")

    vagoesL4 = Array("V_L4_64", "V_L4_63", "V_L4_62", "V_L4_61", "V_L4_60", "V_L4_59", "V_L4_58", "V_L4_57", _
                     "V_L4_56", "V_L4_55", "V_L4_54", "V_L4_53", "V_L4_52", "V_L4_51", "V_L4_50", "V_L4_49")

    alvoAB16 = ws.Range("AB16").Left

    ' ================= RESET =================
    Limpar_Tudo ws
    ws.Shapes("V_L1_16_A").Visible = True
    ws.Shapes("V_L1_12_A").Visible = True
    ws.Shapes("V_L2_32_A").Visible = True
    ws.Shapes("V_L2_28_A").Visible = True

    ws.Shapes("PORTICO_L1_CABO").ZOrder msoBringToFront
    ws.Shapes("PORTICO_L2_CABO").ZOrder msoBringToFront

    ' ================= ENTRADA L1, L2, L3 =================
    Dim todasL As Variant
    todasL = Array(vagoesL1, vagoesL2, vagoesL3)

    For i = 0 To 2
        For j = 0 To 15
            Pintar_Par ws, todasL(i)(j), azul
            DoEvents
        Next j
    Next i

    ' ================= ENTRADA L4 + LIMPEZA LADO A =================
    indexL4 = 0
    For j = 0 To 15
        Pintar_Lado_A_Apenas ws, vagoesL1(j), branco
        Pintar_Lado_A_Apenas ws, vagoesL2(j), branco

        If j >= 8 Then
            If indexL4 <= 15 Then
                Pintar_Par ws, vagoesL4(indexL4), azul
                indexL4 = indexL4 + 1
            End If
        End If
        DoEvents
    Next j

    Do While indexL4 <= 15
        Pintar_Par ws, vagoesL4(indexL4), azul
        indexL4 = indexL4 + 1
        DoEvents
    Loop

    ' ================= SAÍDA LADO B =================
    For j = 15 To 0 Step -1
        Pintar_Lado_B_Apenas ws, vagoesL3(j), branco
        Pintar_Lado_B_Apenas ws, vagoesL4(j), branco
        DoEvents
    Next j

    ' ================= DESCARGA 1 =================
    ws.Shapes("V_L1_16_A").ZOrder msoBringToFront
    ws.Shapes("V_L1_12_A").ZOrder msoBringToFront

    topoInicialCaboL1 = ws.Shapes("PORTICO_L1_CABO").Top
    topoInicialCaboL2 = ws.Shapes("PORTICO_L2_CABO").Top

    ' Descida
    MoverCaboSimultaneo ws, "PORTICO_L1_CABO", "PORTICO_L2_CABO", 95, 1

    ws.Shapes("V_L1_16_A_BASE").Fill.ForeColor.RGB = cinza
    ws.Shapes("V_L1_12_A_BASE").Fill.ForeColor.RGB = cinza

    ' Içamento curto
    MoverCaboComCargaSimultaneo ws, _
        "PORTICO_L1_CABO", "V_L1_16_A", _
        "PORTICO_L2_CABO", "V_L1_12_A", _
        25, -1

    ' Descida final
    MoverCaboComCargaSimultaneo ws, _
        "PORTICO_L1_CABO", "V_L1_16_A", _
        "PORTICO_L2_CABO", "V_L1_12_A", _
        112, 1

    ' >>> AJUSTE <<< salva posição correta de descarga
    posDescargaL1 = ws.Shapes("TRATOR_L1").Left
    posDescargaL2 = ws.Shapes("TRATOR_L2").Left

    ' ================= MOVIMENTO CONJUNTO =================
    Do While _
        ws.Shapes("TRATOR_L3").Left > posDescargaL1 _
        Or ws.Shapes("TRATOR_L4").Left > posDescargaL2 _
        Or ws.Shapes("PORTICO_L1_CABO").Top > topoInicialCaboL1 _
        Or ws.Shapes("PORTICO_L2_CABO").Top > topoInicialCaboL2

        ws.Shapes("TRATOR_L1").Left = ws.Shapes("TRATOR_L1").Left - 5
        ws.Shapes("V_L1_12_A").Left = ws.Shapes("V_L1_12_A").Left - 5

        ws.Shapes("TRATOR_L2").Left = ws.Shapes("TRATOR_L2").Left - 5
        ws.Shapes("V_L1_16_A").Left = ws.Shapes("V_L1_16_A").Left - 5

        If ws.Shapes("TRATOR_L3").Left > posDescargaL1 Then
            ws.Shapes("TRATOR_L3").Left = ws.Shapes("TRATOR_L3").Left - 5
        End If

        If ws.Shapes("TRATOR_L4").Left > posDescargaL2 Then
            ws.Shapes("TRATOR_L4").Left = ws.Shapes("TRATOR_L4").Left - 5
        End If

        If ws.Shapes("PORTICO_L1_CABO").Top > topoInicialCaboL1 Then
            ws.Shapes("PORTICO_L1_CABO").Top = ws.Shapes("PORTICO_L1_CABO").Top - 2
        End If

        If ws.Shapes("PORTICO_L2_CABO").Top > topoInicialCaboL2 Then
            ws.Shapes("PORTICO_L2_CABO").Top = ws.Shapes("PORTICO_L2_CABO").Top - 2
        End If

        DoEvents
    Loop

    ' ================= SAÍDA DOS TRATORES 1 E 2 =================
    Do While ws.Shapes("TRATOR_L1").Left > -500 Or ws.Shapes("TRATOR_L2").Left > -500

        ws.Shapes("TRATOR_L1").Left = ws.Shapes("TRATOR_L1").Left - 6
        ws.Shapes("V_L1_12_A").Left = ws.Shapes("V_L1_12_A").Left - 6

        ws.Shapes("TRATOR_L2").Left = ws.Shapes("TRATOR_L2").Left - 6
        ws.Shapes("V_L1_16_A").Left = ws.Shapes("V_L1_16_A").Left - 6

        DoEvents
    Loop

    MsgBox "Fluxo Concluído!", vbInformation
End Sub

' ================= SUBS AUXILIARES =================

Sub Limpar_Tudo(ws As Worksheet)
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Name Like "V_L*" Then shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
    Next shp
End Sub

Sub Pintar_Par(ws As Worksheet, p As Variant, c As Long)
    On Error Resume Next
    ws.Shapes(p & "_A").Fill.ForeColor.RGB = c
    ws.Shapes(p & "_B").Fill.ForeColor.RGB = c
    ws.Shapes(p & "_A_BASE").Fill.ForeColor.RGB = c
    ws.Shapes(p & "_B_BASE").Fill.ForeColor.RGB = c
End Sub

Sub Pintar_Lado_A_Apenas(ws As Worksheet, p As Variant, c As Long)
    On Error Resume Next
    ws.Shapes(p & "_A").Fill.ForeColor.RGB = c
    ws.Shapes(p & "_A_BASE").Fill.ForeColor.RGB = c
End Sub

Sub Pintar_Lado_B_Apenas(ws As Worksheet, p As Variant, c As Long)
    On Error Resume Next
    ws.Shapes(p & "_B").Fill.ForeColor.RGB = c
    ws.Shapes(p & "_B_BASE").Fill.ForeColor.RGB = c
End Sub

Sub MoverCaboSimultaneo(ws As Worksheet, c1 As String, c2 As String, d As Integer, s As Integer)
    Dim k As Integer
    For k = 1 To d
        ws.Shapes(c1).Top = ws.Shapes(c1).Top + s
        ws.Shapes(c2).Top = ws.Shapes(c2).Top + s
        DoEvents
    Next k
End Sub

Sub MoverCaboComCargaSimultaneo(ws As Worksheet, _
    c1 As String, v1 As String, _
    c2 As String, v2 As String, _
    d As Integer, s As Integer)

    Dim k As Integer
    For k = 1 To d
        ws.Shapes(c1).Top = ws.Shapes(c1).Top + s
        ws.Shapes(v1).Top = ws.Shapes(v1).Top + s
        ws.Shapes(c2).Top = ws.Shapes(c2).Top + s
        ws.Shapes(v2).Top = ws.Shapes(v2).Top + s
        DoEvents
    Next k
End Sub


