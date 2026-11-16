Attribute VB_Name = "TrocarCor"
'=====================================================================
'  TROCAR COR  v1  -  macro para CorelDRAW (VBA)        by Bruno Faqui
'
'  Troca uma cor por outra sem usar o Localizar e substituir do Corel.
'  - Procura em preenchimento, degradê, contorno, texto (letra por
'    letra), sombra e contour.
'  - Entra em grupos e PowerClips, em todas as camadas, inclusive
'    objetos e camadas travados (destrava, troca e trava de novo).
'  - Compara com tolerância de 1% por canal (acha 99% quando procura
'    100%), no mesmo modelo de cor (CMYK com CMYK, RGB com RGB).
'  - Mostra quantas achou antes de trocar.
'  - Tudo vira um passo só de desfazer: um Ctrl+Z volta tudo.
'
'  COMO USAR
'  1. Clique num objeto que tem a cor que vai sair. Se ele estiver
'     dentro de um grupo, use Ctrl + clique. Sem nada selecionado,
'     a macro pede a cor numa janela.
'  2. Ferramentas > Macros > Executar macro > TrocarCor.
'=====================================================================
Option Explicit

' Diferença aceita em cada canal para considerar a mesma cor
Private Const TOL_CMYK As Long = 1      ' CMYK, em %
Private Const TOL_RGB As Long = 2       ' RGB, de 0 a 255
Private Const TITULO As String = "Trocar cor"

Private tcAlvo As Color          ' cor que vai sair (cópia)
Private tcNova As Color          ' cor que entra
Private tcContar As Boolean      ' True = só conta, não troca
Private tcNoPreench As Boolean, tcNoContorno As Boolean
Private qPreench As Long, qDegrade As Long, qContorno As Long
Private qTexto As Long, qEfeito As Long, qErro As Long

Public Sub TrocarCor()
    Dim doc As Document
    Dim pgOriginal As Page
    Dim resp As VbMsgBoxResult
    Dim onde As String, msg As String
    Dim todas As Boolean, falhou As Boolean
    Dim total As Long

    If Documents.Count = 0 Then
        MsgBox "Abra um documento antes de rodar a macro.", vbExclamation, TITULO
        Exit Sub
    End If
    Set doc = ActiveDocument
    Set pgOriginal = doc.ActivePage

    ' 1. Cor que vai sair
    If Not tcPegarCorAlvo() Then Exit Sub

    ' 2. Cor nova
    MsgBox "Cor que vai sair: " & tcDescrever(tcAlvo) & vbCr & vbCr & _
           "Na próxima janela, escolha a NOVA cor.", vbInformation, TITULO
    Set tcNova = New Color
    tcNova.CopyAssign tcAlvo
    If Not tcNova.UserAssignEx() Then Exit Sub
    If tcNova.IsSame(tcAlvo) Then
        MsgBox "A nova cor é igual à que vai sair. Nada foi trocado.", vbInformation, TITULO
        Exit Sub
    End If

    ' 3. Onde trocar
    onde = InputBox("Trocar " & tcDescrever(tcAlvo) & vbCr & _
                    "por " & tcDescrever(tcNova) & vbCr & vbCr & _
                    "Onde?" & vbCr & _
                    "1 = preenchimento e contorno" & vbCr & _
                    "2 = só preenchimento" & vbCr & _
                    "3 = só contorno", TITULO, "1")
    onde = Trim$(onde)
    If Len(onde) = 0 Then Exit Sub
    tcNoPreench = (onde = "1" Or onde = "2")
    tcNoContorno = (onde = "1" Or onde = "3")
    If Not tcNoPreench And Not tcNoContorno Then
        MsgBox "Opção inválida: " & onde & ". Use 1, 2 ou 3.", vbExclamation, TITULO
        Exit Sub
    End If

    ' 4. Páginas
    todas = True
    If doc.Pages.Count > 1 Then
        resp = MsgBox("Trocar em TODAS as " & doc.Pages.Count & " páginas?" & vbCr & vbCr & _
                      "Sim = todas as páginas" & vbCr & _
                      "Não = só a página " & pgOriginal.Index, vbYesNoCancel + vbQuestion, TITULO)
        If resp = vbCancel Then Exit Sub
        todas = (resp = vbYes)
    End If

    ' 5. Conta antes de trocar
    tcContar = True
    tcZerar
    tcPercorrer doc, pgOriginal, todas
    total = qPreench + qDegrade + qContorno + qTexto + qEfeito
    If total = 0 Then
        MsgBox "Não achei " & tcDescrever(tcAlvo) & IIf(todas, " no arquivo.", " nesta página.") & vbCr & vbCr & _
               "A macro compara no mesmo modelo de cor. Se a arte está em RGB e a cor foi escolhida " & _
               "em CMYK (ou o contrário), clique num objeto da arte que tem a cor e rode de novo.", _
               vbInformation, TITULO
        Exit Sub
    End If
    resp = MsgBox("Achei:" & tcResumo() & vbCr & vbCr & _
                  "Trocar " & tcDescrever(tcAlvo) & vbCr & _
                  "por " & tcDescrever(tcNova) & "?", vbOKCancel + vbQuestion, TITULO)
    If resp <> vbOK Then Exit Sub

    ' 6. Troca
    tcContar = False
    tcZerar
    On Error GoTo Falha
    doc.BeginCommandGroup "Trocar cor"
    Optimization = True
    tcPercorrer doc, pgOriginal, todas

Concluir:
    On Error Resume Next
    pgOriginal.Activate
    Optimization = False
    doc.EndCommandGroup
    ActiveWindow.Refresh
    Application.Refresh
    On Error GoTo 0

    If Not falhou Then
        msg = "Pronto. Trocado:" & tcResumo() & vbCr & vbCr
        If qErro > 0 Then
            msg = msg & "ATENÇÃO: " & qErro & " lugar(es) não aceitaram a troca. Confira na arte." & vbCr & vbCr
        End If
        msg = msg & "Ctrl+Z desfaz tudo de uma vez."
        MsgBox msg, IIf(qErro > 0, vbExclamation, vbInformation), TITULO
    End If
    Exit Sub

Falha:
    falhou = True
    MsgBox "A macro parou com erro: " & Err.Description & vbCr & _
           "Use Ctrl+Z para desfazer o que já foi feito.", vbCritical, TITULO
    Resume Concluir
End Sub

'---------------------------------------------------------------------
'  Cor que vai sair: do objeto selecionado ou escolhida na janela
'---------------------------------------------------------------------
Private Function tcPegarCorAlvo() As Boolean
    Dim s As Shape
    Dim cP As Color, cC As Color
    Dim n As Long, tipoFill As Long, tipoContorno As Long
    Dim resp As VbMsgBoxResult

    Set tcAlvo = New Color
    Set s = Nothing
    On Error Resume Next
    n = 0
    n = ActiveSelectionRange.Count
    If n > 0 Then Set s = ActiveSelectionRange(1)
    On Error GoTo 0

    If s Is Nothing Then
        MsgBox "Nada selecionado." & vbCr & vbCr & _
               "Na próxima janela, escolha a cor que vai SAIR, com o mesmo modelo " & _
               "e os mesmos valores usados na arte.", vbInformation, TITULO
        tcPegarCorAlvo = tcAlvo.UserAssignEx()
        Exit Function
    End If

    ' Cópias das cores do objeto (não mudam durante a troca)
    On Error Resume Next
    tipoFill = -1
    tipoFill = s.Fill.Type
    If tipoFill = cdrUniformFill Then
        Set cP = New Color
        Err.Clear
        cP.CopyAssign s.Fill.UniformColor
        If Err.Number <> 0 Then Set cP = Nothing
    End If
    tipoContorno = cdrNoOutline
    tipoContorno = s.Outline.Type
    If tipoContorno <> cdrNoOutline Then
        Set cC = New Color
        Err.Clear
        cC.CopyAssign s.Outline.Color
        If Err.Number <> 0 Then Set cC = Nothing
    End If
    On Error GoTo 0

    If cP Is Nothing And cC Is Nothing Then
        MsgBox "O objeto selecionado não tem preenchimento liso nem contorno " & _
               "(pode ser grupo, degradê ou imagem). Dentro de grupo, use Ctrl + clique." & vbCr & vbCr & _
               "Na próxima janela, escolha a cor que vai SAIR.", vbInformation, TITULO
        tcPegarCorAlvo = tcAlvo.UserAssignEx()
        Exit Function
    End If

    If Not cP Is Nothing And Not cC Is Nothing Then
        If Not cP.IsSame(cC) Then
            resp = MsgBox("Qual cor do objeto selecionado vai sair?" & vbCr & vbCr & _
                          "Sim = a do preenchimento: " & tcDescrever(cP) & vbCr & _
                          "Não = a do contorno: " & tcDescrever(cC), vbYesNoCancel + vbQuestion, TITULO)
            If resp = vbCancel Then Exit Function
            If resp = vbNo Then Set cP = Nothing
        End If
    End If
    If Not cP Is Nothing Then tcAlvo.CopyAssign cP Else tcAlvo.CopyAssign cC
    tcPegarCorAlvo = True
End Function

'---------------------------------------------------------------------
'  Páginas e camadas
'---------------------------------------------------------------------
Private Sub tcPercorrer(doc As Document, pgAtual As Page, ByVal todas As Boolean)
    Dim pg As Page
    If todas Then
        For Each pg In doc.Pages
            tcProcessarPagina pg
        Next pg
        tcProcessarPagina doc.MasterPage
    Else
        tcProcessarPagina pgAtual
    End If
End Sub

Private Sub tcProcessarPagina(pg As Page)
    Dim lyr As Layer
    Dim lista As ShapeRange
    Dim s As Shape
    Dim travada As Boolean

    If Not tcContar Then tcAtivar pg
    For Each lyr In pg.Layers
        If Not tcCamadaEspecial(lyr) Then
            travada = False
            If Not tcContar Then
                travada = Not tcEditavel(lyr)
                If travada Then tcDefinirEditavel lyr, True
            End If
            Set lista = tcFormas(lyr)
            If Not lista Is Nothing Then
                For Each s In lista
                    tcProcessar s
                Next s
            End If
            If travada Then tcDefinirEditavel lyr, False
        End If
    Next lyr
End Sub

Private Sub tcAtivar(pg As Page)
    On Error Resume Next
    pg.Activate
End Sub

' Camadas de guias, grade e área de trabalho ficam de fora
Private Function tcCamadaEspecial(lyr As Layer) As Boolean
    On Error Resume Next
    tcCamadaEspecial = lyr.IsDesktopLayer Or lyr.IsGuidesLayer Or lyr.IsGridLayer
End Function

Private Function tcEditavel(lyr As Layer) As Boolean
    tcEditavel = True
    On Error Resume Next
    tcEditavel = lyr.Editable
End Function

Private Sub tcDefinirEditavel(lyr As Layer, ByVal valor As Boolean)
    On Error Resume Next
    lyr.Editable = valor
End Sub

Private Function tcFormas(lyr As Layer) As ShapeRange
    On Error Resume Next
    Set tcFormas = lyr.Shapes.All
End Function

'---------------------------------------------------------------------
'  Um objeto e tudo o que está dentro dele (grupo, PowerClip, efeitos)
'---------------------------------------------------------------------
Private Sub tcProcessar(s As Shape)
    Dim tipo As Long, n As Long
    Dim travado As Boolean
    Dim f As Shape
    Dim pc As PowerClip

    On Error Resume Next
    tipo = -1
    tipo = s.Type
    If tipo = -1 Or tipo = cdrGuidelineShape Then Exit Sub

    travado = False
    If Not tcContar Then
        travado = s.Locked
        If travado Then s.Locked = False
    End If

    If tipo = cdrTextShape Then
        tcTexto s
    ElseIf tipo <> cdrGroupShape And tipo <> cdrBitmapShape Then
        If tcNoPreench Then tcPreenchimento s, False
        If tcNoContorno Then tcContorno s, False
    End If
    tcEfeitos s

    ' Grupos e efeitos com objetos dentro (mistura, contour, sombra...)
    n = 0
    If tipo <> cdrTextShape And tipo <> cdrBitmapShape Then n = s.Shapes.Count
    If n > 0 Then
        For Each f In s.Shapes
            tcProcessar f
        Next f
    End If

    ' Conteúdo do PowerClip
    Set pc = Nothing
    Set pc = s.PowerClip
    If Not pc Is Nothing Then
        For Each f In pc.Shapes
            tcProcessar f
        Next f
    End If

    If travado Then s.Locked = True
End Sub

' Preenchimento de um objeto ou de uma letra. True = a cor estava lá.
Private Function tcPreenchimento(ByVal alvo As Object, ByVal ehTexto As Boolean) As Boolean
    Dim fl As Object
    Dim c As Color
    Dim tipo As Long, i As Long, n As Long
    Dim achou As Boolean

    On Error Resume Next
    Set fl = Nothing
    Set fl = alvo.Fill
    If fl Is Nothing Then Exit Function
    tipo = -1
    tipo = fl.Type

    If tipo = cdrUniformFill Then
        Set c = Nothing
        Set c = fl.UniformColor
        If tcIgual(c) Then
            If tcAplicarPreench(fl) Then
                tcPreenchimento = True
                If Not ehTexto Then qPreench = qPreench + 1
            End If
        End If

    ElseIf tipo = cdrFountainFill Then
        n = -1
        n = fl.Fountain.Colors.Count
        For i = 0 To n + 1          ' cor inicial = 0, cor final = Count + 1
            Set c = Nothing
            Set c = fl.Fountain.Colors(i).Color
            If tcIgual(c) Then
                If tcAplicarCor(c) Then achou = True
            End If
        Next i
        If achou Then
            tcPreenchimento = True
            If Not ehTexto Then qDegrade = qDegrade + 1
        End If

    ElseIf tipo <> cdrNoFill And tipo <> -1 Then
        ' Padrão de duas cores
        Set c = Nothing
        Set c = fl.Pattern.FrontColor
        If tcIgual(c) Then
            If tcAplicarCor(c) Then achou = True
        End If
        Set c = Nothing
        Set c = fl.Pattern.BackColor
        If tcIgual(c) Then
            If tcAplicarCor(c) Then achou = True
        End If
        If achou Then
            tcPreenchimento = True
            If Not ehTexto Then qPreench = qPreench + 1
        End If
    End If
End Function

' Contorno de um objeto ou de uma letra. True = a cor estava lá.
Private Function tcContorno(ByVal alvo As Object, ByVal ehTexto As Boolean) As Boolean
    Dim ol As Object
    Dim c As Color
    Dim tipo As Long

    On Error Resume Next
    Set ol = Nothing
    Set ol = alvo.Outline
    If ol Is Nothing Then Exit Function
    tipo = cdrNoOutline
    tipo = ol.Type
    If tipo = cdrNoOutline Then Exit Function
    Set c = Nothing
    Set c = ol.Color
    If tcIgual(c) Then
        If tcAplicarContorno(ol, c) Then
            tcContorno = True
            If Not ehTexto Then qContorno = qContorno + 1
        End If
    End If
End Function

' Texto: confere letra por letra (palavras com cores diferentes)
Private Sub tcTexto(s As Shape)
    Dim st As TextRange, ch As TextRange
    Dim i As Long, n As Long
    Dim achou As Boolean

    On Error Resume Next
    Set st = Nothing
    Set st = s.Text.Story
    n = 0
    If Not st Is Nothing Then n = st.Characters.Count

    If n = 0 Then
        ' Sem acesso às letras: confere o texto inteiro
        If tcNoPreench Then
            If tcPreenchimento(s, True) Then achou = True
        End If
        If tcNoContorno Then
            If tcContorno(s, True) Then achou = True
        End If
    Else
        For i = 1 To n
            Set ch = Nothing
            Set ch = st.Characters(i)
            If Not ch Is Nothing Then
                If tcNoPreench Then
                    If tcPreenchimento(ch, True) Then achou = True
                End If
                If tcNoContorno Then
                    If tcContorno(ch, True) Then achou = True
                End If
            End If
        Next i
    End If
    If achou Then qTexto = qTexto + 1
End Sub

' Sombra e contour (efeito)
Private Sub tcEfeitos(s As Shape)
    Dim ef As Object
    Dim c As Color
    Dim n As Long
    Dim achou As Boolean

    On Error Resume Next
    n = 0
    n = s.Effects.Count
    If n = 0 Then Exit Sub
    For Each ef In s.Effects
        Set c = Nothing
        Set c = ef.DropShadow.Color
        If tcIgual(c) Then
            If tcAplicarCor(c) Then achou = True
        End If
        If tcNoPreench Then
            Set c = Nothing
            Set c = ef.Contour.FillColor
            If tcIgual(c) Then
                If tcAplicarCor(c) Then achou = True
            End If
        End If
        If tcNoContorno Then
            Set c = Nothing
            Set c = ef.Contour.OutlineColor
            If tcIgual(c) Then
                If tcAplicarCor(c) Then achou = True
            End If
        End If
    Next ef
    If achou Then qEfeito = qEfeito + 1
End Sub

'---------------------------------------------------------------------
'  Aplicar a cor nova (no modo de contar, só confirma)
'---------------------------------------------------------------------
Private Function tcAplicarPreench(fl As Object) As Boolean
    If tcContar Then tcAplicarPreench = True: Exit Function
    On Error GoTo Falhou
    fl.ApplyUniformFill tcNova
    tcAplicarPreench = True
    Exit Function
Falhou:
    qErro = qErro + 1
End Function

Private Function tcAplicarCor(c As Color) As Boolean
    If tcContar Then tcAplicarCor = True: Exit Function
    On Error GoTo Falhou
    c.CopyAssign tcNova
    tcAplicarCor = True
    Exit Function
Falhou:
    qErro = qErro + 1
End Function

Private Function tcAplicarContorno(ol As Object, c As Color) As Boolean
    If tcContar Then tcAplicarContorno = True: Exit Function
    On Error GoTo Falhou
    c.CopyAssign tcNova
    ' Se o Corel não aplicou direto na cor, aplica pelas propriedades do contorno
    If Not ol.Color.IsSame(tcNova) Then ol.SetProperties Color:=tcNova
    tcAplicarContorno = True
    Exit Function
Falhou:
    qErro = qErro + 1
End Function

'---------------------------------------------------------------------
'  Comparar e descrever cores
'---------------------------------------------------------------------
Private Function tcIgual(c As Color) As Boolean
    Dim a As Object, b As Object
    Dim r As Long

    On Error GoTo Nao
    If c Is Nothing Then Exit Function
    Set a = c
    Set b = tcAlvo
    If a.Type <> b.Type Then Exit Function
    Select Case a.Type
        Case cdrColorCMYK
            tcIgual = Abs(a.CMYKCyan - b.CMYKCyan) <= TOL_CMYK And _
                      Abs(a.CMYKMagenta - b.CMYKMagenta) <= TOL_CMYK And _
                      Abs(a.CMYKYellow - b.CMYKYellow) <= TOL_CMYK And _
                      Abs(a.CMYKBlack - b.CMYKBlack) <= TOL_CMYK
        Case cdrColorRGB
            tcIgual = Abs(a.RGBRed - b.RGBRed) <= TOL_RGB And _
                      Abs(a.RGBGreen - b.RGBGreen) <= TOL_RGB And _
                      Abs(a.RGBBlue - b.RGBBlue) <= TOL_RGB
        Case cdrColorPantone, 25            ' 25 = cor especial (spot)
            r = tcEspecial(a, b)
            If r = -1 Then tcIgual = c.IsSame(tcAlvo) Else tcIgual = (r = 1)
        Case Else
            tcIgual = c.IsSame(tcAlvo)
    End Select
    Exit Function
Nao:
    tcIgual = False
End Function

' Pantone/especial: mesmo nome e mesma porcentagem. -1 = não deu para comparar
Private Function tcEspecial(a As Object, b As Object) As Long
    On Error GoTo NaoDeu
    If Len(a.Name) = 0 Then
        tcEspecial = -1
    ElseIf a.Name <> b.Name Then
        tcEspecial = 0
    ElseIf Abs(a.Tint - b.Tint) <= 1 Then
        tcEspecial = 1
    Else
        tcEspecial = 0
    End If
    Exit Function
NaoDeu:
    tcEspecial = -1
End Function

Private Function tcDescrever(c As Color) As String
    Dim o As Object
    Dim t As String
    Dim tipo As Long

    On Error Resume Next
    Set o = c
    tipo = -1
    tipo = o.Type
    Select Case tipo
        Case cdrColorCMYK
            t = "CMYK " & o.CMYKCyan & "/" & o.CMYKMagenta & "/" & o.CMYKYellow & "/" & o.CMYKBlack
        Case cdrColorRGB
            t = "RGB " & o.RGBRed & "/" & o.RGBGreen & "/" & o.RGBBlue
        Case cdrColorGray
            t = "Cinza " & o.Gray
        Case cdrColorPantone, 25
            t = o.Name
            If o.Tint < 100 Then t = t & " " & o.Tint & "%"
        Case Else
            t = o.Name
    End Select
    If Len(t) = 0 Then t = o.Name
    If Len(t) = 0 Then t = "(cor escolhida)"
    tcDescrever = t
End Function

Private Function tcResumo() As String
    Dim t As String
    If qPreench > 0 Then t = t & vbCr & "   preenchimentos: " & qPreench
    If qDegrade > 0 Then t = t & vbCr & "   degradês: " & qDegrade
    If qContorno > 0 Then t = t & vbCr & "   contornos: " & qContorno
    If qTexto > 0 Then t = t & vbCr & "   textos: " & qTexto
    If qEfeito > 0 Then t = t & vbCr & "   sombras / contour: " & qEfeito
    tcResumo = t
End Function

Private Sub tcZerar()
    qPreench = 0: qDegrade = 0: qContorno = 0
    qTexto = 0: qEfeito = 0: qErro = 0
End Sub
