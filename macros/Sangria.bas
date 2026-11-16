Attribute VB_Name = "Sangria"
'=====================================================================
'  SANGRIA PELO CONTORNO (NAVALHA)  v13  -  macro para CorelDRAW (VBA)
'  by Bruno Faqui
'
'  COMO USAR: selecione o contorno externo da arte (a navalha/faca) OU
'  a própria moldura do PowerClip que tem o formato da faca, e rode
'  SangrarPeloContorno. Pode selecionar várias de uma vez.
'
'  v12: selecionando a moldura do PowerClip, ela mesma recebe a sangria
'  (antes ela era pulada e só as molduras de dentro cresciam).
'  v13: sangra todas as páginas (ou as escolhidas). Nas outras páginas a
'  macro usa a moldura (ou navalha) do mesmo tamanho e no mesmo lugar da
'  que está selecionada; se não achar, avisa a página no fim.
'
'  Para cada navalha selecionada a macro faz o mesmo que o processo
'  manual, sem deformar nada:
'    1. cria um contorno externo da navalha com a medida da sangria
'       (feito numa cópia: a navalha não é alterada);
'    2. separa o contorno (Ctrl+K);
'    3. solda o contorno na moldura do PowerClip que cobre a navalha.
'  A moldura passa a ter o formato da navalha + a sangria, certinho nos
'  alinhamentos da faca. O conteúdo do PowerClip não sai, não se move e
'  não estica: só aparece o que já sobrava para fora da moldura.
'
'  - Moldura "que cobre a navalha" = moldura de PowerClip (solta, em
'    grupo ou dentro de outro PowerClip) que ocupa a navalha inteira.
'  - Moldura retangular que já passa da linha de sangria não é mexida.
'  - Tira o contorno (linha) das molduras soldadas.
'  - Tudo vira um passo só de desfazer: um Ctrl+Z volta tudo.
'=====================================================================
Option Explicit

' Folga (mm) para considerar que a moldura cobre a navalha
Private Const FOLGA_MM As Double = 1
Private Const EPS As Double = 0.01

Private sang As Double
Private qtdSoldadas As Long, qtdJaTinham As Long, qtdErros As Long
Private qtdSemMoldura As Long, qtdAbertas As Long, qtdIgnoradas As Long
Private qtdNaoAchou As Long

' v13: o que foi selecionado (tamanho e lugar na página), para achar o mesmo nas outras páginas
Private Const TOL_IGUAL As Double = 0.5
Private modX() As Double, modY() As Double, modW() As Double, modH() As Double
Private modClip() As Boolean, modTipo() As Long, modNos() As Long
Private pgL As Double, pgB As Double
Private achado As Shape
Private buscaX As Double, buscaY As Double, buscaW As Double, buscaH As Double
Private buscaClip As Boolean, buscaTipo As Long, buscaNos As Long
Private relatorio As String

' Molduras de PowerClip da página ativa
Private molduras() As Shape
Private nMolduras As Long

Public Sub SangrarPeloContorno()
    Dim doc As Document
    Dim pgOriginal As Page, pg As Page
    Dim sel As ShapeRange
    Dim navalhas As New Collection
    Dim s As Shape, nav As Shape
    Dim resp As String, rotulo As String
    Dim unidadeAntiga As cdrUnit
    Dim i As Long, n As Long, nPaginas As Long
    Dim x As Double, y As Double, w As Double, h As Double
    Dim escolhidas() As Boolean, todas As Boolean
    Dim falhou As Boolean, grupoAberto As Boolean

    If Documents.Count = 0 Then
        MsgBox "Abra um documento antes de rodar a macro.", vbExclamation, "Sangria pelo contorno"
        Exit Sub
    End If
    Set doc = ActiveDocument
    Set pgOriginal = doc.ActivePage
    Set sel = ActiveSelectionRange
    If sel.Count = 0 Then
        MsgBox "Selecione a moldura do PowerClip com o formato da faca (ou a navalha) e rode a macro de novo." & vbCr & vbCr & _
               "Pode selecionar várias de uma vez.", vbExclamation, "Sangria pelo contorno"
        Exit Sub
    End If

    resp = InputBox("Sangria em milímetros:", "Sangria pelo contorno", "3")
    If Len(Trim$(resp)) = 0 Then Exit Sub
    sang = Val(Replace(resp, ",", "."))
    If sang <= 0 Or sang > 50 Then
        MsgBox "Valor de sangria inválido: " & resp, vbExclamation, "Sangria pelo contorno"
        Exit Sub
    End If

    ' v13: quais páginas (nas outras, procura a mesma moldura no mesmo lugar)
    If doc.Pages.Count > 1 Then
        resp = InputBox("Quais páginas sangrar?" & vbCr & vbCr & _
                        "todas  =  o documento inteiro" & vbCr & _
                        "atual  =  só a página " & pgOriginal.Index & vbCr & _
                        "ou uma lista, por exemplo:  1-3, 5, 8" & vbCr & vbCr & _
                        "Nas outras páginas a macro usa a moldura (ou navalha) do mesmo tamanho " & _
                        "e no mesmo lugar da que está selecionada.", "Sangria pelo contorno", "todas")
        If Len(Trim$(resp)) = 0 Then Exit Sub
        If Not LerPaginas(resp, doc.Pages.Count, pgOriginal.Index, escolhidas, todas) Then
            MsgBox "Não entendi as páginas: " & resp & vbCr & vbCr & _
                   "Use todas, atual ou uma lista como 1-3, 5, 8 (de 1 a " & doc.Pages.Count & ").", _
                   vbExclamation, "Sangria pelo contorno"
            Exit Sub
        End If
    Else
        ReDim escolhidas(1 To 1)
        escolhidas(1) = True
    End If

    qtdSoldadas = 0: qtdJaTinham = 0: qtdErros = 0
    qtdSemMoldura = 0: qtdAbertas = 0: qtdIgnoradas = 0: qtdNaoAchou = 0
    relatorio = "": nMolduras = 0

    unidadeAntiga = doc.Unit
    doc.Unit = cdrMillimeter

    For Each s In sel
        JuntarNavalhas s, navalhas
    Next s
    n = navalhas.Count
    If n = 0 Then
        doc.Unit = unidadeAntiga
        MsgBox "Nenhum contorno fechado na seleção." & vbCr & vbCr & _
               "Selecione a moldura ou a navalha (curva fechada, retângulo, elipse ou polígono).", _
               vbExclamation, "Sangria pelo contorno"
        Exit Sub
    End If

    ' Tamanho e lugar de cada uma (em relação à página), para achar a mesma nas outras páginas
    LerBordas pgOriginal
    ReDim modX(1 To n): ReDim modY(1 To n): ReDim modW(1 To n): ReDim modH(1 To n)
    ReDim modClip(1 To n): ReDim modTipo(1 To n): ReDim modNos(1 To n)
    For i = 1 To n
        Set nav = navalhas(i)
        nav.GetBoundingBox x, y, w, h
        modX(i) = x - pgL: modY(i) = y - pgB: modW(i) = w: modH(i) = h
        modClip(i) = TemPowerClip(nav)
        modTipo(i) = nav.Type
        modNos(i) = QuantosNos(nav)
    Next i

    On Error GoTo Falha
    doc.BeginCommandGroup "Sangria pelo contorno"
    grupoAberto = True
    Optimization = True

    For Each pg In doc.Pages
        If escolhidas(pg.Index) Then
            nPaginas = nPaginas + 1
            AtivarPagina pg               ' o Corel só mexe direito em PowerClip na página ativa
            LerBordas pg
            ListarMolduras pg
            For i = 1 To n
                rotulo = "Pág. " & pg.Index
                If n > 1 Then rotulo = rotulo & " (" & i & ")"
                If pg.Index = pgOriginal.Index Then
                    Set nav = navalhas(i)
                Else
                    Set nav = ProcurarIgual(pg, i)
                End If
                If nav Is Nothing Then
                    qtdNaoAchou = qtdNaoAchou + 1
                    If modClip(i) Then
                        relatorio = relatorio & "   " & rotulo & "  –  não achei a moldura igual à selecionada" & vbCr
                    Else
                        relatorio = relatorio & "   " & rotulo & "  –  não achei a navalha igual à selecionada" & vbCr
                    End If
                Else
                    SangrarNavalha nav, rotulo
                End If
            Next i
        End If
    Next pg

Concluir:
    On Error Resume Next
    pgOriginal.Activate
    Optimization = False
    If grupoAberto Then doc.EndCommandGroup
    doc.Unit = unidadeAntiga
    sel.CreateSelection
    ActiveWindow.Refresh
    Application.Refresh
    On Error GoTo 0

    If Not falhou Then MostrarResultado nPaginas, n
    Exit Sub

Falha:
    falhou = True
    MsgBox "A macro parou com erro: " & Err.Description & vbCr & _
           "Use Ctrl+Z para desfazer o que já foi feito.", vbCritical, "Sangria pelo contorno"
    Resume Concluir
End Sub

' Atalho com o nome antigo (para botões que já chamavam a macro anterior)
Public Sub SangrarTodasAsPaginas()
    SangrarPeloContorno
End Sub

'---------------------------------------------------------------------
'  Navalhas selecionadas (grupos são abertos)
'---------------------------------------------------------------------
Private Sub JuntarNavalhas(s As Shape, lista As Collection)
    Dim f As Shape
    On Error GoTo Fim
    Select Case s.Type
        Case cdrGroupShape
            ' grupo selecionado: cada parte vale como uma navalha
            For Each f In s.Shapes
                JuntarNavalhas f, lista
            Next f
        Case cdrCurveShape
            If s.Curve.Closed Then
                lista.Add s
            Else
                qtdAbertas = qtdAbertas + 1
            End If
        Case cdrRectangleShape, cdrEllipseShape, cdrPolygonShape
            lista.Add s
        Case Else
            qtdIgnoradas = qtdIgnoradas + 1
    End Select
Fim:
End Sub

'---------------------------------------------------------------------
'  Molduras de PowerClip da página (também em grupos e dentro de PowerClips)
'---------------------------------------------------------------------
Private Sub ListarMolduras(pg As Page)
    Dim lyr As Layer
    Dim s As Shape
    nMolduras = 0
    For Each lyr In pg.Layers
        If Not CamadaEspecial(lyr) Then
            If lyr.Visible And lyr.Printable Then
                For Each s In lyr.Shapes
                    ProcurarMolduras s
                Next s
            End If
        End If
    Next lyr
End Sub

Private Sub ProcurarMolduras(s As Shape)
    Dim f As Shape
    On Error GoTo Fim
    If s.Type = cdrGroupShape Then
        For Each f In s.Shapes
            ProcurarMolduras f
        Next f
        Exit Sub
    End If
    If TemPowerClip(s) Then
        nMolduras = nMolduras + 1
        ReDim Preserve molduras(1 To nMolduras)
        Set molduras(nMolduras) = s
        For Each f In s.PowerClip.Shapes
            ProcurarMolduras f
        Next f
    End If
Fim:
End Sub

Private Function CamadaEspecial(lyr As Layer) As Boolean
    On Error Resume Next
    CamadaEspecial = lyr.IsDesktopLayer Or lyr.IsGuidesLayer Or lyr.IsGridLayer
End Function

Private Function TemPowerClip(s As Shape) As Boolean
    On Error Resume Next
    TemPowerClip = Not (s.PowerClip Is Nothing)
End Function

'---------------------------------------------------------------------
'  Uma navalha: contorno externo + solda nas molduras que cobrem a navalha
'---------------------------------------------------------------------
Private Sub SangrarNavalha(nav As Shape, ByVal nome As String)
    Dim nx As Double, ny As Double, nw As Double, nh As Double
    Dim cx As Double, cy As Double, cw As Double, ch As Double
    Dim fx As Double, fy As Double, fw As Double, fh As Double
    Dim contorno As Shape
    Dim f As Shape
    Dim i As Long
    Dim achou As Boolean

    On Error GoTo Falhou
    nav.GetBoundingBox nx, ny, nw, nh

    Set contorno = CriarContorno(nav)
    If contorno Is Nothing Then
        qtdErros = qtdErros + 1
        relatorio = relatorio & "   " & nome & "  –  não deu para criar o contorno externo" & vbCr
        Exit Sub
    End If
    contorno.GetBoundingBox cx, cy, cw, ch

    For i = 1 To nMolduras
        Set f = molduras(i)
        ' a própria moldura selecionada também entra (v12)
        f.GetBoundingBox fx, fy, fw, fh
        If Cobre(fx, fy, fw, fh, nx, ny, nw, nh, FOLGA_MM) Then
            achou = True
            If f.Type = cdrRectangleShape And Cobre(fx, fy, fw, fh, cx, cy, cw, ch, EPS) Then
                qtdJaTinham = qtdJaTinham + 1          ' moldura retangular que já passa da sangria
            Else
                SoldarNaMoldura f, contorno, nome
            End If
        End If
    Next i

    If Not achou Then
        qtdSemMoldura = qtdSemMoldura + 1
        relatorio = relatorio & "   " & nome & "  –  nenhuma moldura de PowerClip cobre esta navalha" & vbCr
    End If

Fim:
    On Error Resume Next
    If Not contorno Is Nothing Then contorno.Delete
    Exit Sub

Falhou:
    qtdErros = qtdErros + 1
    relatorio = relatorio & "   " & nome & "  –  erro: " & Err.Description & vbCr
    Resume Fim
End Sub

' A caixa (x, y, w, h) da moldura cobre a caixa da navalha (com folga)
Private Function Cobre(ByVal fx As Double, ByVal fy As Double, ByVal fw As Double, ByVal fh As Double, _
                       ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, _
                       ByVal folga As Double) As Boolean
    Cobre = (fx <= x + folga) And (fy <= y + folga) And _
            (fx + fw >= x + w - folga) And (fy + fh >= y + h - folga)
End Function

' Contorno externo da navalha com a medida da sangria, já separado (Ctrl+K).
' Trabalha numa cópia: a navalha original não muda. Devolve uma curva só.
Private Function CriarContorno(nav As Shape) As Shape
    Dim copia As Shape, p As Shape, resultado As Shape
    Dim efeito As Effect
    Dim partes As ShapeRange, soltas As ShapeRange
    Dim idCopia As Long
    Dim k As Long

    On Error GoTo Falhou
    ' cópia só do formato, solta na camada (fora de grupo, sem efeitos)
    Set copia = nav.Layer.CreateCurve(nav.DisplayCurve)
    copia.Outline.SetNoOutline
    idCopia = copia.StaticID

    ' 1) contorno externo, 1 passo, distância = sangria
    Set efeito = copia.CreateContour(cdrContourOutside, sang, 1)
    ' 2) Ctrl+K: separa o contorno da cópia
    Set partes = efeito.Separate

    ' 3) junta as curvas do contorno numa só
    For Each p In partes
        If p.StaticID <> idCopia Then
            If p.Type = cdrGroupShape Then
                Set soltas = p.UngroupAllEx
                For k = 1 To soltas.Count
                    Set resultado = JuntarCurva(resultado, soltas(k))
                Next k
            Else
                Set resultado = JuntarCurva(resultado, p)
            End If
        End If
    Next p

    copia.Delete
    If Not resultado Is Nothing Then
        If resultado.Type <> cdrCurveShape Then resultado.ConvertToCurves
    End If
    Set CriarContorno = resultado
    Exit Function

Falhou:
    On Error Resume Next
    If Not copia Is Nothing Then copia.Delete
    If Not resultado Is Nothing Then resultado.Delete
    Set CriarContorno = Nothing
End Function

Private Function JuntarCurva(atual As Shape, nova As Shape) As Shape
    If atual Is Nothing Then
        Set JuntarCurva = nova
    Else
        Set JuntarCurva = nova.Weld(atual)
    End If
End Function

'---------------------------------------------------------------------
'  Solda o contorno na moldura, sem tirar o conteúdo do PowerClip
'---------------------------------------------------------------------
Private Sub SoldarNaMoldura(f As Shape, contorno As Shape, nome As String)
    Dim travado As Boolean, camadaTravada As Boolean
    Dim conteudoTravado As Boolean, mexeuNaTrava As Boolean
    Dim antes As Long
    Dim formato As Shape, copiaContorno As Shape, uniao As Shape
    Dim lyr As Layer
    Dim soldou As Boolean

    On Error GoTo Falhou
    Set lyr = f.Layer
    camadaTravada = Not lyr.Editable
    If camadaTravada Then lyr.Editable = True
    travado = f.Locked
    If travado Then f.Locked = False

    ' o conteúdo fica parado no lugar (não acompanha a moldura)
    If Not DestravarConteudo(f, conteudoTravado, mexeuNaTrava) Then
        Err.Raise vbObjectError + 513, , "o Corel não deixou destravar o conteúdo do PowerClip"
    End If
    antes = f.PowerClip.Shapes.Count

    ' formato atual da moldura numa curva à parte (sem o conteúdo) + contorno = solda
    Set formato = contorno.Layer.CreateCurve(f.DisplayCurve)
    Set copiaContorno = contorno.Duplicate
    Set uniao = copiaContorno.Weld(formato)
    soldou = True

    ' a moldura recebe o formato da solda; o conteúdo continua dentro dela
    If f.Type <> cdrCurveShape Then f.ConvertToCurves
    f.Curve.CopyAssign uniao.Curve
    uniao.Delete
    Set uniao = Nothing

    If f.PowerClip Is Nothing Then
        Err.Raise vbObjectError + 514, , "o PowerClip se perdeu (use Ctrl+Z)"
    ElseIf f.PowerClip.Shapes.Count <> antes Then
        Err.Raise vbObjectError + 515, , "o conteúdo do PowerClip mudou (use Ctrl+Z)"
    End If
    f.Outline.SetNoOutline
    qtdSoldadas = qtdSoldadas + 1

Fim:
    On Error Resume Next
    ' sobras: antes da solda, o formato e a cópia do contorno; depois, a união
    If Not soldou Then
        If Not formato Is Nothing Then formato.Delete
        If Not copiaContorno Is Nothing Then copiaContorno.Delete
    End If
    If Not uniao Is Nothing Then uniao.Delete
    If mexeuNaTrava Then f.PowerClip.ContentsLocked = conteudoTravado
    If travado Then f.Locked = True
    If camadaTravada Then lyr.Editable = False
    Exit Sub

Falhou:
    qtdErros = qtdErros + 1
    relatorio = relatorio & "   " & nome & "  –  moldura não soldada: " & Err.Description & vbCr
    Resume Fim
End Sub

' Destrava o conteúdo do PowerClip (assim ele não acompanha a moldura).
' Devolve False se o Corel não aceitar.
Private Function DestravarConteudo(s As Shape, ByRef estava As Boolean, ByRef mexeu As Boolean) As Boolean
    On Error GoTo Nao
    estava = s.PowerClip.ContentsLocked
    If estava Then
        mexeu = True
        s.PowerClip.ContentsLocked = False
        If s.PowerClip.ContentsLocked Then GoTo Nao
    End If
    DestravarConteudo = True
    Exit Function
Nao:
    DestravarConteudo = False
End Function

'---------------------------------------------------------------------
'  Resultado numa mensagem do próprio Corel
'---------------------------------------------------------------------
Private Sub MostrarResultado(ByVal nPaginas As Long, ByVal nNavalhas As Long)
    Dim msg As String
    Dim icone As VbMsgBoxStyle

    msg = "Sangria de " & sang & " mm pelo contorno em " & nPaginas & " página(s)" & vbCr & _
          "Molduras soldadas: " & qtdSoldadas & vbCr & vbCr

    If qtdJaTinham > 0 Then
        msg = msg & qtdJaTinham & " moldura(s) retangular(es) já passavam da linha de sangria (não mexi)." & vbCr
    End If
    If qtdAbertas > 0 Then
        msg = msg & qtdAbertas & " curva(s) aberta(s) na seleção foram ignoradas (a navalha precisa ser fechada)." & vbCr
    End If
    If qtdIgnoradas > 0 Then
        msg = msg & qtdIgnoradas & " objeto(s) da seleção não são contorno (texto, imagem...) e foram ignorados." & vbCr
    End If
    If Len(relatorio) > 0 Then
        msg = msg & vbCr & "CONFIRA:" & vbCr & relatorio
    End If
    If Len(msg) > 950 Then msg = Left$(msg, 920) & vbCr & "   … (lista cortada)" & vbCr
    msg = msg & vbCr & "Ctrl+Z desfaz tudo de uma vez."

    icone = vbInformation
    If qtdErros > 0 Or qtdSemMoldura > 0 Or qtdNaoAchou > 0 Or qtdSoldadas = 0 Then icone = vbExclamation
    MsgBox msg, icone, "Sangria pelo contorno"
End Sub

'---------------------------------------------------------------------
'  Outras páginas: a mesma moldura (ou navalha) no mesmo lugar
'---------------------------------------------------------------------
Private Sub AtivarPagina(pg As Page)
    On Error Resume Next
    pg.Activate
End Sub

Private Sub LerBordas(pg As Page)
    On Error GoTo SemPropriedades
    pgL = pg.LeftX: pgB = pg.BottomY
    Exit Sub
SemPropriedades:
    pgL = 0: pgB = 0
End Sub

Private Function QuantosNos(s As Shape) As Long
    QuantosNos = -1
    On Error Resume Next
    If s.Type = cdrCurveShape Then QuantosNos = s.Curve.Nodes.Count
End Function

Private Function ProcurarIgual(pg As Page, ByVal i As Long) As Shape
    Dim lyr As Layer
    Dim s As Shape
    buscaX = modX(i) + pgL: buscaY = modY(i) + pgB: buscaW = modW(i): buscaH = modH(i)
    buscaClip = modClip(i): buscaTipo = modTipo(i): buscaNos = modNos(i)
    Set achado = Nothing
    For Each lyr In pg.Layers
        If Not CamadaEspecial(lyr) Then
            If lyr.Visible And lyr.Printable Then
                For Each s In lyr.Shapes
                    ProcurarIgualEm s
                    If Not achado Is Nothing Then Exit For
                Next s
            End If
        End If
        If Not achado Is Nothing Then Exit For
    Next lyr
    Set ProcurarIgual = achado
End Function

' Procura em grupos e dentro de PowerClips; fica com a primeira (a de fora)
Private Sub ProcurarIgualEm(s As Shape)
    Dim f As Shape
    Dim x As Double, y As Double, w As Double, h As Double
    Dim clip As Boolean

    On Error GoTo Fim
    If Not achado Is Nothing Then Exit Sub
    If s.Type = cdrGroupShape Then
        For Each f In s.Shapes
            ProcurarIgualEm f
            If Not achado Is Nothing Then Exit Sub
        Next f
        Exit Sub
    End If

    clip = TemPowerClip(s)
    If clip = buscaClip And s.Type = buscaTipo Then
        If QuantosNos(s) = buscaNos Then
            s.GetBoundingBox x, y, w, h
            If Abs(x - buscaX) <= TOL_IGUAL And Abs(y - buscaY) <= TOL_IGUAL And _
               Abs(w - buscaW) <= TOL_IGUAL And Abs(h - buscaH) <= TOL_IGUAL Then
                Set achado = s
                Exit Sub
            End If
        End If
    End If
    If clip Then
        For Each f In s.PowerClip.Shapes
            ProcurarIgualEm f
            If Not achado Is Nothing Then Exit Sub
        Next f
    End If
Fim:
End Sub

'---------------------------------------------------------------------
'  Escolha das páginas
'---------------------------------------------------------------------

' Lê "todas", "atual" ou uma lista como "1-3, 5, 8" (também aceita "1 a 3" e "2 e 5").
' Devolve False se não entender o texto.
Private Function LerPaginas(ByVal texto As String, ByVal total As Long, ByVal atual As Long, _
                            escolhidas() As Boolean, ByRef todas As Boolean) As Boolean
    Dim partes() As String
    Dim t As String, item As String
    Dim i As Long, a As Long, b As Long, traco As Long
    Dim algum As Boolean

    ReDim escolhidas(1 To total)
    todas = False
    t = LCase$(Trim$(texto))

    If t = "todas" Or t = "todos" Or t = "tudo" Then
        For i = 1 To total: escolhidas(i) = True: Next i
        todas = True
        LerPaginas = True
        Exit Function
    End If
    If t = "atual" Or t = "esta" Then
        escolhidas(atual) = True
        todas = (total = 1)
        LerPaginas = True
        Exit Function
    End If

    t = Replace(t, " a ", "-")
    t = Replace(t, " e ", ",")
    t = Replace(t, ";", ",")
    Do While InStr(t, " -") > 0: t = Replace(t, " -", "-"): Loop
    Do While InStr(t, "- ") > 0: t = Replace(t, "- ", "-"): Loop
    t = Replace(t, " ", ",")
    partes = Split(t, ",")

    For i = 0 To UBound(partes)
        item = Trim$(partes(i))
        If Len(item) > 0 Then
            traco = InStr(item, "-")
            If traco > 0 Then
                If Not Inteiro(Left$(item, traco - 1), a) Then Exit Function
                If Not Inteiro(Mid$(item, traco + 1), b) Then Exit Function
            Else
                If Not Inteiro(item, a) Then Exit Function
                b = a
            End If
            If a > b Then traco = a: a = b: b = traco
            If a < 1 Or b > total Then Exit Function
            For traco = a To b: escolhidas(traco) = True: Next traco
            algum = True
        End If
    Next i

    If algum Then
        todas = True
        For i = 1 To total
            If Not escolhidas(i) Then todas = False
        Next i
    End If
    LerPaginas = algum
End Function

Private Function Inteiro(ByVal s As String, ByRef n As Long) As Boolean
    Dim i As Long
    s = Trim$(s)
    If Len(s) = 0 Or Len(s) > 6 Then Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next i
    n = CLng(s)
    Inteiro = True
End Function
