// SeparacoesCorel - painel WPF hospedado pelo CorelDRAW.
// Adaptado de Separacoes(1).cs. Compilar como DLL, nunca como programa EXE.
// O Corel entrega sua propria Application ao construtor DockerUI(object app).
// Nao usa GetActiveObject, CreateInstance, SetParent nem Application.Run.
// A previa CMYK/Pantone e raster e aproximada; nao e uma prova de separacoes de impressao.
using System;
using System.Collections.Generic;

#if !MOTOR_APENAS
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using PathIO = System.IO.Path;
using ShapePath = System.Windows.Shapes.Path;
#endif

namespace SeparacoesCorel
{
    // Motor independente de Corel e de WPF, tambem usado nos testes automaticos.
    internal sealed class ImagemCmyk
    {
        internal const int MaxPixels = 18000000; // v2.3 (antes 4 milhoes)
        internal readonly int Largura, Altura;
        internal readonly byte[] Pixels;
        internal readonly TintaEspecial[] Especiais;

        internal ImagemCmyk(int largura, int altura, byte[] pixels)
            : this(largura, altura, pixels, new TintaEspecial[0]) { }

        internal ImagemCmyk(int largura, int altura, byte[] pixels, TintaEspecial[] especiais)
        {
            long tamanho = (long)largura * altura;
            if (largura < 1 || altura < 1 || tamanho > MaxPixels ||
                pixels == null || pixels.LongLength != tamanho * 4)
                throw new ArgumentException("Dimensões ou dados CMYK inválidos. Limite: 18 milhões de pixels.");
            Largura = largura;
            Altura = altura;
            Pixels = pixels;
            Especiais = especiais ?? new TintaEspecial[0];
            if (Especiais.Length > 60 || tamanho * (4 + Especiais.Length) > 320L * 1024 * 1024)
                throw new ArgumentException("Muitas chapas ou prévia muito grande. Reduza a resolução.");
            foreach (TintaEspecial tinta in Especiais)
                if (tinta == null || tinta.Cobertura == null || tinta.Cobertura.Length != tamanho)
                    throw new ArgumentException("Dados da chapa especial inválidos.");
            // v3.7: sem a conta da carga maxima (a linha saiu do painel)
        }

        internal static int Percentual(int valor) { return (int)Math.Round(valor * 100.0 / 255.0); }

        internal int[] Amostra(int x, int y)
        {
            if (x < 0 || y < 0 || x >= Largura || y >= Altura) return null;
            int i = (y * Largura + x) * 4;
            return new int[] { Percentual(Pixels[i]), Percentual(Pixels[i + 1]),
                Percentual(Pixels[i + 2]), Percentual(Pixels[i + 3]),
                Percentual(Pixels[i] + Pixels[i + 1] + Pixels[i + 2] + Pixels[i + 3]) };
        }
    }

    internal sealed class TintaEspecial
    {
        internal string Nome;
        internal byte[] Cobertura;
        internal byte R, G, B;
    }

    internal static class Renderizador
    {
        internal static byte[] Desenhar(ImagemCmyk img, bool[] visivel, bool emPreto,
                                        bool destacar, int limitePct)
        { return Desenhar(img, visivel, emPreto, destacar, limitePct, null); }

        internal static byte[] Desenhar(ImagemCmyk img, bool[] visivel, bool emPreto,
                                        bool destacar, int limitePct, bool[] especiaisVisiveis)
        {
            if (img == null || visivel == null || visivel.Length != 4)
                throw new ArgumentException("Informe a imagem e as quatro chapas.");
            if (limitePct < 0 || limitePct > 400) throw new ArgumentOutOfRangeException("limitePct");
            if (especiaisVisiveis != null && especiaisVisiveis.Length != img.Especiais.Length)
                throw new ArgumentException("A visibilidade das tintas especiais não corresponde às chapas.");
            byte[] saida = new byte[img.Pixels.Length];
            byte[] p = img.Pixels;
            for (int i = 0; i < p.Length; i += 4)
            {
                if (destacar && (p[i] + p[i + 1] + p[i + 2] + p[i + 3]) * 100 > limitePct * 255)
                {
                    saida[i] = 0; saida[i + 1] = 230; saida[i + 2] = 0;
                }
                else
                {
                    int c = visivel[0] ? p[i] : 0, m = visivel[1] ? p[i + 1] : 0;
                    int y = visivel[2] ? p[i + 2] : 0, k = visivel[3] ? p[i + 3] : 0;
                    if (emPreto)
                    {
                        int cinza = (int)Math.Round(255.0 * (255 - c) / 255 * (255 - m) / 255 *
                                                  (255 - y) / 255 * (255 - k) / 255);
                        saida[i] = saida[i + 1] = saida[i + 2] = (byte)cinza;
                    }
                    else
                    {
                        // BGRA. Visualizacao indicativa, sem simular papel/ICC na tela.
                        saida[i] = (byte)((255 - y) * (255 - k) / 255);
                        saida[i + 1] = (byte)((255 - m) * (255 - k) / 255);
                        saida[i + 2] = (byte)((255 - c) * (255 - k) / 255);
                    }
                    // Os Pantones sao chapas independentes. A cor RGB serve somente
                    // para representar a tinta na tela; nao entra nos canais CMYK.
                    for (int spot = 0; spot < img.Especiais.Length; spot++)
                    {
                        if (especiaisVisiveis != null && !especiaisVisiveis[spot]) continue;
                        TintaEspecial tinta = img.Especiais[spot];
                        int cobertura = tinta.Cobertura[i / 4];
                        int r = emPreto ? 0 : tinta.R, g = emPreto ? 0 : tinta.G, b = emPreto ? 0 : tinta.B;
                        saida[i] = (byte)(saida[i] * (65025 - cobertura * (255 - b)) / 65025);
                        saida[i + 1] = (byte)(saida[i + 1] * (65025 - cobertura * (255 - g)) / 65025);
                        saida[i + 2] = (byte)(saida[i + 2] * (65025 - cobertura * (255 - r)) / 65025);
                    }
                }
                saida[i + 3] = 255;
            }
            return saida;
        }

        internal static bool LocalizarPixel(double mx, double my, double areaW, double areaH,
                                            int w, int h, out int x, out int y)
        {
            x = y = -1;
            if (areaW <= 0 || areaH <= 0 || w <= 0 || h <= 0) return false;
            double escala = Math.Min(areaW / w, areaH / h);
            x = (int)Math.Floor((mx - (areaW - w * escala) / 2) / escala);
            y = (int)Math.Floor((my - (areaH - h * escala) / 2) / escala);
            return x >= 0 && y >= 0 && x < w && y < h;
        }
    }

#if !MOTOR_APENAS
    // v2.6: cores na tela pelo perfil CMYK (como no Illustrator), em vez da conta simples RGB = 1 - CMY.
    // Se o Windows nao converter, volta para a conta simples.
    internal static class CorDeTela
    {
        static bool indisponivel;

        internal static string LocalizarPerfil(string motor)
        {
            try
            {
                if (!String.IsNullOrEmpty(motor))
                {
                    string pastaBin = PathIO.GetDirectoryName(motor);
                    string raizGs = String.IsNullOrEmpty(pastaBin) ? null : PathIO.GetDirectoryName(pastaBin);
                    if (!String.IsNullOrEmpty(raizGs))
                    {
                        string perfil = PathIO.Combine(raizGs, "iccprofiles", "default_cmyk.icc");
                        if (File.Exists(perfil)) return perfil;
                    }
                }
                string windows = PathIO.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System),
                                                "spool", "drivers", "color", "RSWOP.icm");
                if (File.Exists(windows)) return windows;
            }
            catch { }
            return null;
        }

        internal static byte[] Desenhar(ImagemCmyk img, bool[] visivel, bool emPreto, bool destacar, int limitePct,
                                        bool[] especiaisVisiveis, string perfil)
        {
            if (emPreto || perfil == null || indisponivel)
                return Renderizador.Desenhar(img, visivel, emPreto, destacar, limitePct, especiaisVisiveis);
            try
            {
                int w = img.Largura, h = img.Altura;
                byte[] p = img.Pixels;
                byte[] cmyk = new byte[p.Length];
                bool vc = visivel[0], vm = visivel[1], vy = visivel[2], vk = visivel[3];
                for (int i = 0; i < p.Length; i += 4)
                {
                    cmyk[i] = vc ? p[i] : (byte)0;
                    cmyk[i + 1] = vm ? p[i + 1] : (byte)0;
                    cmyk[i + 2] = vy ? p[i + 2] : (byte)0;
                    cmyk[i + 3] = vk ? p[i + 3] : (byte)0;
                }
                BitmapSource origem = BitmapSource.Create(w, h, 96, 96, PixelFormats.Cmyk32, null, cmyk, w * 4);
                ColorConvertedBitmap convertido = new ColorConvertedBitmap(origem, new ColorContext(new Uri(perfil)),
                    new ColorContext(PixelFormats.Bgra32), PixelFormats.Bgra32);
                byte[] saida = new byte[checked(w * h * 4)];
                convertido.CopyPixels(saida, w * 4, 0);
                for (int i = 0; i < saida.Length; i += 4)
                {
                    if (destacar && (p[i] + p[i + 1] + p[i + 2] + p[i + 3]) * 100 > limitePct * 255)
                    {
                        saida[i] = 0; saida[i + 1] = 230; saida[i + 2] = 0; saida[i + 3] = 255;
                        continue;
                    }
                    for (int spot = 0; spot < img.Especiais.Length; spot++)
                    {
                        if (especiaisVisiveis != null && !especiaisVisiveis[spot]) continue;
                        TintaEspecial tinta = img.Especiais[spot];
                        int cobertura = tinta.Cobertura[i / 4];
                        saida[i] = (byte)(saida[i] * (65025 - cobertura * (255 - tinta.B)) / 65025);
                        saida[i + 1] = (byte)(saida[i + 1] * (65025 - cobertura * (255 - tinta.G)) / 65025);
                        saida[i + 2] = (byte)(saida[i + 2] * (65025 - cobertura * (255 - tinta.R)) / 65025);
                    }
                    saida[i + 3] = 255;
                }
                return saida;
            }
            catch (Exception ex)
            {
                indisponivel = true;
                Corel.Registrar(ex);
                return Renderizador.Desenhar(img, visivel, emPreto, destacar, limitePct, especiaisVisiveis);
            }
        }
    }

    internal sealed class Captura
    {
        internal string Arquivo, Chave, Descricao, Diretorio, Motor;
        internal int Dpi;
        internal int Fator = 1; // v2.2: suavizacao (renderiza Fator vezes maior e reduz)
        internal int DpiMaximo; // v2.3: maior resolucao permitida para esta pagina
        internal string PastaPdf; // v2.4: pasta do PDF, guardada para o detalhe da area visivel
        internal double LarguraIn, AlturaIn; // v2.4: tamanho da pagina em polegadas
    }

    internal sealed class VistaDocumento
    {
        internal IntPtr Janela, JanelaAplicacao;
        internal int Esquerda, Topo, Direita, Base;
    }

    internal sealed class Corel
    {
        // Referencia emprestada do host: nunca liberar/fechar a Application do Corel.
        readonly object aplicacao;
        internal object Aplicacao { get { return aplicacao; } } // v2.9

        internal Corel(object app)
        {
            if (app == null) throw new ArgumentNullException("app");
            aplicacao = app;
        }

        internal object Documento()
        {
            dynamic app = aplicacao;
            if ((int)app.Documents.Count == 0) return null;
            return (object)app.ActiveDocument;
        }

        internal string ChaveAtual()
        {
            object doc = Documento();
            if (doc == null) return null;
            IntPtr ptr = Marshal.GetIUnknownForObject(doc);
            try { return ptr.ToInt64().ToString("X") + ":" + ((dynamic)doc).ActivePage.Index; }
            finally { Marshal.Release(ptr); } // equilibra apenas o AddRef acima.
        }

        internal bool LerSobreimpressao()
        {
            if (Documento() == null) throw new InvalidOperationException("Abra um documento.");
            return (bool)((dynamic)aplicacao).ActiveWindow.ActiveView.SimulateOverprints;
        }

        internal void DefinirSobreimpressao(bool ativar)
        {
            if (Documento() == null) throw new InvalidOperationException("Abra um documento.");
            ((dynamic)aplicacao).ActiveWindow.ActiveView.SimulateOverprints = ativar;
        }

        // v2.3: pixels de tela por polegada em que a pagina aparece agora (zoom atual do Corel).
        internal double DpiDaTela()
        {
            dynamic app = aplicacao;
            object objDoc = Documento();
            if (objDoc == null) return 0;
            dynamic doc = objDoc;
            double w = (double)app.ConvertUnits((double)doc.ActivePage.SizeWidth, doc.Unit, 1);
            if (double.IsNaN(w) || double.IsInfinity(w) || w <= 0) return 0;
            VistaDocumento vista = LerVista();
            return Math.Abs(vista.Direita - vista.Esquerda) / w;
        }

        internal VistaDocumento LerVista()
        {
            dynamic app = aplicacao;
            dynamic doc = Documento();
            if (doc == null) throw new InvalidOperationException("Abra um documento.");
            dynamic janela = app.ActiveWindow;
            dynamic pagina = doc.ActivePage;
            int x1 = 0, y1 = 0, x2 = 0, y2 = 0;
            janela.DocumentToScreen((double)pagina.LeftX, (double)pagina.TopY, out x1, out y1);
            janela.DocumentToScreen((double)pagina.RightX, (double)pagina.BottomY, out x2, out y2);
            return new VistaDocumento { Janela = Identificador(janela.Handle),
                JanelaAplicacao = Identificador(app.AppWindow.Handle),
                Esquerda = x1, Topo = y1, Direita = x2, Base = y2 };
        }

        static IntPtr Identificador(object valor)
        {
            long n = Convert.ToInt64(valor);
            if (IntPtr.Size == 8 && n < 0 && n >= Int32.MinValue) n = unchecked((uint)(int)n);
            return new IntPtr(n);
        }

        internal Captura Exportar(int dpiSolicitado)
        {
            string motor = MotorSeparacoes.LocalizarExecutavel();
            dynamic app = aplicacao;
            object objDoc = Documento();
            if (objDoc == null) throw new InvalidOperationException("Abra uma arte no CorelDRAW.");
            dynamic doc = objDoc;
            // Unidade documentada do CorelDRAW: polegadas=1.
            // Nao depende de uma DLL Corel.Interop de versao especifica.
            double w = (double)app.ConvertUnits((double)doc.ActivePage.SizeWidth, doc.Unit, 1);
            double h = (double)app.ConvertUnits((double)doc.ActivePage.SizeHeight, doc.Unit, 1);
            if (double.IsNaN(w) || double.IsInfinity(w) || double.IsNaN(h) || double.IsInfinity(h) || w <= 0 || h <= 0)
                throw new InvalidOperationException("Não foi possível determinar o tamanho da página.");
            // Margem para arredondamento do exportador; evita grandes alocacoes no host.
            // v2.3: ate 16 milhoes de pixels por pagina (antes 3,5 milhoes), lado maximo 8000 px.
            double dpiMax = Math.Min(Math.Sqrt(16000000.0 / (w * h)), 8000 / Math.Max(w, h));
            if (dpiMax < 1) throw new InvalidOperationException("Página grande demais para esta prévia.");
            int dpi = Math.Max(1, Math.Min(dpiSolicitado, (int)Math.Floor(dpiMax)));
            // v2.2: renderiza ate cerca de 400 dpi e reduz para a resolucao da previa,
            // com no maximo ~40 milhoes de pixels por chapa durante a renderizacao.
            int fator = Math.Max(1, Math.Min(4, (int)Math.Ceiling(400.0 / dpi)));
            while (fator > 1 && (w * dpi * fator) * (h * dpi * fator) > 48000000.0) fator--;
            string diretorio = PathIO.Combine(PathIO.GetTempPath(), "Imprell_Separacoes_" + Guid.NewGuid().ToString("N"));
            // v2.4: o PDF fica numa pasta propria, guardada ate a proxima atualizacao.
            string pastaPdf = PathIO.Combine(PathIO.GetTempPath(), "Imprell_Separacoes_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(diretorio);
            Directory.CreateDirectory(pastaPdf);
            string arquivo = PathIO.Combine(pastaPdf, "pagina.pdf");
            try
            {
                string chave = ChaveAtual();
                ExportacaoPdf.Exportar(objDoc, arquivo);
                if (!File.Exists(arquivo)) throw new IOException("O Corel não gerou o PDF temporário.");
                string descricao = "Página " + doc.ActivePage.Index + " · " + dpi + " dpi";
                if (dpi != dpiSolicitado) descricao += " (reduzido para a prévia)";
                if (fator > 1) descricao += " · suavizada";
                return new Captura { Arquivo = arquivo, Chave = chave, Descricao = descricao,
                    Diretorio = diretorio, Motor = motor, Dpi = dpi, Fator = fator, DpiMaximo = (int)Math.Floor(dpiMax),
                    PastaPdf = pastaPdf, LarguraIn = w, AlturaIn = h };
            }
            catch { MotorSeparacoes.Limpar(diretorio); MotorSeparacoes.Limpar(pastaPdf); throw; }
        }

        internal static void ApagarTemporario(string arquivo)
        {
            if (String.IsNullOrEmpty(arquivo)) return;
            try { File.Delete(arquivo); } catch (Exception ex) { Registrar(ex); }
        }

        internal static void Registrar(Exception ex)
        {
            try
            {
                string dir = PathIO.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Imprell", "SeparacoesCorel");
                Directory.CreateDirectory(dir);
                string log = PathIO.Combine(dir, "erros.log");
                if (File.Exists(log) && new FileInfo(log).Length > 1024 * 1024) File.Delete(log);
                File.AppendAllText(log, DateTime.Now.ToString("s") + " " + ex + Environment.NewLine);
            }
            catch { } // O log nunca deve derrubar o processo do Corel.
        }

        internal static string Explicar(Exception ex)
        {
            AggregateException agregado = ex as AggregateException;
            if (agregado != null) ex = agregado.Flatten().InnerException;
            COMException com = ex as COMException;
            if (com != null && (com.ErrorCode == unchecked((int)0x80010001) || com.ErrorCode == unchecked((int)0x8001010A)))
                return "O Corel está ocupado. Feche a caixa de diálogo e clique em Atualizar.";
            return ex.Message;
        }
    }

    // v2.9: confere a pagina atual do Corel procurando cores fora do CMYK: RGB (e HSB, HLS, Lab, YIQ),
    // imagens RGB/Lab/em paleta e cor de registro. So le o arquivo; nao altera nada.
    internal sealed class ItemRgb
    {
        internal object Objeto;   // objeto que tem a cor
        internal object Topo;     // objeto da camada que contem o item
        internal int IndiceTopo;
        internal List<object> Molduras = new List<object>(); // PowerClips por fora do item (de fora para dentro)
        internal bool Tela, Registro, Imagem, Preto;
        internal string Descricao;
    }

    internal sealed class ResultadoRgb
    {
        internal readonly List<ItemRgb> Itens = new List<ItemRgb>();
        internal int Objetos, Imagens, Registro, Visitados, Topos, Preto;
        internal bool Parcial, Completa;
        internal string Chave;
    }

    internal static class VerificadorRgb
    {
        // v3.6: a lista guarda ate 3000 itens; "Conferir a página toda" tira o limite de 5 s
        const int LimiteObjetos = 40000, LimiteObjetosCompleta = 400000, LimitePalavras = 1500;
        internal const int LimiteItens = 3000;
        static readonly TimeSpan LimiteTempo = TimeSpan.FromSeconds(5), LimiteTempoCompleta = TimeSpan.FromSeconds(120);
        // Tipos de forma que contem outras: grupo, mistura, extrusao, contorno, chanfro, sombra, midia artistica, efeito.
        static readonly int[] Grupos = { 7, 10, 11, 13, 15, 16, 18, 22 };
        static readonly Dictionary<int, string> modelos = new Dictionary<int, string>(); // tipo de cor -> modelo
        static readonly HashSet<int> modosVistos = new HashSet<int>();
        static int textoCor = -1; // Color.ToString do Corel: -1 nao testado, 0 nao funciona, 1 funciona
        static int falhas;
        static List<object> moldurasAbertas = new List<object>();

        sealed class Visita
        {
            internal ResultadoRgb Resultado;
            internal System.Diagnostics.Stopwatch Relogio;
            internal bool Parar;
            internal int MaxObjetos;     // v3.6
            internal TimeSpan MaxTempo;  // v3.6
        }

        sealed class Achado
        {
            internal readonly List<string> Partes = new List<string>();
            internal bool Tela, Registro, Imagem, Preto;
        }

        internal static ResultadoRgb Verificar(object documento) { return Verificar(documento, false); }

        // v3.6: completa = sem o limite de 5 s (botao "Conferir a página toda")
        internal static ResultadoRgb Verificar(object documento, bool completa)
        {
            if (documento == null) throw new InvalidOperationException("Abra uma arte no CorelDRAW.");
            falhas = 0;
            dynamic doc = documento;
            dynamic pagina = doc.ActivePage;
            double pl = (double)pagina.LeftX, pr = (double)pagina.RightX, pt = (double)pagina.TopY, pb = (double)pagina.BottomY;
            ResultadoRgb r = new ResultadoRgb { Completa = completa };
            Visita v = new Visita { Resultado = r, Relogio = System.Diagnostics.Stopwatch.StartNew(),
                MaxObjetos = completa ? LimiteObjetosCompleta : LimiteObjetos, MaxTempo = completa ? LimiteTempoCompleta : LimiteTempo };
            List<object> camadas = new List<object>();
            AdicionarCamadas(camadas, (object)pagina.Layers, false);
            try { AdicionarCamadas(camadas, (object)doc.MasterPage.Layers, true); } catch (Exception ex) { Falha(ex); }
            foreach (object camada in camadas)
            {
                if (v.Parar) break;
                try
                {
                    dynamic formas = ((dynamic)camada).Shapes;
                    int n = (int)formas.Count;
                    for (int i = 1; i <= n && !v.Parar; i++)
                    {
                        object s = formas.Item(i);
                        if (!NaPagina(s, pl, pr, pt, pb)) continue;
                        r.Topos++;
                        Visitar(s, s, r.Topos, new List<object>(), false, v);
                    }
                }
                catch (Exception ex) { Falha(ex); }
            }
            r.Parcial = v.Parar;
            MotorRegiao.Diagnostico("RGB" + (completa ? " (completa)" : "") + ": " + r.Visitados + " objetos em " + v.Relogio.ElapsedMilliseconds + " ms; " +
                r.Objetos + " objetos RGB, " + r.Imagens + " imagens RGB, " + r.Registro + " em registro" + (r.Parcial ? " (parcial)" : ""));
            return r;
        }

        static void AdicionarCamadas(List<object> lista, object colecao, bool mestre)
        {
            dynamic camadas = colecao;
            int n = (int)camadas.Count;
            for (int i = 1; i <= n; i++)
            {
                object c = camadas.Item(i);
                if (Logico(c, "IsDesktopLayer", false) || Logico(c, "IsGuidesLayer", false) || Logico(c, "IsGridLayer", false)) continue;
                if (!mestre && Logico(c, "Master", false)) continue; // camada mestre: entra pela pagina mestre
                if (!Logico(c, "Visible", true) || !Logico(c, "Printable", true)) continue;
                lista.Add(c);
            }
        }

        static bool Logico(object alvo, string nome, bool padrao)
        {
            try
            {
                return Convert.ToBoolean(alvo.GetType().InvokeMember(nome, System.Reflection.BindingFlags.GetProperty, null, alvo, null,
                    System.Globalization.CultureInfo.InvariantCulture));
            }
            catch { return padrao; }
        }

        static bool NaPagina(object forma, double pl, double pr, double pt, double pb)
        {
            try
            {
                dynamic s = forma;
                double l = (double)s.LeftX, rx = (double)s.RightX, t = (double)s.TopY, b = (double)s.BottomY;
                return rx >= pl && l <= pr && t >= pb && b <= pt;
            }
            catch { return true; }
        }

        static void Visitar(object forma, object topo, int indiceTopo, List<object> molduras, bool emGrupo, Visita v)
        {
            if (v.Parar) return;
            ResultadoRgb r = v.Resultado;
            if (++r.Visitados > v.MaxObjetos || v.Relogio.Elapsed > v.MaxTempo) { v.Parar = true; return; }
            dynamic s = forma;
            int tipo;
            try { tipo = (int)s.Type; } catch (Exception ex) { Falha(ex); return; }
            if (tipo == 9) return; // guia
            bool grupo = Array.IndexOf(Grupos, tipo) >= 0;
            Achado a = new Achado();
            // forma (object) nas chamadas: ligacao estatica dos metodos, so as propriedades do Corel sao dinamicas
            if (tipo == 5) ConferirImagem(forma, a);
            else if (!grupo)
            {
                ConferirPreenchimento(forma, a);
                ConferirContorno(forma, a);
                if (tipo == 6 && !a.Tela) ConferirTexto(forma, a);
            }
            ConferirEfeitos(forma, a);
            if (a.Partes.Count > 0) Anotar(forma, topo, indiceTopo, molduras, emGrupo, tipo, a, r);

            if (grupo)
            {
                try
                {
                    dynamic filhos = s.Shapes;
                    int n = (int)filhos.Count;
                    for (int i = 1; i <= n && !v.Parar; i++) Visitar((object)filhos.Item(i), topo, indiceTopo, molduras, true, v);
                }
                catch (Exception ex) { Falha(ex); }
            }
            object pc = null;
            try { pc = s.PowerClip; } catch { }
            if (pc != null)
            {
                try
                {
                    List<object> dentro = new List<object>(molduras);
                    dentro.Add(forma);
                    dynamic conteudo = ((dynamic)pc).Shapes;
                    int n = (int)conteudo.Count;
                    for (int i = 1; i <= n && !v.Parar; i++) Visitar((object)conteudo.Item(i), topo, indiceTopo, dentro, emGrupo, v);
                }
                catch (Exception ex) { Falha(ex); }
            }
        }

        static void Anotar(object forma, object topo, int indiceTopo, List<object> molduras, bool emGrupo, int tipo, Achado a, ResultadoRgb r)
        {
            if (a.Tela) { if (a.Imagem) r.Imagens++; else r.Objetos++; }
            if (a.Registro) r.Registro++;
            if (a.Preto) r.Preto++; // v3.2
            if (r.Itens.Count >= LimiteItens) return;
            string onde = molduras.Count > 0 ? " · dentro de PowerClip" : emGrupo ? " · em grupo" : "";
            r.Itens.Add(new ItemRgb { Objeto = forma, Topo = topo, IndiceTopo = indiceTopo, Molduras = molduras,
                Tela = a.Tela, Registro = a.Registro, Imagem = a.Imagem, Preto = a.Preto,
                Descricao = NomeDoTipo(tipo) + " · " + String.Join(", ", a.Partes.ToArray()) + onde });
        }

        static void ConferirPreenchimento(dynamic s, Achado a)
        {
            try
            {
                dynamic f = s.Fill;
                int tipo = (int)f.Type;
                if (tipo == 1) Cor((object)f.UniformColor, "preenchimento", a);
                else if (tipo == 2)
                {
                    // Degrade: cor inicial no indice 0, final em Count + 1.
                    dynamic cores = f.Fountain.Colors;
                    int n = (int)cores.Count;
                    for (int i = 0; i <= n + 1; i++)
                        if (Cor((object)cores.Item(i).Color, "degradê", a)) break;
                }
            }
            catch (Exception ex) { Falha(ex); }
        }

        static void ConferirContorno(dynamic s, Achado a)
        {
            try
            {
                dynamic o = s.Outline;
                if ((double)o.Width <= 0) return;
                Cor((object)o.Color, "contorno", a);
            }
            catch (Exception ex) { Falha(ex); }
        }

        // Texto com palavras de cores diferentes: confere palavra por palavra.
        static void ConferirTexto(dynamic s, Achado a)
        {
            try
            {
                dynamic palavras = s.Text.Story.Words;
                int n = Math.Min((int)palavras.Count, LimitePalavras);
                for (int i = 1; i <= n; i++)
                {
                    dynamic f = palavras.Item(i).Fill;
                    if ((int)f.Type == 1 && Cor((object)f.UniformColor, "texto", a)) return;
                }
            }
            catch (Exception ex) { Falha(ex); }
        }

        static void ConferirImagem(dynamic s, Achado a)
        {
            try
            {
                int modo = (int)s.Bitmap.Mode;
                if (modosVistos.Add(modo)) MotorRegiao.Diagnostico("RGB: imagem com modo " + modo);
                string nome = (modo == 4 || modo == 9) ? "RGB" : modo == 7 ? "Lab" : (modo == 1 || modo == 3) ? "em paleta (RGB)" : null;
                if (nome != null) { a.Tela = true; a.Imagem = true; a.Partes.Add(nome); DetalharImagem((object)s, a); }
            }
            catch (Exception ex) { Falha(ex); }
        }

        // v3.6: tamanho em pixels e nome do arquivo (quando vinculada), para reconhecer a imagem na lista
        static void DetalharImagem(object forma, Achado a)
        {
            try
            {
                dynamic b = ((dynamic)forma).Bitmap;
                int w = (int)b.SizeWidth, h = (int)b.SizeHeight;
                if (w > 0 && h > 0) a.Partes.Add(w + "×" + h + " px");
            }
            catch { }
            try
            {
                dynamic b = ((dynamic)forma).Bitmap;
                if ((bool)b.ExternallyLinked)
                {
                    string arquivo = (string)b.LinkFileName;
                    if (!String.IsNullOrEmpty(arquivo)) a.Partes.Add("vínculo " + PathIO.GetFileName(arquivo));
                }
            }
            catch { }
        }

        static void ConferirEfeitos(dynamic s, Achado a)
        {
            try
            {
                dynamic efeitos = s.Effects;
                int n = (int)efeitos.Count;
                for (int i = 1; i <= n; i++)
                {
                    try
                    {
                        object sombra = efeitos.Item(i).DropShadow;
                        if (sombra != null) Cor((object)((dynamic)sombra).Color, "sombra", a);
                    }
                    catch { } // efeito que nao e sombra
                }
            }
            catch (Exception ex) { Falha(ex); }
        }

        // Retorna true se a cor e de tela (RGB e afins).
        static bool Cor(object cor, string parte, Achado a)
        {
            if (cor == null) return false;
            string modelo = Modelo(cor);
            if (modelo == "registro")
            {
                if (!a.Registro) { a.Registro = true; a.Partes.Add(parte + " em cor de registro"); }
                return false;
            }
            if (String.IsNullOrEmpty(modelo)) { ConferirPreto(cor, parte, a); return false; } // v3.2
            a.Tela = true;
            string valores = "";
            if (modelo == "RGB")
            {
                try { dynamic c = cor; valores = " " + (int)c.RGBRed + "/" + (int)c.RGBGreen + "/" + (int)c.RGBBlue; } catch { }
            }
            a.Partes.Add(parte + " " + modelo + valores);
            return true;
        }

        // v3.3: preto e cinza que nao sao so K.
        //  - preto com mistura: K 90 a 100 com C, M ou Y (ex.: C20 M5 Y10 K100)
        //  - preto nao 100%: K 90 a 99 sozinho (ex.: K95)
        //  - cinza com mistura: K 1 a 89 com um pouco de C, M ou Y (ex.: M5 Y2 K50).
        //    Conta como cinza quando C, M e Y sao todos ate 10%, ou quase iguais entre si
        //    (diferenca ate 15%); cor de verdade (ex.: M80 Y90 K20) nao entra.
        //  - cinza sem K: C, M e Y quase iguais (diferenca ate 8%, todos a partir de 5%) e K 0
        //    (ex.: C20 M15 Y15), um cinza que podia ser so K.
        static readonly Dictionary<int, bool> tiposCmyk = new Dictionary<int, bool>();

        static void ConferirPreto(object cor, string parte, Achado a)
        {
            if (a.Preto) return;
            try
            {
                dynamic c = cor;
                int tipo = (int)c.Type;
                if (!EhCmyk(cor, tipo)) return;
                int k = (int)c.CMYKBlack;
                int ci = (int)c.CMYKCyan, m = (int)c.CMYKMagenta, y = (int)c.CMYKYellow;
                int maior = Math.Max(ci, Math.Max(m, y)), menor = Math.Min(ci, Math.Min(m, y));
                string valores = "C" + ci + " M" + m + " Y" + y + " K" + k;
                string motivo = null;
                if (ci + m + y == 0)
                {
                    if (k >= 90 && k < 100) motivo = "preto em " + k + "%, não 100%";
                }
                else if (k >= 90) motivo = "preto com mistura";
                else if (k > 0 && (maior <= 10 || maior - menor <= 15)) motivo = "cinza com mistura, podia ser só K";
                else if (k == 0 && menor >= 5 && maior - menor <= 8) motivo = "cinza feito com C, M e Y, sem K";
                if (motivo == null) return;
                a.Preto = true;
                a.Partes.Add(parte + " " + valores + " (" + motivo + ")");
            }
            catch (Exception ex) { Falha(ex); }
        }

        static bool EhCmyk(object cor, int tipo)
        {
            bool sim;
            if (tiposCmyk.TryGetValue(tipo, out sim)) return sim;
            sim = tipo == 2;
            string texto = TextoDaCor(cor);
            if (!String.IsNullOrEmpty(texto))
            {
                string t = texto.TrimStart().ToUpperInvariant();
                if (t.StartsWith("CMYK", StringComparison.Ordinal)) sim = true;
                else if (PorTexto(texto) != null) sim = false; // outro modelo conhecido
            }
            tiposCmyk[tipo] = sim;
            MotorRegiao.Diagnostico("Conferencia: cor tipo " + tipo + (sim ? " e CMYK" : " nao e CMYK"));
            return sim;
        }

        // Modelo de cor pelo tipo do Corel. Na primeira vez que aparece um tipo, confere tambem pelo
        // texto da cor (quando o Corel fornece) e anota em diagnostico.log.
        static string Modelo(object cor)
        {
            int tipo = (int)((dynamic)cor).Type;
            string rotulo;
            if (modelos.TryGetValue(tipo, out rotulo)) return rotulo;
            rotulo = PorNumero(tipo);
            string texto = TextoDaCor(cor);
            string pelo = PorTexto(texto);
            if (pelo != null) rotulo = pelo;
            modelos[tipo] = rotulo;
            MotorRegiao.Diagnostico("RGB: cor tipo " + tipo + " = " + (rotulo.Length == 0 ? "CMYK/ok" : rotulo) + (texto != null ? " · " + texto : ""));
            return rotulo;
        }

        static string PorNumero(int tipo)
        {
            switch (tipo)
            {
                case 5: return "RGB";
                case 6: return "HSB";
                case 7: return "HLS";
                case 11: return "YIQ";
                case 12: return "Lab";
                case 20: return "registro";
                default: return "";
            }
        }

        static string PorTexto(string texto)
        {
            if (String.IsNullOrEmpty(texto)) return null;
            string t = texto.TrimStart().ToUpperInvariant();
            if (t.StartsWith("RGB", StringComparison.Ordinal)) return "RGB";
            if (t.StartsWith("HSB", StringComparison.Ordinal)) return "HSB";
            if (t.StartsWith("HLS", StringComparison.Ordinal)) return "HLS";
            if (t.StartsWith("LAB", StringComparison.Ordinal)) return "Lab";
            if (t.StartsWith("YIQ", StringComparison.Ordinal)) return "YIQ";
            if (t.StartsWith("REGISTRATION", StringComparison.Ordinal)) return "registro";
            foreach (string ok in new string[] { "CMY", "GRAY", "GREY", "BW", "BLACK", "PANTONE", "SPOT", "USERINK", "MULTICHANNEL" })
                if (t.StartsWith(ok, StringComparison.Ordinal)) return "";
            return null;
        }

        static string TextoDaCor(object cor)
        {
            if (textoCor == 0) return null;
            try
            {
                string s = cor.GetType().InvokeMember("ToString", System.Reflection.BindingFlags.InvokeMethod, null, cor, null,
                    System.Globalization.CultureInfo.InvariantCulture) as string;
                if (String.IsNullOrEmpty(s) || s.StartsWith("System.", StringComparison.Ordinal)) { textoCor = 0; return null; }
                textoCor = 1; return s;
            }
            catch { textoCor = 0; return null; }
        }

        static string NomeDoTipo(int tipo)
        {
            switch (tipo)
            {
                case 1: return "Retângulo";
                case 2: return "Elipse";
                case 3: return "Curva";
                case 4: return "Polígono";
                case 5: return "Imagem";
                case 6: return "Texto";
                case 7: return "Grupo";
                default: return "Objeto";
            }
        }

        static void Falha(Exception ex) { if (falhas++ < 20) Corel.Registrar(ex); }

        // Sai dos PowerClips que o painel abriu para mostrar um item.
        static void SairDasMolduras()
        {
            for (int i = moldurasAbertas.Count - 1; i >= 0; i--)
                try { ((dynamic)moldurasAbertas[i]).PowerClip.LeaveEditMode(); } catch { }
            moldurasAbertas = new List<object>();
        }

        static void AjustarZoom(object aplicacao)
        {
            try { ((dynamic)aplicacao).ActiveWindow.ActiveView.ToFitSelection(); } catch (Exception ex) { Falha(ex); }
        }

        internal static string SelecionarTodos(object aplicacao, object documento, List<ItemRgb> itens)
        {
            SairDasMolduras();
            dynamic doc = documento;
            try { doc.ClearSelection(); } catch (Exception ex) { Falha(ex); }
            HashSet<int> feitos = new HashSet<int>();
            int ok = 0;
            foreach (ItemRgb item in itens)
            {
                if (!feitos.Add(item.IndiceTopo)) continue;
                try { ((dynamic)item.Topo).AddToSelection(); ok++; } catch (Exception ex) { Falha(ex); }
            }
            if (ok == 0) return "Não foi possível selecionar. A camada pode estar travada ou a arte mudou: clique em Atualizar página.";
            AjustarZoom(aplicacao);
            return ok + (ok == 1 ? " objeto selecionado" : " objetos selecionados") + " no Corel. Quando o RGB está dentro de grupo ou PowerClip, " +
                "fica selecionado o objeto de fora; clique no item da lista para ir direto nele.";
        }

        internal static string SelecionarItem(object aplicacao, object documento, ItemRgb item)
        {
            SairDasMolduras();
            dynamic doc = documento;
            try { doc.ClearSelection(); } catch (Exception ex) { Falha(ex); }
            List<object> abertas = new List<object>();
            try
            {
                foreach (object m in item.Molduras) { ((dynamic)m).PowerClip.EnterEditMode(); abertas.Add(m); }
                ((dynamic)item.Objeto).CreateSelection();
                moldurasAbertas = abertas;
                AjustarZoom(aplicacao);
                return item.Molduras.Count > 0
                    ? "Selecionado dentro do PowerClip. Quando terminar, clique em Concluir no Corel antes de atualizar o painel."
                    : "Selecionado no Corel.";
            }
            catch (Exception ex)
            {
                Falha(ex);
                for (int i = abertas.Count - 1; i >= 0; i--)
                    try { ((dynamic)abertas[i]).PowerClip.LeaveEditMode(); } catch { }
                try
                {
                    ((dynamic)item.Topo).CreateSelection();
                    AjustarZoom(aplicacao);
                    return "Selecionado o objeto de fora (não deu para entrar no grupo/PowerClip).";
                }
                catch (Exception ex2)
                {
                    Falha(ex2);
                    return "Não foi possível selecionar. O objeto pode ter sido apagado ou a camada está travada: clique em Atualizar página.";
                }
            }
        }
    }

    // Visualizacao somente em memoria: uma janela FILHA do documento, sem aba,
    // sem janela de programa, sem inserir um bitmap/objeto no arquivo CDR.
    // Este modo nao implementa o renderer de separacoes do Illustrator.
    internal sealed class PreviaNaArte : IDisposable
    {
        readonly HwndSource fonte;
        readonly Grid raiz;
        readonly Canvas folha;
        readonly Image imagem;
        readonly IntPtr pai;
        readonly Action sair;
        readonly Action atualizar;
        readonly Action aoZoom; // v2.3
        readonly Action<double, double> aoMedir; // v2.5
        // v2.4: camada com a area visivel na resolucao da tela, por cima da pagina inteira.
        // v2.6: fica fora do zoom, em pixels exatos da tela (Stretch None e DPI da tela).
        readonly Image detalheImg = new Image { Stretch = Stretch.None, SnapsToDevicePixels = true, Visibility = Visibility.Collapsed,
            HorizontalAlignment = HorizontalAlignment.Left, VerticalAlignment = VerticalAlignment.Top, IsHitTestVisible = false };
        int versaoVista; // v2.6: muda a cada zoom, arraste ou mudanca de tamanho
        Matrix matrizDetalhe = Matrix.Identity; // v2.8: zoom/posicao da previa quando o detalhe exibido foi pedido
        readonly MatrixTransform transformacao = new MatrixTransform();
        Point arraste;
        bool arrastando;
        int larguraJanela, alturaJanela;
        bool encerrada;

        internal PreviaNaArte(VistaDocumento vista, Action aoSair, Action aoAtualizar, Action aoZoomTela,
                              Action<double, double> aoMedirTela)
        {
            pai = vista.Janela; sair = aoSair; atualizar = aoAtualizar; aoZoom = aoZoomTela; aoMedir = aoMedirTela;
            uint processo;
            Native.GetWindowThreadProcessId(pai, out processo);
            if (pai == IntPtr.Zero || pai == vista.JanelaAplicacao || !Native.IsWindow(pai) ||
                processo != (uint)System.Diagnostics.Process.GetCurrentProcess().Id ||
                (Native.GetWindowLong(pai, -16) & 0x40000000) == 0)
                throw new InvalidOperationException("Não foi possível localizar com segurança a área do documento. A prévia permanece no painel.");
            Native.RECT ret;
            if (!Native.GetClientRect(pai, out ret) || ret.Right < 80 || ret.Bottom < 80)
                throw new InvalidOperationException("A janela do documento não está disponível para a prévia.");
            larguraJanela = ret.Right; alturaJanela = ret.Bottom;
            HwndSourceParameters config = new HwndSourceParameters("SeparacoesCorel_PreviaNaArte");
            config.ParentWindow = pai;
            config.WindowStyle = unchecked((int)0x50000000); // WS_CHILD | WS_VISIBLE
            config.ExtendedWindowStyle = 0; // nao e top-level nem taskbar
            config.PositionX = 0; config.PositionY = 0;
            config.Width = larguraJanela; config.Height = alturaJanela;
            fonte = new HwndSource(config);
            try
            {
                raiz = new Grid { Background = new SolidColorBrush(Color.FromRgb(82, 82, 82)),
                    ClipToBounds = true, Focusable = true };
                folha = new Canvas { ClipToBounds = false, RenderTransform = transformacao };
                imagem = new Image { Stretch = Stretch.Fill, SnapsToDevicePixels = true };
                // v2.2: reducao/ampliacao suave (sem o efeito "craquelado" do filtro padrao).
                RenderOptions.SetBitmapScalingMode(imagem, BitmapScalingMode.HighQuality);
                folha.Children.Add(imagem); raiz.Children.Add(folha); raiz.Children.Add(detalheImg);
                RenderOptions.SetBitmapScalingMode(detalheImg, BitmapScalingMode.NearestNeighbor); // 1:1, sem filtro
                DockPanel barra = new DockPanel { LastChildFill = false };
                Button fechar = new Button { Content = "Voltar à edição  ·  Esc", Padding = new Thickness(12, 5, 12, 5),
                    Margin = new Thickness(8, 6, 8, 6) };
                fechar.Click += delegate { SairDepoisDoEvento(); };
                DockPanel.SetDock(fechar, Dock.Right); barra.Children.Add(fechar);
                TextBlock aviso = new TextBlock { Text = "PRÉVIA CMYK + ESPECIAIS  ·  roda: zoom  ·  arraste: mover  ·  F5: atualizar",
                    Foreground = Brushes.White, FontFamily = new FontFamily("Segoe UI"), FontSize = 12,
                    VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(12, 5, 8, 5),
                    TextWrapping = TextWrapping.Wrap };
                barra.Children.Add(aviso);
                Border faixa = new Border { Background = new SolidColorBrush(Color.FromRgb(43, 43, 43)),
                    Child = barra, VerticalAlignment = VerticalAlignment.Top };
                raiz.Children.Add(faixa);
                raiz.PreviewKeyDown += delegate(object s, KeyEventArgs e)
                {
                    if (e.Key == Key.Escape) SairDepoisDoEvento();
                    else if (e.Key == Key.F5) raiz.Dispatcher.BeginInvoke(atualizar);
                    e.Handled = true; // a imagem e de consulta, nao um editor de objetos.
                };
                raiz.PreviewMouseWheel += delegate(object s, MouseWheelEventArgs e)
                {
                    Matrix m = transformacao.Matrix;
                    double fator = e.Delta > 0 ? 1.2 : 1.0 / 1.2;
                    if (m.M11 * fator < 0.05 || m.M11 * fator > 32) { e.Handled = true; return; } // v2.8: no limite, a roda nao vai para o Corel
                    Point p = e.GetPosition(raiz);
                    m.ScaleAt(fator, fator, p.X, p.Y); transformacao.Matrix = m; e.Handled = true; MudouVista();
                    if (aoZoom != null) aoZoom(); // v2.3: depois do zoom, gera de novo mais nitida
                };
                folha.MouseLeftButtonDown += delegate(object s, MouseButtonEventArgs e)
                {
                    arraste = e.GetPosition(raiz); arrastando = true; folha.CaptureMouse();
                    raiz.Focus(); e.Handled = true;
                };
                // O fundo tambem pode iniciar o deslocamento.
                raiz.MouseLeftButtonDown += delegate(object s, MouseButtonEventArgs e)
                {
                    if (e.Handled) return;
                    arraste = e.GetPosition(raiz); arrastando = true; folha.CaptureMouse(); raiz.Focus();
                };
                folha.MouseMove += delegate(object s, MouseEventArgs e)
                {
                    if (!arrastando) return;
                    Point p = e.GetPosition(raiz); Matrix m = transformacao.Matrix;
                    m.Translate(p.X - arraste.X, p.Y - arraste.Y); transformacao.Matrix = m; arraste = p; MudouVista();
                };
                folha.MouseLeftButtonUp += delegate
                {
                    arrastando = false; folha.ReleaseMouseCapture();
                    if (aoZoom != null) aoZoom(); // v2.4: depois de arrastar, gera o detalhe da nova area
                };
                folha.LostMouseCapture += delegate { arrastando = false; };
                // v2.5: leitura das tintas onde o mouse passa sobre a arte (0 a 1 na pagina)
                raiz.MouseMove += delegate(object s, MouseEventArgs e)
                {
                    if (aoMedir == null || arrastando) return;
                    double esquerda = Canvas.GetLeft(imagem), topo = Canvas.GetTop(imagem);
                    if (double.IsNaN(esquerda) || double.IsNaN(topo) || imagem.Width <= 0 || imagem.Height <= 0) return;
                    Point ponto = e.GetPosition(folha);
                    aoMedir((ponto.X - esquerda) / imagem.Width, (ponto.Y - topo) / imagem.Height);
                };
                raiz.MouseLeave += delegate { if (aoMedir != null) aoMedir(-1, -1); };
                fonte.RootVisual = raiz;
                PosicionarPagina(vista);
                Native.SetWindowPos(fonte.Handle, IntPtr.Zero, 0, 0, larguraJanela, alturaJanela, 0x0010 | 0x0040);
            }
            catch { fonte.Dispose(); throw; }
        }

        void PosicionarPagina(VistaDocumento vista)
        {
            Native.POINT a = new Native.POINT { X = vista.Esquerda, Y = vista.Topo };
            Native.POINT b = new Native.POINT { X = vista.Direita, Y = vista.Base };
            if (!Native.ScreenToClient(pai, ref a) || !Native.ScreenToClient(pai, ref b) || b.X <= a.X || b.Y <= a.Y)
                throw new InvalidOperationException("Não foi possível alinhar a prévia à página. Use uma vista sem rotação e tente novamente.");
            Matrix dpi = fonte.CompositionTarget.TransformFromDevice;
            Point inicio = dpi.Transform(new Point(a.X, a.Y));
            Point fim = dpi.Transform(new Point(b.X, b.Y));
            imagem.Width = fim.X - inicio.X; imagem.Height = fim.Y - inicio.Y;
            Canvas.SetLeft(imagem, inicio.X); Canvas.SetTop(imagem, inicio.Y);
        }

        // v2.3: zoom feito com a roda do mouse sobre a previa (1 = pagina do tamanho do Corel).
        internal double Escala { get { return transformacao.Matrix.M11; } }

        // v2.6: o detalhe so aparece se foi feito para a vista atual (zoom/posicao).
        internal int VersaoVista { get { return versaoVista; } }

        // v2.8: zoom/posicao atuais da previa, guardados junto com cada pedido de detalhe.
        internal Matrix MatrizVista { get { return transformacao.Matrix; } }

        void MudouVista()
        {
            versaoVista++;
            // v2.8: o detalhe nitido nao some mais no zoom/arraste; ele acompanha a nova vista
            // ate o detalhe da nova area ficar pronto (como no Acrobat e no Illustrator).
            AjustarDetalhe();
        }

        // v2.8: posiciona o detalhe exibido na vista atual. Mesma vista do pedido: 1:1, sem filtro.
        // Vista diferente: o mesmo detalhe ampliado/reduzido/deslocado junto com a pagina.
        void AjustarDetalhe()
        {
            if (encerrada || detalheImg.Source == null) return;
            Matrix atual = transformacao.Matrix;
            if (atual == matrizDetalhe)
            {
                detalheImg.RenderTransform = Transform.Identity;
                RenderOptions.SetBitmapScalingMode(detalheImg, BitmapScalingMode.NearestNeighbor); // 1:1, sem filtro
            }
            else
            {
                Matrix inversa = matrizDetalhe;
                if (!inversa.HasInverse) { OcultarDetalhe(); return; }
                inversa.Invert();
                Thickness o = detalheImg.Margin;
                Matrix r = Matrix.Identity;
                r.Translate(o.Left, o.Top);   // ponto do detalhe -> ponto na janela (vista do pedido)
                r.Append(inversa);            // -> ponto na pagina
                r.Append(atual);              // -> ponto na janela (vista atual)
                r.Translate(-o.Left, -o.Top); // -> coordenadas do proprio detalhe
                detalheImg.RenderTransform = new MatrixTransform(r);
                RenderOptions.SetBitmapScalingMode(detalheImg, BitmapScalingMode.Linear);
            }
            detalheImg.Visibility = Visibility.Visible;
        }

        // v2.6: area visivel da pagina alinhada aos pixels da tela.
        // u/v: fracao da pagina (0 a 1, a partir do canto superior esquerdo); px: posicao e tamanho em
        // pixels da tela; dpiTela: pixels da tela por polegada da pagina; escala: pixels por unidade WPF.
        internal bool RegiaoTela(double larguraPaginaIn, out double u0, out double v0, out double u1, out double v1,
                                 out int px0, out int py0, out int pxLargura, out int pxAltura, out double dpiTela, out double escala)
        {
            u0 = v0 = u1 = v1 = dpiTela = 0; px0 = py0 = pxLargura = pxAltura = 0; escala = 1;
            if (encerrada || raiz.ActualWidth <= 0 || raiz.ActualHeight <= 0 || larguraPaginaIn <= 0) return false;
            double L = Canvas.GetLeft(imagem), T = Canvas.GetTop(imagem), W = imagem.Width, H = imagem.Height;
            if (double.IsNaN(L) || double.IsNaN(T) || double.IsNaN(W) || double.IsNaN(H) || W <= 0 || H <= 0) return false;
            Matrix m = transformacao.Matrix;
            Matrix dispositivo = fonte.CompositionTarget.TransformToDevice;
            double sx = dispositivo.M11, sy = dispositivo.M22;
            if (sx <= 0 || sy <= 0) return false;
            Point a = m.Transform(new Point(L, T)), b = m.Transform(new Point(L + W, T + H));
            double ax = Math.Min(a.X, b.X) * sx, bx = Math.Max(a.X, b.X) * sx;
            double ay = Math.Min(a.Y, b.Y) * sy, by = Math.Max(a.Y, b.Y) * sy;
            if (bx - ax < 1 || by - ay < 1) return false;
            int x0 = (int)Math.Floor(Math.Max(0, ax)), x1 = (int)Math.Ceiling(Math.Min(raiz.ActualWidth * sx, bx));
            int y0 = (int)Math.Floor(Math.Max(0, ay)), y1 = (int)Math.Ceiling(Math.Min(raiz.ActualHeight * sy, by));
            if (x1 - x0 < 2 || y1 - y0 < 2) return false;
            u0 = (x0 - ax) / (bx - ax); u1 = (x1 - ax) / (bx - ax);
            v0 = (y0 - ay) / (by - ay); v1 = (y1 - ay) / (by - ay);
            px0 = x0; py0 = y0; pxLargura = x1 - x0; pxAltura = y1 - y0;
            dpiTela = (bx - ax) / larguraPaginaIn;
            escala = sx;
            return true;
        }

        // v2.8: matriz = zoom/posicao da previa quando o detalhe foi pedido. O detalhe anterior so e
        // trocado quando o novo esta pronto, sem passar pela pagina inteira desfocada.
        internal void ExibirDetalhe(BitmapSource bitmap, int px0, int py0, double escala, Matrix matriz)
        {
            if (encerrada) return;
            if (bitmap == null || escala <= 0) { OcultarDetalhe(); return; }
            detalheImg.Margin = new Thickness(px0 / escala, py0 / escala, 0, 0);
            detalheImg.Source = bitmap; matrizDetalhe = matriz;
            AjustarDetalhe();
        }

        internal void OcultarDetalhe()
        {
            if (encerrada) return;
            detalheImg.Source = null; detalheImg.Visibility = Visibility.Collapsed;
            detalheImg.RenderTransform = Transform.Identity;
        }

        internal void Exibir(BitmapSource bitmap)
        {
            if (encerrada) return;
            imagem.Source = bitmap;
        }

        // v2.8: ao atualizar sem fechar a previa, confere se a pagina mudou de lugar ou de tamanho no Corel.
        // Retorna true se mudou (o detalhe antigo deixa de valer).
        internal bool Alinhar(VistaDocumento vista)
        {
            if (encerrada) return false;
            double l = Canvas.GetLeft(imagem), t = Canvas.GetTop(imagem), w = imagem.Width, h = imagem.Height;
            PosicionarPagina(vista);
            bool igual = Math.Abs(Canvas.GetLeft(imagem) - l) < 0.01 && Math.Abs(Canvas.GetTop(imagem) - t) < 0.01 &&
                         Math.Abs(imagem.Width - w) < 0.01 && Math.Abs(imagem.Height - h) < 0.01;
            if (!igual) { OcultarDetalhe(); MudouVista(); }
            return !igual;
        }

        // Resize da janela filha, sem alterar o zoom ou o documento do Corel.
        internal bool Conferir(VistaDocumento vista)
        {
            if (encerrada || !Native.IsWindow(fonte.Handle) || vista.Janela != pai || !Native.IsWindowVisible(pai)) return false;
            Native.RECT ret;
            if (!Native.GetClientRect(pai, out ret) || ret.Right < 80 || ret.Bottom < 80) return false;
            if (ret.Right != larguraJanela || ret.Bottom != alturaJanela)
            {
                larguraJanela = ret.Right; alturaJanela = ret.Bottom;
                Native.SetWindowPos(fonte.Handle, IntPtr.Zero, 0, 0, larguraJanela, alturaJanela, 0x0010);
                MudouVista();
                if (aoZoom != null) aoZoom(); // v2.8
            }
            return true;
        }

        void SairDepoisDoEvento()
        {
            if (!encerrada) raiz.Dispatcher.BeginInvoke(sair);
        }

        public void Dispose()
        {
            if (encerrada) return;
            encerrada = true;
            if (folha.IsMouseCaptured) folha.ReleaseMouseCapture();
            imagem.Source = null;
            detalheImg.Source = null;
            fonte.RootVisual = null;
            fonte.Dispose();
        }

        static class Native
        {
            [StructLayout(LayoutKind.Sequential)] internal struct RECT { internal int Left, Top, Right, Bottom; }
            [StructLayout(LayoutKind.Sequential)] internal struct POINT { internal int X, Y; }
            [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)] internal static extern bool IsWindow(IntPtr hwnd);
            [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)] internal static extern bool IsWindowVisible(IntPtr hwnd);
            [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)] internal static extern bool GetClientRect(IntPtr hwnd, out RECT rect);
            [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)] internal static extern bool ScreenToClient(IntPtr hwnd, ref POINT point);
            [DllImport("user32.dll")] internal static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint process);
            [DllImport("user32.dll", EntryPoint = "GetWindowLongW")] internal static extern int GetWindowLong(IntPtr hwnd, int index);
            [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool SetWindowPos(IntPtr hwnd, IntPtr after, int x, int y, int cx, int cy, uint flags);
        }
    }

    // Esta classe e instanciada pelo WpfHost definido em AppUI.xslt.
    // Sem construtor sem argumentos: uma instancia externa do Corel nunca e criada.
    public sealed class DockerUI : UserControl
    {
        readonly Corel corel;
        readonly bool[] visivel = { true, true, true, true };
        readonly Button[] linhas = new Button[5];
        readonly ShapePath[] olhos = new ShapePath[5];
        readonly TextBlock[] percentuais = new TextBlock[4];
        readonly StackPanel linhasEspeciais = new StackPanel();
        readonly List<ShapePath> olhosEspeciais = new List<ShapePath>();
        readonly List<Button> botoesEspeciais = new List<Button>();
        readonly List<TextBlock> percentuaisEspeciais = new List<TextBlock>();
        readonly Dictionary<string, bool> estadoEspeciais = new Dictionary<string, bool>(StringComparer.Ordinal);
        bool[] especiaisVisiveis = new bool[0];
        string chaveEspeciais;
        readonly DispatcherTimer timer;
        readonly DispatcherTimer timerNitidez; // v2.3
        int dpiImagem, dpiMaximoImagem; // v2.3
        // v2.4: PDF da ultima atualizacao e detalhe da area visivel
        string pdfCache, pastaPdfCache, motorCache;
        double paginaLarguraIn, paginaAlturaIn;
        ImagemCmyk detalhe;
        double detU0, detV0, detU1, detV1;
        bool renderizandoDetalhe, detalhePendente;
        int geracaoDetalhe;
        int detPx0, detPy0, detVersao = -1, geracaoDesenhoDetalhe; // v2.6
        double detEscala = 1;
        Matrix detMatriz = Matrix.Identity; // v2.8: zoom/posicao da previa quando o detalhe foi pedido
        // v2.9: conferencia de cores fora do CMYK
        readonly TextBlock rgbResumo = new TextBlock();
        readonly StackPanel rgbLista = new StackPanel();
        readonly Button rgbSelecionar = new Button();
        // v3.6: lista completa, filtro por tipo e Anterior/Próximo com zoom
        readonly WrapPanel rgbFiltros = new WrapPanel();
        readonly Grid rgbNavegar = new Grid();
        readonly TextBlock rgbPosicao = new TextBlock();
        readonly Button rgbCompleta = new Button();
        readonly List<Button> linhasRgb = new List<Button>();
        List<ItemRgb> filtradosRgb = new List<ItemRgb>();
        Button rgbMais;
        int filtroRgb, posicaoRgb = -1;
        const int BlocoListaRgb = 150;
        static readonly string[] NomesFiltroRgb = { "Todos", "Imagens RGB", "Cores RGB", "Preto/cinza", "Registro" };
        static readonly Brush LinhaRgb = Pincel(79, 79, 79);
        static readonly Brush LinhaRgbAtual = Pincel(52, 96, 130);
        ResultadoRgb resultadoRgb;
        bool rgbPendente;
        int geracaoRgb;
        string perfilCmyk; // v2.6: perfil CMYK para as cores na tela
        readonly Image previa = new Image();
        readonly Grid area = new Grid();
        readonly TextBlock vazio = new TextBlock();
        readonly TextBlock status = new TextBlock();
        readonly TextBlock leitura = new TextBlock();
        readonly CheckBox sobreimpressao = new CheckBox();
        readonly CheckBox naArte = new CheckBox();
        readonly Button voltar = new Button();
        readonly ComboBox resolucao = new ComboBox();
        // v3.6: sem "Mostrar chapas em preto" e "Carga CMYK acima de"
        readonly Button atualizar = new Button();
        ImagemCmyk imagem;
        PreviaNaArte tela;
        bool telaSuspensa;
        bool ativo, carregando, renderizando, renderPendente, sincronizando;
        int geracao, geracaoRender;
        string chaveAtual;
        DateTime proximaTentativa = DateTime.MinValue;
        int ticksOcioso; // v3.5: sem a previa sobre a arte, confere o Corel com menos frequencia
        static readonly Brush Fundo = Pincel(67, 67, 67);
        static readonly Brush Texto = Pincel(235, 235, 235);
        static readonly Brush Suave = Pincel(190, 190, 190);
        static readonly Brush Alerta = Pincel(255, 184, 77);  // v2.9
        static readonly Brush Certo = Pincel(140, 210, 140);  // v2.9

        public DockerUI(object app)
        {
            corel = new Corel(app);
            FontFamily = new FontFamily("Segoe UI");
            FontSize = 12;
            Background = Fundo;
            Foreground = Texto;
            MinWidth = 260;
            MinHeight = 390;
            ConstruirInterface();
            timer = new DispatcherTimer(DispatcherPriority.Background, Dispatcher);
            timer.Interval = TimeSpan.FromMilliseconds(500);
            timer.Tick += AoTick;
            // v2.3: depois de um zoom, gera a previa de novo na resolucao da tela.
            timerNitidez = new DispatcherTimer(DispatcherPriority.Background, Dispatcher);
            timerNitidez.Interval = TimeSpan.FromMilliseconds(250);
            timerNitidez.Tick += delegate { AfiarSeNecessario(); };
            Loaded += delegate { Ativar(); };
            Unloaded += delegate { Desativar(); };
            IsVisibleChanged += delegate { if (IsVisible && IsLoaded) Ativar(); else Desativar(); };
            // F5 apenas quando o foco esta no painel.
            PreviewKeyDown += delegate(object s, KeyEventArgs e)
            {
                if (e.Key == Key.Escape) { VoltarEdicao(); e.Handled = true; }
                else if (e.Key == Key.F5) { Atualizar(); e.Handled = true; }
            };
        }

        static Brush Pincel(byte r, byte g, byte b)
        {
            SolidColorBrush p = new SolidColorBrush(Color.FromRgb(r, g, b));
            p.Freeze(); return p;
        }

        static TextBlock Rotulo(string texto, Brush cor)
        {
            return new TextBlock { Text = texto, Foreground = cor, TextWrapping = TextWrapping.Wrap,
                                   VerticalAlignment = VerticalAlignment.Center };
        }

        // v2.9: botao cinza simples, igual as linhas das tintas, mesmo com tema claro do Windows.
        static ControlTemplate ModeloCinza()
        {
            FrameworkElementFactory borda = new FrameworkElementFactory(typeof(Border));
            borda.SetValue(Border.BackgroundProperty, new TemplateBindingExtension(Button.BackgroundProperty));
            borda.SetValue(Border.BorderBrushProperty, new TemplateBindingExtension(Button.BorderBrushProperty));
            borda.SetValue(Border.BorderThicknessProperty, new TemplateBindingExtension(Button.BorderThicknessProperty));
            borda.SetValue(Border.PaddingProperty, new TemplateBindingExtension(Button.PaddingProperty));
            FrameworkElementFactory apresentador = new FrameworkElementFactory(typeof(ContentPresenter));
            apresentador.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            borda.AppendChild(apresentador);
            ControlTemplate modelo = new ControlTemplate(typeof(Button)); modelo.VisualTree = borda;
            Trigger hover = new Trigger { Property = Button.IsMouseOverProperty, Value = true };
            hover.Setters.Add(new Setter(Button.BackgroundProperty, Pincel(97, 97, 97))); modelo.Triggers.Add(hover);
            return modelo;
        }

        static void ConfigurarCaixa(CheckBox c, string texto)
        {
            c.Content = Rotulo(texto, Texto);
            c.Margin = new Thickness(10, 5, 8, 5);
            c.Foreground = Texto;
        }

        void ConstruirInterface()
        {
            Grid raiz = new Grid();
            // v2.5: sem a miniatura da pagina no painel; as opcoes ocupam a altura toda.
            raiz.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            raiz.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            raiz.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            Content = raiz;

            StackPanel topo = new StackPanel();
            ScrollViewer topoRolavel = new ScrollViewer { Content = topo,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto, HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled };
            raiz.Children.Add(topoRolavel);
            TextBlock titulo = Rotulo("CMYK e Pantone · v3.8", Texto);
            titulo.FontWeight = FontWeights.SemiBold;
            titulo.Margin = new Thickness(10, 9, 8, 7);
            topo.Children.Add(titulo);
            ConfigurarCaixa(naArte, "Mostrar a prévia na área da arte");
            naArte.IsChecked = true;
            naArte.ToolTip = "Ao desligar uma chapa, mostra uma simulação sobre a área do documento. Esc volta à edição. Não altera os objetos.";
            naArte.Click += delegate { telaSuspensa = false; AtualizarTela(); };
            topo.Children.Add(naArte);
            naArte.Visibility = Visibility.Collapsed; // v2.5: a previa e sempre sobre a arte
            voltar.Content = "Voltar à edição  ·  Esc";
            voltar.Margin = new Thickness(10, 0, 10, 5); voltar.Height = 27;
            voltar.Click += delegate { VoltarEdicao(); }; topo.Children.Add(voltar);
            ConfigurarCaixa(sobreimpressao, "Sobreimpressão na página do Corel");
            sobreimpressao.ToolTip = "Controla a vista original do Corel. As chapas da prévia preservam a sobreimpressão exportada no PDF; a aparência da composição é aproximada.";
            sobreimpressao.IsEnabled = false;
            sobreimpressao.Click += delegate
            {
                if (sincronizando) return;
                try { VoltarEdicao(); corel.DefinirSobreimpressao(sobreimpressao.IsChecked == true); SincronizarSobreimpressao(); }
                catch (Exception ex) { Erro(ex, false); SincronizarSobreimpressao(); }
            };
            topo.Children.Add(sobreimpressao);

            string[] nomes = { "CMYK", "Ciano", "Magenta", "Amarelo", "Preto" };
            Brush[] tintas = { Suave, Pincel(0, 174, 239), Pincel(236, 0, 140), Pincel(255, 242, 0), Brushes.Black };
            for (int i = 0; i < 5; i++)
            {
                int indice = i;
                Grid conteudo = new Grid { Margin = new Thickness(9, 0, 9, 0) };
                conteudo.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(25) });
                conteudo.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(21) });
                conteudo.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                conteudo.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(40) });
                ShapePath olho = new ShapePath { Data = Geometry.Parse("M 1,6 C 4,0 12,0 15,6 C 12,12 4,12 1,6 Z M 10,6 A 2,2 0 1 1 6,6 A 2,2 0 1 1 10,6"),
                    Fill = Suave, Width = 16, Height = 12, Stretch = Stretch.Uniform, HorizontalAlignment = HorizontalAlignment.Left,
                    VerticalAlignment = VerticalAlignment.Center };
                olhos[i] = olho;
                conteudo.Children.Add(olho);
                Border amostra = new Border { Background = tintas[i], BorderBrush = Pincel(35, 35, 35), BorderThickness = new Thickness(1),
                    Width = 12, Height = 12, HorizontalAlignment = HorizontalAlignment.Left, VerticalAlignment = VerticalAlignment.Center };
                if (i == 0)
                {
                    Grid quatro = new Grid();
                    quatro.ColumnDefinitions.Add(new ColumnDefinition()); quatro.ColumnDefinitions.Add(new ColumnDefinition());
                    quatro.RowDefinitions.Add(new RowDefinition()); quatro.RowDefinitions.Add(new RowDefinition());
                    for (int n = 0; n < 4; n++)
                    {
                        Border cor = new Border { Background = tintas[n + 1] };
                        Grid.SetRow(cor, n / 2); Grid.SetColumn(cor, n % 2); quatro.Children.Add(cor);
                    }
                    amostra.Child = quatro;
                }
                Grid.SetColumn(amostra, 1); conteudo.Children.Add(amostra);
                TextBlock nome = Rotulo(nomes[i], Texto);
                Grid.SetColumn(nome, 2); conteudo.Children.Add(nome);
                if (i > 0)
                {
                    TextBlock pct = Rotulo("—", Suave); pct.TextAlignment = TextAlignment.Right;
                    percentuais[i - 1] = pct; Grid.SetColumn(pct, 3); conteudo.Children.Add(pct);
                }
                Button linha = new Button { Content = conteudo, Height = 27, Foreground = Texto,
                    Background = Pincel(79, 79, 79), BorderBrush = Pincel(94, 94, 94), BorderThickness = new Thickness(0, 0, 0, 1),
                    HorizontalContentAlignment = HorizontalAlignment.Stretch, Padding = new Thickness(0),
                    ToolTip = i == 0 ? "Mostrar ou ocultar as quatro chapas CMYK; não altera os Pantones" : "Clique: mostrar/ocultar. Shift + clique ou botão direito: isolar esta chapa." };
                // Template simples para manter o estilo cinza mesmo com tema claro do Windows.
                FrameworkElementFactory borda = new FrameworkElementFactory(typeof(Border));
                borda.SetValue(Border.BackgroundProperty, new TemplateBindingExtension(Button.BackgroundProperty));
                borda.SetValue(Border.BorderBrushProperty, new TemplateBindingExtension(Button.BorderBrushProperty));
                borda.SetValue(Border.BorderThicknessProperty, new TemplateBindingExtension(Button.BorderThicknessProperty));
                FrameworkElementFactory apresentador = new FrameworkElementFactory(typeof(ContentPresenter));
                apresentador.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
                borda.AppendChild(apresentador);
                ControlTemplate modelo = new ControlTemplate(typeof(Button)); modelo.VisualTree = borda;
                Trigger hover = new Trigger { Property = Button.IsMouseOverProperty, Value = true };
                hover.Setters.Add(new Setter(Button.BackgroundProperty, Pincel(97, 97, 97))); modelo.Triggers.Add(hover);
                Trigger foco = new Trigger { Property = Button.IsKeyboardFocusedProperty, Value = true };
                foco.Setters.Add(new Setter(Button.BorderBrushProperty, Pincel(132, 188, 221))); modelo.Triggers.Add(foco);
                linha.Template = modelo;
                linha.Click += delegate { Alternar(indice, (Keyboard.Modifiers & ModifierKeys.Shift) != 0); };
                linha.PreviewMouseRightButtonUp += delegate(object s, MouseButtonEventArgs e)
                { if (indice > 0) { Alternar(indice, true); e.Handled = true; } };
                linhas[i] = linha; topo.Children.Add(linha);
            }

            topo.Children.Add(linhasEspeciais);
            TextBlock dica = Rotulo("Desligue uma tinta para visualizar · Shift + clique isola", Suave);
            dica.Margin = new Thickness(10, 5, 8, 8); dica.FontSize = 11; topo.Children.Add(dica);
            // v3.6: sem "Mostrar chapas em preto" e "Carga CMYK acima de" (deixavam o painel mais pesado)
            // v3.7: sem a linha "Carga máxima CMYK"

            Grid acoes = new Grid { Margin = new Thickness(10, 0, 10, 3) };
            acoes.ColumnDefinitions.Add(new ColumnDefinition());
            acoes.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(85) });
            atualizar.Content = "Atualizar página  ·  F5"; atualizar.Height = 29; atualizar.Margin = new Thickness(0, 0, 7, 0);
            atualizar.Click += delegate { Atualizar(); }; acoes.Children.Add(atualizar);
            resolucao.Items.Add("Tela"); // v2.3: acompanha o zoom do Corel e da previa
            foreach (int dpi in new int[] { 72, 100, 150, 200, 300 }) resolucao.Items.Add(dpi + " dpi");
            resolucao.SelectedIndex = 0; resolucao.ToolTip = "Tela: acompanha o zoom do Corel e da prévia (mais nítida). Ou escolha uma resolução fixa.";
            resolucao.SelectionChanged += delegate { if (ativo) Atualizar(); };
            Grid.SetColumn(resolucao, 1); acoes.Children.Add(resolucao); topo.Children.Add(acoes);
            // v2.8: sem a opcao "Atualizar automaticamente (a cada 5 s)".

            // v2.9: cores fora do CMYK (RGB e afins, imagens RGB, cor de registro) na pagina atual
            Border separador = new Border { Height = 1, Background = Pincel(94, 94, 94), Margin = new Thickness(10, 10, 10, 7) };
            topo.Children.Add(separador);
            TextBlock tituloRgb = Rotulo("Conferência de cores", Texto);
            tituloRgb.FontWeight = FontWeights.SemiBold; tituloRgb.Margin = new Thickness(10, 0, 8, 3);
            tituloRgb.ToolTip = "Procura RGB (e HSB, HLS, Lab), imagens RGB, cor de registro e preto ou cinza que não é só K (ex.: M5 Y2 K50) na página, inclusive dentro de grupos e PowerClips. Confere a cada Atualizar página.";
            topo.Children.Add(tituloRgb);
            rgbResumo.Text = "—"; rgbResumo.Foreground = Suave; rgbResumo.TextWrapping = TextWrapping.Wrap;
            rgbResumo.Margin = new Thickness(10, 0, 8, 5);
            topo.Children.Add(rgbResumo);
            // v3.6: conferencia parcial (parou em 5 s) -> conferir sem limite de tempo
            rgbCompleta.Content = "Conferir a página toda"; rgbCompleta.Height = 25; rgbCompleta.Margin = new Thickness(10, 0, 10, 5);
            rgbCompleta.ToolTip = "A conferência parou em 5 s para não travar o Corel. Clique para conferir tudo (em artes grandes pode demorar).";
            rgbCompleta.Visibility = Visibility.Collapsed;
            rgbCompleta.Click += delegate { ConferirRgbCompleta(); };
            topo.Children.Add(rgbCompleta);
            // v3.6: filtro por tipo (Imagens RGB, Cores RGB, Preto/cinza, Registro)
            rgbFiltros.Margin = new Thickness(10, 0, 6, 3);
            rgbFiltros.Visibility = Visibility.Collapsed;
            topo.Children.Add(rgbFiltros);
            // v3.6: Anterior / Próximo: seleciona e dá zoom item por item
            rgbNavegar.Margin = new Thickness(10, 1, 10, 4);
            rgbNavegar.ColumnDefinitions.Add(new ColumnDefinition());
            rgbNavegar.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            rgbNavegar.ColumnDefinitions.Add(new ColumnDefinition());
            Button rgbAnterior = new Button { Content = "‹ Anterior", Height = 27,
                ToolTip = "Seleciona o item anterior da lista e dá zoom nele no Corel" };
            rgbAnterior.Click += delegate { IrParaRgb(posicaoRgb - 1); };
            rgbNavegar.Children.Add(rgbAnterior);
            rgbPosicao.Foreground = Texto; rgbPosicao.FontSize = 11; rgbPosicao.MinWidth = 56;
            rgbPosicao.TextAlignment = TextAlignment.Center; rgbPosicao.VerticalAlignment = VerticalAlignment.Center;
            rgbPosicao.Margin = new Thickness(8, 0, 8, 0);
            Grid.SetColumn(rgbPosicao, 1); rgbNavegar.Children.Add(rgbPosicao);
            Button rgbProximo = new Button { Content = "Próximo ›", Height = 27,
                ToolTip = "Seleciona o próximo item da lista e dá zoom nele no Corel" };
            rgbProximo.Click += delegate { IrParaRgb(posicaoRgb + 1); };
            Grid.SetColumn(rgbProximo, 2); rgbNavegar.Children.Add(rgbProximo);
            rgbNavegar.Visibility = Visibility.Collapsed;
            topo.Children.Add(rgbNavegar);
            rgbSelecionar.Content = "Selecionar no Corel"; rgbSelecionar.Height = 27; rgbSelecionar.Margin = new Thickness(10, 1, 10, 5);
            rgbSelecionar.ToolTip = "Volta à edição e seleciona de uma vez os objetos da lista (o objeto de fora, quando estão em grupo ou PowerClip).";
            rgbSelecionar.Visibility = Visibility.Collapsed;
            rgbSelecionar.Click += delegate { SelecionarRgb(null); };
            topo.Children.Add(rgbSelecionar);
            rgbLista.Margin = new Thickness(10, 0, 10, 8);
            topo.Children.Add(rgbLista);

            area.Background = Pincel(43, 43, 43);
            area.Margin = new Thickness(10, 4, 10, 4);
            area.MinHeight = 75;
            previa.Stretch = Stretch.Uniform;
            RenderOptions.SetBitmapScalingMode(previa, BitmapScalingMode.HighQuality); // v2.2
            area.Children.Add(previa);
            vazio.Text = "Abra uma arte no CorelDRAW."; vazio.Foreground = Suave;
            vazio.TextAlignment = TextAlignment.Center; vazio.TextWrapping = TextWrapping.Wrap;
            vazio.VerticalAlignment = VerticalAlignment.Center; vazio.Margin = new Thickness(12);
            area.Children.Add(vazio);
            area.MouseMove += MouseNaPrevia;
            area.MouseLeave += delegate { LimparLeitura(); };
            area.Visibility = Visibility.Collapsed; // v2.5: miniatura removida (a imagem continua servindo a previa sobre a arte)
            Grid.SetRow(area, 1); raiz.Children.Add(area);

            StackPanel rodape = new StackPanel { Margin = new Thickness(10, 0, 10, 9) };
            leitura.Text = "Desligue uma chapa e passe o mouse na arte para medir as tintas.";
            leitura.Foreground = Texto; leitura.FontSize = 11; leitura.TextWrapping = TextWrapping.Wrap;
            rodape.Children.Add(leitura);
            status.TextWrapping = TextWrapping.Wrap; status.Foreground = Suave; status.FontSize = 11;
            status.Margin = new Thickness(0, 6, 0, 4); rodape.Children.Add(status);
            TextBlock aviso = Rotulo("Modo de consulta: Esc volta à edição. CMYK e tintas especiais em chapas independentes; aparência das cores aproximada. Atualize após editar.", Suave);
            aviso.FontSize = 10; rodape.Children.Add(aviso);
            // assinatura
            TextBlock assinatura = Rotulo("by Bruno Faqui", Suave);
            assinatura.FontSize = 11; assinatura.FontStyle = FontStyles.Italic; assinatura.Margin = new Thickness(0, 8, 0, 0);
            rodape.Children.Add(assinatura);
            Grid.SetRow(rodape, 2); raiz.Children.Add(rodape);
        }

        void Ativar()
        {
            if (ativo || !IsVisible) return;
            ativo = true; timer.Start();
            // v3.5: nao gera a previa sozinho; so quando pedir (Atualizar página/F5 ou clicar numa tinta)
            if (imagem == null)
            {
                status.Text = "Clique em Atualizar página (F5) ou numa tinta para ler a página.";
                status.Foreground = Suave;
            }
        }

        void Desativar()
        {
            if (!ativo) return;
            ativo = false; timer.Stop(); timerNitidez.Stop();
            // v3.5: guarda a leitura (ao voltar, nao gera de novo se a pagina for a mesma); so fecha a previa
            FecharTela();
        }

        void AoTick(object sender, EventArgs e)
        {
            if (!ativo || carregando || DateTime.UtcNow < proximaTentativa) return;
            // v3.5: sem a previa sobre a arte, confere so a cada 1,5 s
            if (tela == null && (++ticksOcioso % 3) != 0) return;
            try
            {
                string chave = corel.ChaveAtual();
                if (chave != chaveAtual)
                {
                    telaSuspensa = true;
                    LimparImagem(); chaveAtual = chave;
                    // v3.5: nao gera a previa sozinho ao trocar de pagina ou de arquivo
                    if (chave != null)
                    {
                        status.Text = "A página ou o arquivo mudou. Clique em Atualizar página (F5) para ler.";
                        status.Foreground = Suave;
                    }
                }
                if (tela != null && !tela.Conferir(corel.LerVista())) { telaSuspensa = true; FecharTela(); }
                SincronizarSobreimpressao();
                if (chave == null) status.Text = "Abra uma arte no CorelDRAW.";
            }
            catch (Exception ex) { Erro(ex, true); }
        }

        void SincronizarSobreimpressao()
        {
            sincronizando = true;
            try { sobreimpressao.IsChecked = corel.LerSobreimpressao(); sobreimpressao.IsEnabled = true; }
            catch { sobreimpressao.IsChecked = false; sobreimpressao.IsEnabled = false; }
            finally { sincronizando = false; }
        }

        void HabilitarChapas(bool habilitar)
        {
            foreach (Button linha in linhas) linha.IsEnabled = habilitar;
            linhasEspeciais.IsEnabled = habilitar;
        }

        // v2.8: com a previa aberta sobre a arte na mesma pagina, atualiza sem fechar a previa:
        // o zoom e a posicao ficam onde estao e a imagem e trocada no lugar.
        void Atualizar()
        {
            bool manter = false;
            if (tela != null && imagem != null && chaveAtual != null)
            {
                try { manter = corel.ChaveAtual() == chaveAtual; } catch (Exception ex) { Corel.Registrar(ex); }
            }
            rgbPendente = true; // v2.9: confere as cores depois desta atualizacao
            Atualizar(manter);
        }

        // v2.3: manterTela = gera de novo em resolucao maior sem fechar a previa sobre a arte.
        void Atualizar(bool manterTela)
        {
            if (!ativo || carregando) return;
            carregando = true; atualizar.IsEnabled = false;
            HabilitarChapas(false);
            resolucao.IsEnabled = false;
            if (!manterTela) LimparImagem();
            int minhaGeracao = geracao;
            status.Text = manterTela ? "Atualizando a prévia…" : "Lendo a página atual…"; status.Foreground = Suave;
            if (!manterTela) vazio.Text = "Gerando prévia…";
            Captura captura;
            try
            {
                captura = corel.Exportar(DpiPedido());
                chaveAtual = captura.Chave;
                SincronizarSobreimpressao();
            }
            catch (Exception ex)
            {
                carregando = false; atualizar.IsEnabled = true; resolucao.IsEnabled = true; HabilitarChapas(true);
                Erro(ex, !manterTela); return;
            }
            // Nenhum acesso ao Corel em segundo plano; apenas leitura e calculos.
            Task.Factory.StartNew(delegate
            {
                try { return LeitorSeparacoes.Ler(captura); }
                finally { MotorSeparacoes.Limpar(captura.Diretorio); }
            }).ContinueWith(delegate(Task<ImagemCmyk> tarefa)
            {
                Exception falha = tarefa.IsFaulted ? tarefa.Exception : null; // observa a excecao mesmo ao fechar o host.
                if (Dispatcher.HasShutdownStarted) { MotorSeparacoes.Limpar(captura.PastaPdf); return; }
                Dispatcher.BeginInvoke(new Action(delegate
                {
                    carregando = false; atualizar.IsEnabled = true; resolucao.IsEnabled = true; HabilitarChapas(true);
                    if (!ativo || minhaGeracao != geracao) { MotorSeparacoes.Limpar(captura.PastaPdf); return; }
                    try
                    {
                        if (corel.ChaveAtual() != captura.Chave) { MotorSeparacoes.Limpar(captura.PastaPdf); LimparImagem(); chaveAtual = null; return; }
                        if (falha != null) { MotorSeparacoes.Limpar(captura.PastaPdf); Erro(falha, !manterTela); return; }
                        imagem = tarefa.Result;
                        dpiImagem = captura.Dpi; dpiMaximoImagem = captura.DpiMaximo;
                        AdotarPdf(captura);
                        ConstruirEspeciais(captura.Chave);
                        proximaTentativa = DateTime.MinValue;
                        status.Text = captura.Descricao + " · " + imagem.Largura + " × " + imagem.Altura + " px · " + imagem.Especiais.Length + " especiais\nAtualizada às " + DateTime.Now.ToString("HH:mm:ss");
                        status.Foreground = Suave;
                        Redesenhar();
                        if (rgbPendente) { rgbPendente = false; AgendarRgb(captura.Chave); } // v2.9
                    }
                    catch (Exception ex) { Erro(ex, !manterTela); }
                }), DispatcherPriority.Background);
            }, TaskScheduler.Default);
        }

        // v2.3: resolucao pedida ao Ghostscript. "Tela" = como a pagina aparece agora no Corel,
        // com folga de 25% e o zoom da previa sobre a arte. O limite por pagina fica em Exportar.
        int DpiPedido()
        {
            int[] dpis = { 72, 100, 150, 200, 300 };
            int indice = resolucao.SelectedIndex;
            if (indice > 0) return dpis[Math.Min(indice - 1, dpis.Length - 1)];
            double dpi = 0;
            try { dpi = corel.DpiDaTela(); } catch (Exception ex) { Corel.Registrar(ex); }
            if (double.IsNaN(dpi) || double.IsInfinity(dpi) || dpi <= 0) return 150;
            // v2.8: com o detalhe por area, a pagina inteira nao precisa crescer com o zoom da previa
            // (a atualizacao com a previa aberta continua rapida).
            double escala = (tela != null && !MotorRegiao.Disponivel()) ? Math.Max(1.0, tela.Escala) : 1.0;
            dpi = dpi * escala * 1.25;
            return (int)Math.Max(100, Math.Min(1200, Math.Ceiling(dpi)));
        }

        void ReiniciarNitidez()
        {
            timerNitidez.Stop();
            timerNitidez.Start();
        }

        // v2.3: se o zoom pede mais detalhe do que a imagem atual tem, gera de novo sem fechar a previa.
        void AfiarSeNecessario()
        {
            timerNitidez.Stop();
            if (!ativo || carregando || imagem == null || tela == null) return;
            if (DateTime.UtcNow < proximaTentativa) return;
            // v2.4: gera so a area visivel, na resolucao da tela
            if (pdfCache != null && MotorRegiao.Disponivel()) { PedirDetalhe(); return; }
            // sem detalhe por area: como na v2.3, gera a pagina inteira maior
            if (resolucao.SelectedIndex > 0 || dpiImagem >= dpiMaximoImagem) return;
            if (DpiPedido() > dpiImagem * 1.3) Atualizar(true);
        }

        // v2.4: guarda o PDF desta atualizacao (o anterior e apagado)
        void AdotarPdf(Captura c)
        {
            if (pastaPdfCache != null && pastaPdfCache != c.PastaPdf) MotorSeparacoes.Limpar(pastaPdfCache);
            pastaPdfCache = c.PastaPdf; pdfCache = c.Arquivo; motorCache = c.Motor;
            paginaLarguraIn = c.LarguraIn; paginaAlturaIn = c.AlturaIn;
            perfilCmyk = CorDeTela.LocalizarPerfil(c.Motor);
            // v2.8: o detalhe anterior continua na tela ate o novo ficar pronto (sem piscar).
            geracaoDetalhe++; detVersao = -1;
            if (tela != null) ReiniciarNitidez();
        }

        void LimparPdfCache()
        {
            if (pastaPdfCache != null) MotorSeparacoes.Limpar(pastaPdfCache);
            pastaPdfCache = null; pdfCache = null; motorCache = null;
        }

        // v2.6: pede ao Ghostscript a area visivel com exatamente um pixel por pixel da tela.
        void PedirDetalhe()
        {
            if (tela == null || imagem == null || pdfCache == null || paginaLarguraIn <= 0 || paginaAlturaIn <= 0) return;
            if (renderizandoDetalhe) { detalhePendente = true; return; }
            double u0, v0, u1, v1, dpi, escala;
            int px0, py0, pxL, pxA;
            if (!tela.RegiaoTela(paginaLarguraIn, out u0, out v0, out u1, out v1, out px0, out py0, out pxL, out pxA, out dpi, out escala)) return;
            int versao = tela.VersaoVista;
            Matrix matriz = tela.MatrizVista; // v2.8
            if (detalhe != null && detVersao == versao) return; // ja existe para esta vista
            if ((long)pxL * pxA > 16000000L) return;
            int fator = (long)pxL * pxA * 4 <= 48000000L ? 2 : 1;
            double W = paginaLarguraIn * 72.0, H = paginaAlturaIn * 72.0;
            double x0 = u0 * W, y0 = (1 - v1) * H;
            renderizandoDetalhe = true; detalhePendente = false;
            int minha = ++geracaoDetalhe;
            ImagemCmyk pagina = imagem;
            string exe = motorCache, pdf = pdfCache;
            Task.Factory.StartNew(delegate { return MotorRegiao.Renderizar(exe, pdf, pagina, dpi, fator, x0, y0, pxL, pxA, u0, v0, u1, v1); })
                .ContinueWith(delegate(Task<ImagemCmyk> tarefa)
                {
                    Exception falha = tarefa.IsFaulted ? tarefa.Exception : null;
                    if (Dispatcher.HasShutdownStarted) return;
                    Dispatcher.BeginInvoke(new Action(delegate
                    {
                        renderizandoDetalhe = false;
                        if (falha != null) Corel.Registrar(falha);
                        // v2.8: aceita o detalhe mesmo se a vista mudou durante a geracao; ele aparece ajustado
                        // a vista atual e um novo e pedido em seguida.
                        bool valido = ativo && minha == geracaoDetalhe && pagina == imagem && tela != null;
                        ImagemCmyk resultado = (falha == null && valido) ? tarefa.Result : null;
                        if (resultado != null)
                        {
                            detalhe = resultado; detU0 = u0; detV0 = v0; detU1 = u1; detV1 = v1;
                            detPx0 = px0; detPy0 = py0; detEscala = escala; detVersao = versao; detMatriz = matriz;
                            DesenharDetalhe();
                            if (tela.VersaoVista != versao) detalhePendente = true;
                        }
                        else if (ativo && falha == null && !MotorRegiao.Disponivel())
                        {
                            status.Text = "O Ghostscript não gerou só a área visível; usando a página inteira. Detalhes em diagnostico.log.";
                            AfiarSeNecessario();
                        }
                        else if (ativo && tela != null && tela.VersaoVista != versao) detalhePendente = true;
                        if (detalhePendente && ativo) PedirDetalhe();
                    }), DispatcherPriority.Background);
                }, TaskScheduler.Default);
        }

        // v2.6: aplica as chapas ligadas/desligadas ao detalhe (cores pelo perfil CMYK) e mostra sobre a arte.
        void DesenharDetalhe()
        {
            if (tela == null) return;
            if (detalhe == null || detalhe.Especiais.Length != especiaisVisiveis.Length)
            { tela.OcultarDetalhe(); return; }
            ImagemCmyk img = detalhe;
            bool[] chapas = (bool[])visivel.Clone(), especiais = (bool[])especiaisVisiveis.Clone();
            bool preto = false, carga = false; int pct = 300; // v3.6: opcoes removidas do painel
            string perfil = perfilCmyk;
            int px0 = detPx0, py0 = detPy0;
            Matrix matriz = detMatriz; // v2.8
            double escala = detEscala;
            int minha = ++geracaoDesenhoDetalhe;
            Task.Factory.StartNew(delegate { return CorDeTela.Desenhar(img, chapas, preto, carga, pct, especiais, perfil); })
                .ContinueWith(delegate(Task<byte[]> tarefa)
                {
                    Exception falha = tarefa.IsFaulted ? tarefa.Exception : null;
                    if (Dispatcher.HasShutdownStarted) return;
                    Dispatcher.BeginInvoke(new Action(delegate
                    {
                        if (falha != null) { Corel.Registrar(falha); return; }
                        if (tela == null || minha != geracaoDesenhoDetalhe || img != detalhe) return;
                        try
                        {
                            BitmapSource bitmap = BitmapSource.Create(img.Largura, img.Altura, 96 * escala, 96 * escala,
                                                                      PixelFormats.Bgra32, null, tarefa.Result, img.Largura * 4);
                            bitmap.Freeze();
                            tela.ExibirDetalhe(bitmap, px0, py0, escala, matriz);
                        }
                        catch (Exception ex) { Corel.Registrar(ex); }
                    }), DispatcherPriority.Background);
                }, TaskScheduler.Default);
        }

        void LimparImagem()
        {
            FecharTela();
            geracao++; geracaoRender++;
            detalhe = null; geracaoDetalhe++; LimparPdfCache(); // v2.4
            imagem = null; previa.Source = null; vazio.Visibility = Visibility.Visible;
            vazio.Text = "Abra uma arte e clique em Atualizar.";
            LimparLeitura();
            geracaoRgb++; resultadoRgb = null; MostrarRgb(); // v2.9
        }

        void Erro(Exception ex, bool limpar)
        {
            Corel.Registrar(ex);
            if (limpar) LimparImagem();
            status.Text = Corel.Explicar(ex); status.Foreground = Pincel(255, 196, 135);
            vazio.Text = "Prévia indisponível. Veja a mensagem abaixo.";
            proximaTentativa = DateTime.UtcNow.AddSeconds(10);
        }

        void Alternar(int linha, bool isolar)
        {
            telaSuspensa = false;
            if (linha == 0)
            {
                bool todas = visivel[0] && visivel[1] && visivel[2] && visivel[3];
                for (int i = 0; i < 4; i++) visivel[i] = !todas;
            }
            else if (isolar)
            {
                for (int i = 0; i < 4; i++) visivel[i] = i == linha - 1;
                for (int i = 0; i < especiaisVisiveis.Length; i++) DefinirEspecial(i, false);
            }
            else visivel[linha - 1] = !visivel[linha - 1];
            AtualizarOlhos();
            if (!PrecisaTela()) FecharTela();
            if (imagem == null && !carregando) Atualizar();
            else Redesenhar();
        }

        void ConstruirEspeciais(string chave)
        {
            if (chaveEspeciais != chave) { estadoEspeciais.Clear(); chaveEspeciais = chave; }
            linhasEspeciais.Children.Clear(); olhosEspeciais.Clear(); botoesEspeciais.Clear(); percentuaisEspeciais.Clear();
            especiaisVisiveis = new bool[imagem.Especiais.Length];
            TextBlock cabecalho = Rotulo(imagem.Especiais.Length == 0 ? "Nenhuma tinta especial nesta página" : "Tintas especiais / Pantone", Suave);
            cabecalho.Margin = new Thickness(10, 7, 8, 4); cabecalho.FontSize = 11;
            linhasEspeciais.Children.Add(cabecalho);
            for (int i = 0; i < imagem.Especiais.Length; i++)
            {
                int indice = i;
                TintaEspecial tinta = imagem.Especiais[i];
                bool estado;
                especiaisVisiveis[i] = !estadoEspeciais.TryGetValue(tinta.Nome, out estado) || estado;
                Grid conteudo = new Grid { Margin = new Thickness(9, 0, 9, 0) };
                conteudo.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(25) });
                conteudo.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(21) });
                conteudo.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                conteudo.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(40) });
                ShapePath olho = new ShapePath { Data = olhos[1].Data, Fill = Suave, Width = 16, Height = 12,
                    Stretch = Stretch.Uniform, HorizontalAlignment = HorizontalAlignment.Left, VerticalAlignment = VerticalAlignment.Center };
                conteudo.Children.Add(olho); olhosEspeciais.Add(olho);
                Border amostra = new Border { Background = Pincel(tinta.R, tinta.G, tinta.B), BorderBrush = Pincel(35, 35, 35),
                    BorderThickness = new Thickness(1), Width = 12, Height = 12,
                    HorizontalAlignment = HorizontalAlignment.Left, VerticalAlignment = VerticalAlignment.Center };
                Grid.SetColumn(amostra, 1); conteudo.Children.Add(amostra);
                TextBlock nome = Rotulo(tinta.Nome, Texto); nome.TextTrimming = TextTrimming.CharacterEllipsis;
                nome.TextWrapping = TextWrapping.NoWrap; Grid.SetColumn(nome, 2); conteudo.Children.Add(nome);
                TextBlock pct = Rotulo("—", Suave); pct.TextAlignment = TextAlignment.Right;
                Grid.SetColumn(pct, 3); conteudo.Children.Add(pct); percentuaisEspeciais.Add(pct);
                Button linha = new Button { Content = conteudo, Height = 27, Foreground = Texto,
                    Background = Pincel(79, 79, 79), BorderBrush = Pincel(94, 94, 94), BorderThickness = new Thickness(0, 0, 0, 1),
                    HorizontalContentAlignment = HorizontalAlignment.Stretch, Padding = new Thickness(0), Template = linhas[1].Template,
                    ToolTip = tinta.Nome + " · Clique: mostrar/ocultar. Shift + clique ou botão direito: isolar esta tinta." };
                linha.Click += delegate { AlternarEspecial(indice, (Keyboard.Modifiers & ModifierKeys.Shift) != 0); };
                linha.PreviewMouseRightButtonUp += delegate(object s, MouseButtonEventArgs e)
                { AlternarEspecial(indice, true); e.Handled = true; };
                linhasEspeciais.Children.Add(linha); botoesEspeciais.Add(linha);
            }
            AtualizarOlhos();
        }

        void DefinirEspecial(int indice, bool mostrar)
        {
            especiaisVisiveis[indice] = mostrar;
            if (imagem != null && indice < imagem.Especiais.Length)
                estadoEspeciais[imagem.Especiais[indice].Nome] = mostrar;
        }

        void AlternarEspecial(int indice, bool isolar)
        {
            if (imagem == null || indice >= especiaisVisiveis.Length) return;
            telaSuspensa = false;
            if (isolar)
            {
                for (int i = 0; i < 4; i++) visivel[i] = false;
                for (int i = 0; i < especiaisVisiveis.Length; i++) DefinirEspecial(i, i == indice);
            }
            else DefinirEspecial(indice, !especiaisVisiveis[indice]);
            AtualizarOlhos();
            if (!PrecisaTela()) FecharTela();
            Redesenhar();
        }

        void AtualizarOlhos()
        {
            int quantas = 0;
            for (int i = 0; i < 4; i++)
            {
                if (visivel[i]) quantas++;
                olhos[i + 1].Opacity = visivel[i] ? 1 : 0.12;
                linhas[i + 1].Opacity = visivel[i] ? 1 : 0.55;
            }
            olhos[0].Opacity = quantas == 4 ? 1 : quantas == 0 ? 0.12 : 0.5;
            for (int i = 0; i < especiaisVisiveis.Length; i++)
            {
                olhosEspeciais[i].Opacity = especiaisVisiveis[i] ? 1 : 0.12;
                botoesEspeciais[i].Opacity = especiaisVisiveis[i] ? 1 : 0.55;
            }
        }

        bool PrecisaTela()
        {
            bool algumaEspecialOculta = Array.IndexOf(especiaisVisiveis, false) >= 0;
            return ativo && naArte.IsChecked == true && !telaSuspensa &&
                (!visivel[0] || !visivel[1] || !visivel[2] || !visivel[3] || algumaEspecialOculta);
        }

        void FecharTela()
        {
            PreviaNaArte anterior = tela; tela = null;
            if (anterior != null) LimparLeitura(); // v2.5
            // v2.8: o detalhe pertence a previa que fechou (a proxima pode abrir com outra vista do Corel).
            if (anterior != null) { detalhe = null; detVersao = -1; geracaoDetalhe++; }
            if (anterior != null)
            {
                try { anterior.Dispose(); } catch (Exception ex) { Corel.Registrar(ex); }
            }
        }

        // v2.9: confere as cores depois que a previa aparece (o Corel fica parado alguns instantes em paginas grandes).
        void AgendarRgb(string chave) { AgendarRgb(chave, false); }

        // v3.6: completa = sem o limite de 5 s (pedido pelo botao "Conferir a página toda")
        void AgendarRgb(string chave, bool completa)
        {
            int minha = ++geracaoRgb;
            rgbResumo.Text = completa ? "Conferindo a página toda… (em artes grandes pode demorar)" : "Conferindo as cores da página…";
            rgbResumo.Foreground = Suave;
            Dispatcher.BeginInvoke(new Action(delegate
            {
                if (!ativo || minha != geracaoRgb) return;
                if (carregando) { rgbPendente = true; return; }
                try
                {
                    if (corel.ChaveAtual() != chave) return;
                    if (completa) Mouse.OverrideCursor = Cursors.Wait;
                    ResultadoRgb r = VerificadorRgb.Verificar(corel.Documento(), completa);
                    r.Chave = chave;
                    resultadoRgb = r;
                }
                catch (Exception ex)
                {
                    Corel.Registrar(ex);
                    resultadoRgb = null; MostrarRgb();
                    rgbResumo.Text = "Não foi possível conferir as cores: " + Corel.Explicar(ex);
                    return;
                }
                finally { if (completa) Mouse.OverrideCursor = null; }
                MostrarRgb();
            }), DispatcherPriority.ApplicationIdle);
        }

        void ConferirRgbCompleta()
        {
            ResultadoRgb r = resultadoRgb;
            if (r == null) return;
            try
            {
                if (corel.ChaveAtual() != r.Chave) { status.Text = "A página mudou. Clique em Atualizar página."; status.Foreground = Suave; return; }
            }
            catch (Exception ex) { Corel.Registrar(ex); return; }
            rgbCompleta.Visibility = Visibility.Collapsed;
            AgendarRgb(r.Chave, true);
        }

        void MostrarRgb()
        {
            rgbLista.Children.Clear(); linhasRgb.Clear(); rgbFiltros.Children.Clear(); rgbMais = null;
            filtradosRgb = new List<ItemRgb>(); posicaoRgb = -1;
            rgbFiltros.Visibility = Visibility.Collapsed; rgbNavegar.Visibility = Visibility.Collapsed;
            rgbSelecionar.Visibility = Visibility.Collapsed; rgbCompleta.Visibility = Visibility.Collapsed;
            ResultadoRgb r = resultadoRgb;
            if (r == null)
            {
                rgbResumo.Text = "—"; rgbResumo.Foreground = Suave;
                return;
            }
            string parcial = "";
            if (r.Parcial)
            {
                if (r.Completa) parcial = "\n(Conferência parcial: a página é grande demais até para a conferência completa.)";
                else
                {
                    parcial = "\n(Conferência parcial: parou em 5 s para não travar o Corel.)";
                    rgbCompleta.Visibility = Visibility.Visible;
                }
            }
            if (r.Itens.Count == 0)
            {
                rgbResumo.Text = "Nenhum RGB, cor de registro, nem preto ou cinza com mistura nesta página." + parcial;
                rgbResumo.Foreground = Certo;
                return;
            }
            List<string> partes = new List<string>();
            if (r.Objetos > 0) partes.Add(r.Objetos + (r.Objetos == 1 ? " objeto" : " objetos"));
            if (r.Imagens > 0) partes.Add(r.Imagens + (r.Imagens == 1 ? " imagem" : " imagens"));
            string texto = partes.Count > 0 ? "RGB nesta página: " + String.Join(" e ", partes.ToArray()) : "Nenhum RGB nesta página.";
            if (r.Preto > 0) // v3.2
                texto += "\nPreto ou cinza que não é só K: " + r.Preto + (r.Preto == 1 ? " objeto" : " objetos");
            if (r.Registro > 0)
                texto += "\nCor de registro: " + r.Registro + (r.Registro == 1 ? " objeto (sai em todas as chapas)" : " objetos (saem em todas as chapas)");
            if (r.Itens.Count >= VerificadorRgb.LimiteItens)
                texto += "\nA lista mostra os primeiros " + VerificadorRgb.LimiteItens + " itens.";
            rgbResumo.Text = texto + parcial;
            rgbResumo.Foreground = Alerta;

            // v3.6: botoes de filtro (so aparecem quando ha mais de um tipo)
            int[] contagem = new int[NomesFiltroRgb.Length];
            foreach (ItemRgb item in r.Itens)
                for (int f = 0; f < contagem.Length; f++) if (NoFiltroRgb(item, f)) contagem[f]++;
            if (filtroRgb < 0 || filtroRgb >= contagem.Length || contagem[filtroRgb] == 0) filtroRgb = 0;
            int tipos = 0;
            for (int f = 1; f < contagem.Length; f++) if (contagem[f] > 0) tipos++;
            if (tipos > 1)
            {
                for (int f = 0; f < contagem.Length; f++)
                {
                    if (contagem[f] == 0) continue;
                    int filtro = f;
                    TextBlock nome = Rotulo(NomesFiltroRgb[f] + " · " + contagem[f], Texto);
                    nome.FontSize = 11;
                    Button chip = new Button { Content = nome, Foreground = Texto,
                        Background = f == filtroRgb ? LinhaRgbAtual : LinhaRgb,
                        BorderBrush = Pincel(94, 94, 94), BorderThickness = new Thickness(1),
                        Padding = new Thickness(7, 2, 7, 2), Margin = new Thickness(0, 0, 4, 4), Cursor = Cursors.Hand,
                        ToolTip = f == 0 ? "Mostrar todos os itens" : "Mostrar só: " + NomesFiltroRgb[f] };
                    chip.Template = ModeloCinza();
                    chip.Click += delegate { if (filtroRgb != filtro) { filtroRgb = filtro; MostrarRgb(); } };
                    rgbFiltros.Children.Add(chip);
                }
                rgbFiltros.Visibility = Visibility.Visible;
            }

            foreach (ItemRgb item in r.Itens) if (NoFiltroRgb(item, filtroRgb)) filtradosRgb.Add(item);
            rgbNavegar.Visibility = Visibility.Visible;
            rgbSelecionar.Content = filtradosRgb.Count == 1 ? "Selecionar no Corel" : "Selecionar os " + filtradosRgb.Count + " no Corel";
            rgbSelecionar.Visibility = Visibility.Visible;
            AtualizarPosicaoRgb();
            MaisLinhasRgb();
        }

        // v3.6: 0 todos, 1 imagens RGB, 2 cores RGB (objetos), 3 preto/cinza com mistura, 4 cor de registro
        static bool NoFiltroRgb(ItemRgb item, int filtro)
        {
            switch (filtro)
            {
                case 1: return item.Imagem;
                case 2: return item.Tela && !item.Imagem;
                case 3: return item.Preto;
                case 4: return item.Registro;
                default: return true;
            }
        }

        // v3.6: a lista entra em blocos de 150 (com milhares de itens o painel nao trava); "Mostrar mais" acrescenta o proximo
        void MaisLinhasRgb()
        {
            if (rgbMais != null) { rgbLista.Children.Remove(rgbMais); rgbMais = null; }
            int de = linhasRgb.Count;
            int ate = Math.Min(filtradosRgb.Count, de + BlocoListaRgb);
            for (int i = de; i < ate; i++)
            {
                int indice = i;
                ItemRgb item = filtradosRgb[i];
                Grid conteudo = new Grid();
                conteudo.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                conteudo.ColumnDefinitions.Add(new ColumnDefinition());
                TextBlock numero = Rotulo((i + 1) + ".", Suave);
                numero.FontSize = 11; numero.Margin = new Thickness(0, 0, 6, 0); numero.VerticalAlignment = VerticalAlignment.Top;
                conteudo.Children.Add(numero);
                TextBlock rotulo = Rotulo(item.Descricao, Texto);
                rotulo.FontSize = 11;
                Grid.SetColumn(rotulo, 1); conteudo.Children.Add(rotulo);
                Button linha = new Button { Content = conteudo, Foreground = Texto,
                    Background = indice == posicaoRgb ? LinhaRgbAtual : LinhaRgb,
                    BorderBrush = Pincel(94, 94, 94), BorderThickness = new Thickness(0, 0, 0, 1),
                    HorizontalContentAlignment = HorizontalAlignment.Stretch, Padding = new Thickness(6, 3, 6, 3),
                    Cursor = Cursors.Hand, ToolTip = "Clique para selecionar e dar zoom neste objeto no Corel" };
                linha.Template = ModeloCinza();
                linha.Click += delegate { IrParaRgb(indice); };
                linhasRgb.Add(linha);
                rgbLista.Children.Add(linha);
            }
            int resto = filtradosRgb.Count - linhasRgb.Count;
            if (resto > 0)
            {
                TextBlock texto = Rotulo("Mostrar mais " + Math.Min(resto, BlocoListaRgb) + "  (faltam " + resto + ")", Texto);
                texto.FontSize = 11; texto.TextAlignment = TextAlignment.Center;
                rgbMais = new Button { Content = texto, Foreground = Texto, Background = Pincel(64, 64, 64),
                    BorderBrush = Pincel(94, 94, 94), BorderThickness = new Thickness(1),
                    Padding = new Thickness(6, 4, 6, 4), Margin = new Thickness(0, 4, 0, 0), Cursor = Cursors.Hand,
                    ToolTip = "Também dá para passar item por item com Anterior / Próximo" };
                rgbMais.Template = ModeloCinza();
                rgbMais.Click += delegate { MaisLinhasRgb(); };
                rgbLista.Children.Add(rgbMais);
            }
        }

        void AtualizarPosicaoRgb()
        {
            int n = filtradosRgb.Count;
            rgbPosicao.Text = posicaoRgb < 0 ? (n == 1 ? "1 item" : n + " itens") : (posicaoRgb + 1) + " de " + n;
        }

        // v3.6: seleciona o item da lista filtrada e da zoom; do ultimo volta ao primeiro
        void IrParaRgb(int indice)
        {
            int n = filtradosRgb.Count;
            if (n == 0) return;
            if (indice < 0) indice = n - 1;
            if (indice >= n) indice = 0;
            int anterior = posicaoRgb;
            posicaoRgb = indice;
            if (anterior >= 0 && anterior < linhasRgb.Count) linhasRgb[anterior].Background = LinhaRgb;
            if (indice < linhasRgb.Count) linhasRgb[indice].Background = LinhaRgbAtual;
            AtualizarPosicaoRgb();
            SelecionarRgb(filtradosRgb[indice]);
        }

        // item == null: seleciona todos os itens da lista (o objeto de fora de cada um)
        void SelecionarRgb(ItemRgb item)
        {
            ResultadoRgb r = resultadoRgb;
            if (r == null) return;
            try
            {
                if (corel.ChaveAtual() != r.Chave) { status.Text = "A página mudou. Clique em Atualizar página."; status.Foreground = Suave; return; }
                if (ForaDaEdicao()) VoltarEdicao(); // v3.6: so mexe nas chapas quando a previa esta aberta ou alguma chapa desligada
                string msg = item == null
                    ? VerificadorRgb.SelecionarTodos(corel.Aplicacao, corel.Documento(), filtradosRgb)
                    : VerificadorRgb.SelecionarItem(corel.Aplicacao, corel.Documento(), item);
                if (item != null && posicaoRgb >= 0 && posicaoRgb < filtradosRgb.Count && ReferenceEquals(filtradosRgb[posicaoRgb], item))
                    msg = (posicaoRgb + 1) + " de " + filtradosRgb.Count + ": " + item.Descricao + "\n" + msg;
                status.Text = msg;
                status.Foreground = Suave;
            }
            catch (Exception ex) { Corel.Registrar(ex); status.Text = "Não foi possível selecionar: " + Corel.Explicar(ex); }
        }

        // v3.6: a previa sobre a arte esta aberta ou alguma chapa esta desligada
        bool ForaDaEdicao()
        {
            return tela != null || !visivel[0] || !visivel[1] || !visivel[2] || !visivel[3] ||
                Array.IndexOf(especiaisVisiveis, false) >= 0;
        }

        void VoltarEdicao()
        {
            telaSuspensa = true; FecharTela();
            for (int i = 0; i < 4; i++) visivel[i] = true;
            for (int i = 0; i < especiaisVisiveis.Length; i++) DefinirEspecial(i, true);
            AtualizarOlhos(); Redesenhar();
        }

        void AtualizarTela()
        {
            if (!PrecisaTela() || imagem == null || previa.Source == null) { FecharTela(); return; }
            try
            {
                if (corel.ChaveAtual() != chaveAtual) { telaSuspensa = true; FecharTela(); return; }
                VistaDocumento vista = corel.LerVista();
                if (tela != null && !tela.Conferir(vista)) FecharTela();
                else if (tela != null && tela.Alinhar(vista)) // v2.8: pagina mudou de lugar/tamanho no Corel
                { detalhe = null; detVersao = -1; geracaoDetalhe++; ReiniciarNitidez(); }
                if (tela == null) { tela = new PreviaNaArte(vista, VoltarEdicao, Atualizar, ReiniciarNitidez, MedirNaArte); ReiniciarNitidez(); }
                tela.Exibir((BitmapSource)previa.Source);
            }
            catch (Exception ex)
            {
                telaSuspensa = true; FecharTela(); Corel.Registrar(ex);
                status.Text = "Não foi possível mostrar sobre a arte: " + Corel.Explicar(ex);
                status.Foreground = Pincel(255, 196, 135);
            }
        }

        void Redesenhar()
        {
            geracaoRender++;
            if (imagem == null || !ativo) return;
            if (renderizando) { renderPendente = true; return; }
            renderizando = true; renderPendente = false;
            int versao = geracaoRender;
            ImagemCmyk img = imagem;
            bool[] chapas = (bool[])visivel.Clone();
            bool[] especiais = (bool[])especiaisVisiveis.Clone();
            bool preto = false, carga = false; int pct = 300; // v3.6: opcoes removidas do painel
            string perfil = perfilCmyk; // v2.6: cores pelo perfil CMYK
            Task.Factory.StartNew(delegate { return CorDeTela.Desenhar(img, chapas, preto, carga, pct, especiais, perfil); })
                .ContinueWith(delegate(Task<byte[]> tarefa)
                {
                    Exception falha = tarefa.IsFaulted ? tarefa.Exception : null;
                    if (Dispatcher.HasShutdownStarted) return;
                    Dispatcher.BeginInvoke(new Action(delegate
                    {
                        renderizando = false;
                        if (ativo && versao == geracaoRender && img == imagem)
                        {
                            try
                            {
                                if (falha != null) Erro(falha, true);
                                else
                                {
                                    BitmapSource bitmap = BitmapSource.Create(img.Largura, img.Altura, 96, 96,
                                        PixelFormats.Bgra32, null, tarefa.Result, img.Largura * 4);
                                    bitmap.Freeze(); previa.Source = bitmap; vazio.Visibility = Visibility.Collapsed;
                                    AtualizarTela();
                                    DesenharDetalhe(); // v2.4
                                }
                            }
                            catch (Exception ex) { Erro(ex, true); }
                        }
                        if (renderPendente && ativo) Redesenhar();
                    }), DispatcherPriority.Background);
                }, TaskScheduler.Default);
        }

        void MouseNaPrevia(object sender, MouseEventArgs e)
        {
            if (imagem == null || previa.Source == null) return;
            Point p = e.GetPosition(previa); int x, y;
            if (!Renderizador.LocalizarPixel(p.X, p.Y, previa.ActualWidth, previa.ActualHeight,
                                            imagem.Largura, imagem.Altura, out x, out y))
            { LimparLeitura(); return; }
            int[] v = imagem.Amostra(x, y);
            for (int i = 0; i < 4; i++) percentuais[i].Text = v[i] + "%";
            for (int i = 0; i < imagem.Especiais.Length; i++)
                percentuaisEspeciais[i].Text = ImagemCmyk.Percentual(imagem.Especiais[i].Cobertura[y * imagem.Largura + x]) + "%";
            leitura.Text = String.Format("C {0}%   M {1}%   Y {2}%   K {3}%\nCarga CMYK: {4}% · leitura inclui tintas ocultas", v[0], v[1], v[2], v[3], v[4]);
        }

        // v2.5: leitura das tintas com o mouse sobre a arte (usa o detalhe quando ele cobre o ponto).
        void MedirNaArte(double u, double v)
        {
            if (imagem == null || u < 0 || v < 0 || u >= 1 || v >= 1) { LimparLeitura(); return; }
            ImagemCmyk img = imagem;
            double fu = u, fv = v;
            if (detalhe != null && detalhe.Especiais.Length == imagem.Especiais.Length &&
                u >= detU0 && u < detU1 && v >= detV0 && v < detV1)
            {
                img = detalhe;
                fu = (u - detU0) / (detU1 - detU0);
                fv = (v - detV0) / (detV1 - detV0);
            }
            int x = Math.Max(0, Math.Min(img.Largura - 1, (int)(fu * img.Largura)));
            int y = Math.Max(0, Math.Min(img.Altura - 1, (int)(fv * img.Altura)));
            int[] amostra = img.Amostra(x, y);
            if (amostra == null) { LimparLeitura(); return; }
            for (int i = 0; i < 4; i++) percentuais[i].Text = amostra[i] + "%";
            for (int i = 0; i < img.Especiais.Length && i < percentuaisEspeciais.Count; i++)
                percentuaisEspeciais[i].Text = ImagemCmyk.Percentual(img.Especiais[i].Cobertura[y * img.Largura + x]) + "%";
            leitura.Text = String.Format("C {0}%   M {1}%   Y {2}%   K {3}%\nCarga CMYK: {4}% · leitura inclui tintas ocultas",
                amostra[0], amostra[1], amostra[2], amostra[3], amostra[4]);
        }

        void LimparLeitura()
        {
            for (int i = 0; i < 4; i++) if (percentuais[i] != null) percentuais[i].Text = "—";
            foreach (TextBlock pct in percentuaisEspeciais) pct.Text = "—";
            leitura.Text = "Desligue uma chapa e passe o mouse na arte para medir as tintas.";
        }
    }
#endif
}
