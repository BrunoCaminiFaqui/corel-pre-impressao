// TrocarCoresCorel - painel "Trocar cores" para o CorelDRAW. by Bruno Faqui
// Lista as cores usadas na arte (inclusive dentro de grupos e PowerClips), seleciona os objetos
// de uma cor e troca uma cor por outra num passo so de desfazer.
// Compilar como DLL. O Corel entrega sua propria Application ao construtor DockerUI(object app).
// As cores sao comparadas pelos valores (modelo + numeros), sem depender do Color.IsSame.
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace TrocarCoresCorel
{
    internal static class Registro
    {
        static string Pasta
        {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Imprell", "TrocarCores"); }
        }

        internal static void Erro(Exception ex) { Gravar("erros.log", ex.ToString(), 1024 * 1024); }

        internal static void Diagnostico(string texto) { Gravar("diagnostico.log", texto, 512 * 1024); }

        static void Gravar(string nome, string texto, long limite)
        {
            try
            {
                Directory.CreateDirectory(Pasta);
                string arquivo = Path.Combine(Pasta, nome);
                if (File.Exists(arquivo) && new FileInfo(arquivo).Length > limite) File.Delete(arquivo);
                File.AppendAllText(arquivo, DateTime.Now.ToString("s") + " " + texto + Environment.NewLine);
            }
            catch { } // o registro nunca deve derrubar o Corel
        }

        internal static string Explicar(Exception ex)
        {
            COMException com = ex as COMException;
            if (com != null && (com.ErrorCode == unchecked((int)0x80010001) || com.ErrorCode == unchecked((int)0x8001010A)))
                return "O Corel está ocupado. Feche a caixa de diálogo e tente de novo.";
            return ex.Message;
        }
    }

    internal sealed class Corel
    {
        // Referencia emprestada do host: nunca liberar/fechar a Application do Corel.
        readonly object aplicacao;
        internal Corel(object app)
        {
            if (app == null) throw new ArgumentNullException("app");
            aplicacao = app;
        }

        internal object Aplicacao { get { return aplicacao; } }

        internal object Documento()
        {
            dynamic app = aplicacao;
            if ((int)app.Documents.Count == 0) return null;
            return (object)app.ActiveDocument;
        }

        // documento + pagina ativa
        internal string ChaveAtual()
        {
            object doc = Documento();
            if (doc == null) return null;
            IntPtr ptr = Marshal.GetIUnknownForObject(doc);
            try { return ptr.ToInt64().ToString("X") + ":" + ((dynamic)doc).ActivePage.Index; }
            finally { Marshal.Release(ptr); }
        }
    }

    internal enum Parte { Preench, Degrade, Contorno, Texto }

    // Cor lida de um objeto: modelo e valores. A Chave identifica a cor.
    internal sealed class InfoCor
    {
        internal int Tipo;
        internal string Modelo;
        internal int[] Numeros;      // CMYK, CMY, RGB ou Cinza
        internal string Descricao;
        internal string Chave { get { return Tipo.ToString(CultureInfo.InvariantCulture) + "|" + Descricao; } }
    }

    internal sealed class CorUsada
    {
        internal string Chave, Modelo, Descricao;
        internal object Amostra;     // copia da cor, para abrir a janela de cores com ela
        internal int Indice;
        internal byte R = 128, G = 128, B = 128;
        internal int Preench, Degrade, Contorno, Texto;
        internal int Total { get { return Preench + Degrade + Contorno + Texto; } }
        internal bool ForaDoCmyk
        {
            get { return Modelo == "RGB" || Modelo == "HSB" || Modelo == "HLS" || Modelo == "Lab" || Modelo == "YIQ"; }
        }
    }

    // v1.3: imagem fora do CMYK (RGB, Lab ou em paleta) para localizar pelo painel
    internal sealed class ImagemAchada
    {
        internal object Forma, Topo, Pagina; // Pagina null: pagina mestre
        internal List<object> Molduras;      // PowerClips por fora da imagem (de fora para dentro)
        internal int NumeroPagina;
        internal string Descricao;
    }

    internal sealed class Leitura
    {
        internal readonly List<CorUsada> Cores = new List<CorUsada>();
        internal readonly List<ImagemAchada> ImagensFora = new List<ImagemAchada>(); // v1.3
        internal int ImagensForaTotal;                                                // v1.3
        internal int Objetos, Imagens, ComEfeito, ForaDaPagina, PowerClips;
        internal bool Parcial, TodasPaginas;
        internal string ChaveCorel;
    }

    internal static class Motor
    {
        // Mistura, extrusao, contour, chanfro, midia artistica e efeito personalizado: os passos gerados
        // pelo Corel nao sao lidos (cada passo teria uma cor intermediaria).
        static readonly int[] Efeitos = { 10, 11, 13, 15, 18, 22 };
        static readonly TimeSpan LimiteTempo = TimeSpan.FromSeconds(8);
        const int LimiteLetras = 5000, LimiteCores = 300, LimiteImagens = 3000;
        static readonly HashSet<int> modosVistos = new HashSet<int>(); // v1.3: modos de imagem ja anotados
        static readonly Dictionary<int, string> modelos = new Dictionary<int, string>();
        static int textoCor = -1; // Color.ToString do Corel: -1 nao testado, 0 nao funciona, 1 funciona
        static int falhas;
        static int diagPowerClip;
        // v1.1: objetos fora de PowerClip (mesma lista para todos, para agrupar a selecao)
        static readonly List<object> SemMolduras = new List<object>();
        // PowerClips que o painel abriu (modo de edicao) para selecionar o conteudo
        static List<object> moldurasAbertas = new List<object>();

        abstract class Visita
        {
            internal bool UsarPreench = true, UsarContorno = true;
            internal bool Trocando;          // destrava, aplica a cor nova e trava de novo
            internal object Nova;
            internal InfoCor NovaInfo;
            internal int Objetos, Imagens, ComEfeito, ForaDaPagina, PowerClips;
            internal int Trocas, Recusadas, ObjetosTrocados;
            internal bool Parar;
            internal object FormaAtual, TopoAtual, UltimaTrocada;
            internal List<object> Molduras = SemMolduras; // PowerClips por fora do objeto atual (de fora para dentro)
            internal object PaginaAtual;   // v1.3: null na pagina mestre
            internal int NumeroPaginaAtual;
            internal readonly System.Diagnostics.Stopwatch Relogio = System.Diagnostics.Stopwatch.StartNew();
            // true = trocar esta cor
            internal abstract bool Ver(InfoCor info, object cor, Parte parte);
            // v1.3: imagem encontrada (so a leitura guarda)
            internal virtual void VerImagem(object forma) { }
        }

        sealed class VisitaLeitura : Visita
        {
            internal readonly Leitura Resultado = new Leitura();
            readonly object app;
            readonly Dictionary<string, CorUsada> porChave = new Dictionary<string, CorUsada>();
            readonly HashSet<long> contadas = new HashSet<long>();
            object formaContada;
            internal VisitaLeitura(object app) { this.app = app; }

            internal override bool Ver(InfoCor info, object cor, Parte parte)
            {
                CorUsada u;
                if (!porChave.TryGetValue(info.Chave, out u))
                {
                    if (Resultado.Cores.Count >= LimiteCores) return false;
                    u = NovaEntrada(info, cor, app);
                    u.Indice = Resultado.Cores.Count;
                    porChave[info.Chave] = u;
                    Resultado.Cores.Add(u);
                }
                // cada objeto conta uma vez por cor e por parte (texto: uma vez, nao por letra)
                if (!ReferenceEquals(formaContada, FormaAtual)) { formaContada = FormaAtual; contadas.Clear(); }
                if (contadas.Add(((long)u.Indice << 3) | (long)(int)parte))
                {
                    switch (parte)
                    {
                        case Parte.Preench: u.Preench++; break;
                        case Parte.Degrade: u.Degrade++; break;
                        case Parte.Contorno: u.Contorno++; break;
                        default: u.Texto++; break;
                    }
                }
                return false;
            }

            // v1.3: guarda as imagens RGB/Lab/paleta para a lista do painel
            internal override void VerImagem(object forma)
            {
                string modo = ModoForaDoCmyk(forma);
                if (modo == null) return;
                Resultado.ImagensForaTotal++;
                if (Resultado.ImagensFora.Count >= LimiteImagens) return;
                string onde = Molduras.Count > 0 ? " · dentro de PowerClip" : !ReferenceEquals(forma, TopoAtual) ? " · em grupo" : "";
                Resultado.ImagensFora.Add(new ImagemAchada { Forma = forma, Topo = TopoAtual, Molduras = Molduras,
                    Pagina = PaginaAtual, NumeroPagina = NumeroPaginaAtual, Descricao = "Imagem " + modo + DetalheImagem(forma) + onde });
            }
        }

        sealed class VisitaTroca : Visita
        {
            readonly string alvo;
            internal VisitaTroca(string alvo) { this.alvo = alvo; Trocando = true; }
            internal override bool Ver(InfoCor info, object cor, Parte parte) { return info.Chave == alvo; }
        }

        // v1.4: converte para CMYK as cores fora do CMYK (RGB, HSB, HLS, Lab, YIQ) e as imagens RGB/Lab/paleta.
        // Cada cor e convertida uma vez pelo proprio Corel (ConvertToCMYK, gerenciamento de cores do documento)
        // e o mesmo resultado vai para todos os lugares onde ela aparece.
        sealed class VisitaConversao : Visita
        {
            readonly Dictionary<string, InfoCor> convertidas = new Dictionary<string, InfoCor>();
            readonly Dictionary<string, object> novas = new Dictionary<string, object>();
            internal readonly HashSet<string> Cores = new HashSet<string>();
            internal int ImagensConvertidas, ImagensRecusadas, CoresRecusadas;
            internal VisitaConversao() { Trocando = true; }

            internal override bool Ver(InfoCor info, object cor, Parte parte)
            {
                if (!ForaDoCmyk(info.Modelo)) return false;
                InfoCor nova;
                if (!convertidas.TryGetValue(info.Chave, out nova))
                {
                    object copia = ParaCmyk(cor);
                    nova = copia == null ? null : LerCor(copia);
                    if (nova != null && nova.Modelo != "CMYK") nova = null;
                    convertidas[info.Chave] = nova;
                    novas[info.Chave] = copia;
                    Registro.Diagnostico("Conversao " + info.Chave + " -> " + (nova != null ? nova.Chave : "recusada"));
                }
                if (nova == null) { CoresRecusadas++; return false; }
                Nova = novas[info.Chave]; NovaInfo = nova;
                Cores.Add(info.Chave);
                return true;
            }

            internal override void VerImagem(object forma)
            {
                if (ModoForaDoCmyk(forma) == null) return;
                if (ImagemParaCmyk(forma)) ImagensConvertidas++; else ImagensRecusadas++;
            }
        }

        static bool ForaDoCmyk(string modelo)
        {
            return modelo == "RGB" || modelo == "HSB" || modelo == "HLS" || modelo == "Lab" || modelo == "YIQ";
        }

        static object ParaCmyk(object cor)
        {
            try
            {
                dynamic copia = ((dynamic)cor).GetCopy();
                copia.ConvertToCMYK();
                return (object)copia;
            }
            catch (Exception ex) { Falha(ex); return null; }
        }

        static int imagensAnotadas;

        // 5 = cdrCMYKColorImage; confere lendo o modo de novo
        static bool ImagemParaCmyk(object forma)
        {
            int antes = -1, depois = -1;
            try { antes = (int)((dynamic)forma).Bitmap.Mode; } catch (Exception ex) { Falha(ex); }
            try { ((dynamic)forma).Bitmap.ConvertTo(5); } catch (Exception ex) { Falha(ex); }
            try { depois = (int)((dynamic)forma).Bitmap.Mode; } catch (Exception ex) { Falha(ex); }
            if (imagensAnotadas++ < 10) Registro.Diagnostico("Imagem para CMYK: modo " + antes + " -> " + depois);
            return depois == 5;
        }

        // Objetos com a cor, agrupados pelo PowerClip em que estao (ou fora de PowerClip)
        sealed class GrupoSelecao
        {
            internal List<object> Molduras;
            internal readonly List<object> Formas = new List<object>();
            internal readonly List<object> Topos = new List<object>();
        }

        sealed class VisitaSelecao : Visita
        {
            readonly string alvo;
            internal readonly List<GrupoSelecao> Grupos = new List<GrupoSelecao>();
            object ultima;
            internal VisitaSelecao(string alvo) { this.alvo = alvo; }
            internal override bool Ver(InfoCor info, object cor, Parte parte)
            {
                if (info.Chave != alvo || ReferenceEquals(ultima, FormaAtual)) return false;
                ultima = FormaAtual;
                GrupoSelecao g = null;
                foreach (GrupoSelecao x in Grupos) if (ReferenceEquals(x.Molduras, Molduras)) { g = x; break; }
                if (g == null) { g = new GrupoSelecao { Molduras = Molduras }; Grupos.Add(g); }
                g.Formas.Add(FormaAtual);
                g.Topos.Add(TopoAtual);
                return false;
            }
        }

        // ---------------- uso pelo painel ----------------

        internal static Leitura Ler(object app, object documento, bool todas, bool preench, bool contorno)
        {
            falhas = 0; diagPowerClip = 0;
            SairDasMolduras();
            VisitaLeitura v = new VisitaLeitura(app) { UsarPreench = preench, UsarContorno = contorno };
            Percorrer(documento, todas, v);
            Leitura r = v.Resultado;
            r.Objetos = v.Objetos; r.Imagens = v.Imagens; r.ComEfeito = v.ComEfeito; r.ForaDaPagina = v.ForaDaPagina;
            r.PowerClips = v.PowerClips;
            r.Parcial = v.Parar; r.TodasPaginas = todas;
            Registro.Diagnostico("Leitura: " + r.Cores.Count + " cores, " + r.Objetos + " objetos, " + r.PowerClips + " PowerClips, " + r.Imagens + " imagens (" +
                r.ImagensForaTotal + " fora do CMYK), " +
                r.ComEfeito + " com efeito, " + r.ForaDaPagina + " fora da pagina, " + v.Relogio.ElapsedMilliseconds + " ms" +
                (r.Parcial ? " (parcial)" : ""));
            for (int i = 0; i < r.Cores.Count && i < 40; i++)
                Registro.Diagnostico("  " + r.Cores[i].Chave + "  x" + r.Cores[i].Total);
            return r;
        }

        // v1.1: seleciona sem passar objetos para o Corel (so AddToSelection em cada um).
        // Objetos dentro de PowerClip: entra no PowerClip (como "Editar PowerClip") e seleciona la dentro.
        // Com a cor em varios lugares, cada clique seleciona o proximo grupo (fora de PowerClip primeiro).
        internal static string Selecionar(object app, object documento, string chave, bool preench, bool contorno,
            int pedido, out int totalGrupos)
        {
            falhas = 0;
            SairDasMolduras();
            VisitaSelecao v = new VisitaSelecao(chave) { UsarPreench = preench, UsarContorno = contorno };
            Percorrer(documento, false, v);
            List<GrupoSelecao> ordem = new List<GrupoSelecao>();
            foreach (GrupoSelecao g in v.Grupos) if (g.Molduras.Count == 0) ordem.Add(g);
            foreach (GrupoSelecao g in v.Grupos) if (g.Molduras.Count > 0) ordem.Add(g);
            totalGrupos = ordem.Count;
            if (ordem.Count == 0) return "Essa cor não está nesta página (pode estar em outra página).";
            int indice = pedido % ordem.Count;
            GrupoSelecao grupo = ordem[indice];
            dynamic doc = documento;

            // entra nos PowerClips, de fora para dentro
            List<object> abertas = new List<object>();
            bool entrou = true;
            foreach (object m in grupo.Molduras)
            {
                try { ((dynamic)m).PowerClip.EnterEditMode(); abertas.Add(m); }
                catch (Exception ex) { Falha(ex); entrou = false; break; }
            }
            if (!entrou)
            {
                for (int i = abertas.Count - 1; i >= 0; i--) { try { ((dynamic)abertas[i]).PowerClip.LeaveEditMode(); } catch { } }
                abertas.Clear();
            }
            moldurasAbertas = abertas;

            try { doc.ClearSelection(); } catch (Exception ex) { Falha(ex); }
            int ok = 0;
            if (entrou)
            {
                for (int i = 0; i < grupo.Formas.Count; i++)
                {
                    bool foi = false;
                    try { ((dynamic)grupo.Formas[i]).AddToSelection(); foi = true; } catch (Exception ex) { Falha(ex); }
                    // fora de PowerClip, se nao deu no objeto (ex.: dentro de grupo), seleciona o objeto de fora
                    if (!foi && grupo.Molduras.Count == 0)
                    {
                        try { ((dynamic)grupo.Topos[i]).AddToSelection(); foi = true; } catch (Exception ex) { Falha(ex); }
                    }
                    if (foi) ok++;
                }
            }
            Registro.Diagnostico("Selecao " + chave + ": grupo " + (indice + 1) + "/" + ordem.Count + ", " + grupo.Formas.Count +
                " objetos, " + grupo.Molduras.Count + " PowerClip(s) por fora, entrou=" + entrou + ", selecionados=" + ok);

            string proximo = ordem.Count > 1
                ? " (" + (indice + 1) + " de " + ordem.Count + "; clique em Selecionar de novo para o próximo)" : "";
            if (ok == 0 && grupo.Molduras.Count > 0)
            {
                try { ((dynamic)grupo.Molduras[0]).AddToSelection(); } catch (Exception ex) { Falha(ex); }
                return "Não deu para entrar no PowerClip: ficou selecionado o PowerClip de fora, com " + grupo.Formas.Count +
                       (grupo.Formas.Count == 1 ? " objeto" : " objetos") + " dessa cor lá dentro" + proximo + ".";
            }
            if (ok == 0) return "Não foi possível selecionar. A camada pode estar travada.";
            string msg = ok + (ok == 1 ? " objeto selecionado" : " objetos selecionados") +
                         (grupo.Molduras.Count > 0 ? " dentro do PowerClip" : "") + proximo + ".";
            if (grupo.Molduras.Count > 0) msg += " Para sair do PowerClip, clique em Concluir no Corel.";
            return msg;
        }

        static void SairDasMolduras()
        {
            for (int i = moldurasAbertas.Count - 1; i >= 0; i--)
                try { ((dynamic)moldurasAbertas[i]).PowerClip.LeaveEditMode(); } catch { }
            moldurasAbertas = new List<object>();
        }

        // v1.3: seleciona uma imagem (vai para a pagina dela e entra nos PowerClips por fora) e da zoom
        internal static string SelecionarImagem(object app, object documento, ImagemAchada img)
        {
            falhas = 0;
            SairDasMolduras();
            dynamic doc = documento;
            string pagina = "";
            if (img.Pagina != null && img.NumeroPagina > 0)
            {
                try
                {
                    if ((int)doc.ActivePage.Index != img.NumeroPagina)
                    {
                        ((dynamic)img.Pagina).Activate();
                        pagina = " (página " + img.NumeroPagina + ")";
                    }
                }
                catch (Exception ex) { Falha(ex); }
            }
            List<object> abertas = new List<object>();
            bool entrou = true;
            foreach (object m in img.Molduras)
            {
                try { ((dynamic)m).PowerClip.EnterEditMode(); abertas.Add(m); }
                catch (Exception ex) { Falha(ex); entrou = false; break; }
            }
            moldurasAbertas = abertas;
            try { doc.ClearSelection(); } catch (Exception ex) { Falha(ex); }
            bool ok = false;
            string msg;
            if (entrou)
            {
                try { ((dynamic)img.Forma).AddToSelection(); ok = true; } catch (Exception ex) { Falha(ex); }
            }
            if (ok) msg = img.Molduras.Count > 0 ? "Imagem selecionada dentro do PowerClip" + pagina + ". Para sair, clique em Concluir no Corel."
                                                 : "Imagem selecionada" + pagina + ".";
            else
            {
                // nao deu para ir direto: seleciona o objeto de fora (PowerClip ou grupo)
                SairDasMolduras();
                object fora = img.Molduras.Count > 0 ? img.Molduras[0] : img.Topo;
                try { doc.ClearSelection(); } catch { }
                try { ((dynamic)fora).AddToSelection(); ok = true; } catch (Exception ex) { Falha(ex); }
                msg = ok ? "Não deu para ir direto na imagem: ficou selecionado o objeto de fora" + pagina + "."
                         : "Não foi possível selecionar. A camada pode estar travada ou a arte mudou: clique em Ler cores.";
            }
            if (ok) { try { ((dynamic)app).ActiveWindow.ActiveView.ToFitSelection(); } catch (Exception ex) { Falha(ex); } }
            Registro.Diagnostico("Selecao de imagem: " + img.Molduras.Count + " PowerClip(s) por fora, entrou=" + entrou + ", ok=" + ok);
            return msg;
        }

        // v1.3: modo da imagem (cdrImageType): 4 e 9 RGB, 7 Lab, 1 e 3 paleta (igual ao painel de separacoes)
        static string ModoForaDoCmyk(object forma)
        {
            try
            {
                int modo = (int)((dynamic)forma).Bitmap.Mode;
                if (modosVistos.Add(modo)) Registro.Diagnostico("Imagem com modo " + modo);
                if (modo == 4 || modo == 9) return "RGB";
                if (modo == 7) return "Lab";
                if (modo == 1 || modo == 3) return "em paleta (RGB)";
            }
            catch (Exception ex) { Falha(ex); }
            return null;
        }

        // v1.3: tamanho em pixels e arquivo vinculado, para reconhecer a imagem na lista
        static string DetalheImagem(object forma)
        {
            string d = "";
            try
            {
                dynamic b = ((dynamic)forma).Bitmap;
                int w = (int)b.SizeWidth, h = (int)b.SizeHeight;
                if (w > 0 && h > 0) d += " · " + w + "×" + h + " px";
            }
            catch { }
            try
            {
                dynamic b = ((dynamic)forma).Bitmap;
                if ((bool)b.ExternallyLinked)
                {
                    string arquivo = (string)b.LinkFileName;
                    if (!String.IsNullOrEmpty(arquivo)) d += " · vínculo " + Path.GetFileName(arquivo);
                }
            }
            catch { }
            return d;
        }

        internal static string Trocar(object app, object documento, string chave, object nova, bool todas, bool preench, bool contorno)
        {
            falhas = 0;
            InfoCor info = LerCor(nova);
            if (info == null) return "Não foi possível ler a cor escolhida. Tente de novo.";
            if (info.Chave == chave) return "A nova cor é igual à antiga. Nada foi trocado.";
            SairDasMolduras();
            VisitaTroca v = new VisitaTroca(chave) { UsarPreench = preench, UsarContorno = contorno, Nova = nova, NovaInfo = info };
            dynamic a = app;
            dynamic doc = documento;
            bool grupoAberto = false, otimizando = false;
            try
            {
                doc.BeginCommandGroup("Trocar cor");
                grupoAberto = true;
                try { a.Optimization = true; otimizando = true; } catch (Exception ex) { Falha(ex); }
                Percorrer(documento, todas, v);
            }
            finally
            {
                if (otimizando) { try { a.Optimization = false; } catch (Exception ex) { Falha(ex); } }
                if (grupoAberto) { try { doc.EndCommandGroup(); } catch (Exception ex) { Falha(ex); } }
                try { a.ActiveWindow.Refresh(); } catch { }
                try { a.Refresh(); } catch { }
            }
            Registro.Diagnostico("Troca " + chave + " -> " + info.Chave + ": " + v.Trocas + " trocadas em " + v.ObjetosTrocados +
                " objetos, " + v.Recusadas + " recusadas");
            if (v.Trocas == 0)
                return v.Recusadas > 0
                    ? "Nada foi trocado: o Corel não aceitou a cor nova em " + v.Recusadas + " lugar(es). Mande o arquivo diagnostico.log."
                    : "Essa cor não foi encontrada. A arte pode ter mudado: clique em Ler cores.";
            string msg = "Cor trocada em " + v.ObjetosTrocados + (v.ObjetosTrocados == 1 ? " objeto" : " objetos") +
                         " (" + v.Trocas + (v.Trocas == 1 ? " lugar" : " lugares") + "). Ctrl+Z desfaz tudo.";
            if (v.Recusadas > 0) msg += " Atenção: " + v.Recusadas + " lugar(es) não aceitaram a troca.";
            return msg;
        }

        // v1.4: converte para CMYK tudo o que esta fora do CMYK (cores e imagens), preenchimento e contorno
        internal static string ConverterParaCmyk(object app, object documento, bool todas)
        {
            falhas = 0; imagensAnotadas = 0;
            SairDasMolduras();
            VisitaConversao v = new VisitaConversao();
            dynamic a = app;
            dynamic doc = documento;
            bool grupoAberto = false, otimizando = false;
            try
            {
                doc.BeginCommandGroup("Converter RGB para CMYK");
                grupoAberto = true;
                try { a.Optimization = true; otimizando = true; } catch (Exception ex) { Falha(ex); }
                Percorrer(documento, todas, v);
            }
            finally
            {
                if (otimizando) { try { a.Optimization = false; } catch (Exception ex) { Falha(ex); } }
                if (grupoAberto) { try { doc.EndCommandGroup(); } catch (Exception ex) { Falha(ex); } }
                try { a.ActiveWindow.Refresh(); } catch { }
                try { a.Refresh(); } catch { }
            }
            Registro.Diagnostico("Conversao para CMYK: " + v.Cores.Count + " cores em " + v.ObjetosTrocados + " objetos (" + v.Trocas +
                " lugares, " + v.Recusadas + " recusados, " + v.CoresRecusadas + " sem conversao); imagens " + v.ImagensConvertidas +
                " convertidas, " + v.ImagensRecusadas + " recusadas");

            if (v.Trocas == 0 && v.ImagensConvertidas == 0 && v.Recusadas == 0 && v.CoresRecusadas == 0 && v.ImagensRecusadas == 0)
                return "Nada fora do CMYK" + (todas ? " no arquivo." : " nesta página.");
            List<string> feito = new List<string>();
            if (v.Trocas > 0)
                feito.Add(v.Cores.Count + (v.Cores.Count == 1 ? " cor" : " cores") + " em " + v.ObjetosTrocados +
                    (v.ObjetosTrocados == 1 ? " objeto" : " objetos"));
            if (v.ImagensConvertidas > 0)
                feito.Add(v.ImagensConvertidas + (v.ImagensConvertidas == 1 ? " imagem" : " imagens"));
            string msg = feito.Count > 0 ? "Convertido para CMYK: " + String.Join(" e ", feito.ToArray()) + ". Ctrl+Z desfaz tudo."
                                         : "Nada foi convertido.";
            if (v.Recusadas + v.CoresRecusadas > 0)
                msg += "\nAtenção: " + (v.Recusadas + v.CoresRecusadas) + " lugar(es) com cor não aceitaram a conversão.";
            if (v.ImagensRecusadas > 0)
                msg += "\nAtenção: " + v.ImagensRecusadas + (v.ImagensRecusadas == 1 ? " imagem não converteu" : " imagens não converteram") +
                       " (vinculada ou travada?). Ela continua na lista de imagens.";
            return msg;
        }

        // ---------------- percurso da arte ----------------

        static void Percorrer(object documento, bool todas, Visita v)
        {
            dynamic doc = documento;
            object ativa = doc.ActivePage;
            List<object> paginas = new List<object>();
            if (todas)
            {
                dynamic pgs = doc.Pages;
                int n = (int)pgs.Count;
                for (int i = 1; i <= n; i++) paginas.Add((object)pgs.Item(i));
            }
            else paginas.Add(ativa);
            object mestre = null;
            try { mestre = doc.MasterPage; } catch (Exception ex) { Falha(ex); }
            try
            {
                foreach (object pagina in paginas)
                {
                    if (v.Parar) break;
                    if (v.Trocando && todas) { try { ((dynamic)pagina).Activate(); } catch (Exception ex) { Falha(ex); } }
                    v.PaginaAtual = pagina; // v1.3
                    try { v.NumeroPaginaAtual = (int)((dynamic)pagina).Index; } catch { v.NumeroPaginaAtual = 0; }
                    PercorrerPagina(pagina, pagina, false, v);
                }
                // camadas mestre: aparecem em todas as paginas, entram uma vez
                if (mestre != null && !v.Parar)
                {
                    v.PaginaAtual = null; v.NumeroPaginaAtual = 0;
                    PercorrerPagina(mestre, ativa, true, v);
                }
            }
            finally
            {
                if (v.Trocando && todas) { try { ((dynamic)ativa).Activate(); } catch (Exception ex) { Falha(ex); } }
            }
        }

        static void PercorrerPagina(object pagina, object paginaDosLimites, bool ehMestre, Visita v)
        {
            double pl = 0, pr = 0, pt = 0, pb = 0;
            bool temLimites = false;
            try
            {
                dynamic p = paginaDosLimites;
                pl = (double)p.LeftX; pr = (double)p.RightX; pt = (double)p.TopY; pb = (double)p.BottomY;
                temLimites = pr > pl && pt > pb;
            }
            catch (Exception ex) { Falha(ex); }
            dynamic camadas = ((dynamic)pagina).Layers;
            int n = (int)camadas.Count;
            for (int i = 1; i <= n && !v.Parar; i++)
            {
                object c = camadas.Item(i);
                if (Logico(c, "IsDesktopLayer", false) || Logico(c, "IsGuidesLayer", false) || Logico(c, "IsGridLayer", false)) continue;
                if (!ehMestre && Logico(c, "Master", false)) continue; // camada mestre entra pela pagina mestre
                if (!Logico(c, "Visible", true)) continue;            // camada oculta
                bool destravada = false;
                if (v.Trocando && !Logico(c, "Editable", true))
                {
                    try { ((dynamic)c).Editable = true; destravada = true; } catch (Exception ex) { Falha(ex); }
                }
                try
                {
                    dynamic formas = ((dynamic)c).Shapes;
                    int m = (int)formas.Count;
                    List<object> lista = new List<object>(m); // lista fixa antes de mexer
                    for (int k = 1; k <= m; k++) lista.Add((object)formas.Item(k));
                    foreach (object s in lista)
                    {
                        if (v.Parar) break;
                        if (temLimites && !NaPagina(s, pl, pr, pt, pb)) { v.ForaDaPagina++; continue; }
                        v.TopoAtual = s;
                        Visitar(s, SemMolduras, v);
                    }
                }
                catch (Exception ex) { Falha(ex); }
                finally
                {
                    if (destravada) { try { ((dynamic)c).Editable = false; } catch (Exception ex) { Falha(ex); } }
                }
            }
        }

        static bool NaPagina(object forma, double pl, double pr, double pt, double pb)
        {
            try
            {
                dynamic s = forma;
                double l = (double)s.LeftX, r = (double)s.RightX, t = (double)s.TopY, b = (double)s.BottomY;
                return r >= pl && l <= pr && t >= pb && b <= pt;
            }
            catch { return true; }
        }

        static void Visitar(object forma, List<object> molduras, Visita v)
        {
            if (v.Parar) return;
            if (!v.Trocando && v.Relogio.Elapsed > LimiteTempo) { v.Parar = true; return; }
            dynamic s = forma;
            int tipo;
            try { tipo = (int)s.Type; } catch (Exception ex) { Falha(ex); return; }
            if (tipo == 9) return; // guia
            bool travado = false;
            if (v.Trocando)
            {
                try { travado = (bool)s.Locked; if (travado) s.Locked = false; }
                catch (Exception ex) { Falha(ex); travado = false; }
            }
            try
            {
                v.Objetos++;
                v.FormaAtual = forma;
                v.Molduras = molduras;
                if (tipo == 5) { v.Imagens++; v.VerImagem(forma); }            // imagem: cor nos pixels (v1.3: RGB vai para a lista)
                else if (Array.IndexOf(Efeitos, tipo) >= 0) v.ComEfeito++;    // passos gerados nao sao lidos
                else
                {
                    bool grupo = tipo == 7 || tipo == 16; // grupo e grupo de sombra
                    if (tipo == 6) TextoDe(forma, v);
                    else if (!grupo) Vagas(forma, Parte.Preench, v);
                    if (grupo) Filhos((object)s.Shapes, molduras, v);
                }
                // conteudo do PowerClip (qualquer tipo de moldura)
                List<object> conteudo = Conteudo(forma);
                if (conteudo.Count > 0)
                {
                    v.PowerClips++;
                    List<object> dentro = new List<object>(molduras);
                    dentro.Add(forma);
                    foreach (object f in conteudo)
                    {
                        if (v.Parar) break;
                        Visitar(f, dentro, v);
                    }
                }
            }
            catch (Exception ex) { Falha(ex); }
            finally
            {
                if (travado) { try { s.Locked = true; } catch (Exception ex) { Falha(ex); } }
            }
        }

        // Objetos dentro do PowerClip. Tenta tres caminhos e anota em diagnostico.log o que achou.
        static List<object> Conteudo(object forma)
        {
            List<object> lista = new List<object>();
            object pc = null;
            try { pc = ((dynamic)forma).PowerClip; } catch { return lista; }
            if (pc == null) return lista;
            string caminho = null;
            bool primeiroDeuCerto = false;
            try
            {
                dynamic formas = ((dynamic)pc).Shapes;
                int n = (int)formas.Count;
                for (int i = 1; i <= n; i++) lista.Add((object)formas.Item(i));
                primeiroDeuCerto = true;
                caminho = "Shapes";
            }
            catch (Exception ex) { Falha(ex); lista.Clear(); }
            if (!primeiroDeuCerto)
            {
                try
                {
                    dynamic todas = ((dynamic)pc).Shapes.All();
                    int n = (int)todas.Count;
                    for (int i = 1; i <= n; i++) lista.Add((object)todas.Item(i));
                    caminho = "Shapes.All";
                }
                catch (Exception ex) { Falha(ex); lista.Clear(); }
            }
            if (caminho == null)
            {
                try
                {
                    object formas = pc.GetType().InvokeMember("Shapes", BindingFlags.GetProperty, null, pc, null, CultureInfo.InvariantCulture);
                    int n = Convert.ToInt32(formas.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, formas, null, CultureInfo.InvariantCulture));
                    for (int i = 1; i <= n; i++)
                        lista.Add(formas.GetType().InvokeMember("Item", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, formas,
                            new object[] { i }, CultureInfo.InvariantCulture));
                    caminho = "reflexao";
                }
                catch (Exception ex) { Falha(ex); lista.Clear(); }
            }
            if (lista.Count > 0 && diagPowerClip < 15)
            {
                diagPowerClip++;
                Registro.Diagnostico("PowerClip: " + lista.Count + " objetos dentro (" + caminho + ")");
            }
            else if (caminho == null && diagPowerClip < 15)
            {
                diagPowerClip++;
                Registro.Diagnostico("PowerClip: nao foi possivel ler o conteudo (veja erros.log)");
            }
            return lista;
        }

        static void Filhos(object colecao, List<object> molduras, Visita v)
        {
            try
            {
                dynamic formas = colecao;
                int n = (int)formas.Count;
                List<object> lista = new List<object>(n);
                for (int i = 1; i <= n; i++) lista.Add((object)formas.Item(i));
                foreach (object f in lista)
                {
                    if (v.Parar) break;
                    Visitar(f, molduras, v);
                }
            }
            catch (Exception ex) { Falha(ex); }
        }

        // Texto: letra por letra (palavras com cores diferentes)
        static void TextoDe(object forma, Visita v)
        {
            try
            {
                dynamic letras = ((dynamic)forma).Text.Story.Characters;
                int n = (int)letras.Count;
                if (!v.Trocando) n = Math.Min(n, LimiteLetras);
                for (int i = 1; i <= n && !v.Parar; i++)
                    Vagas((object)letras.Item(i), Parte.Texto, v);
            }
            catch (Exception ex)
            {
                // sem acesso as letras: confere o texto inteiro
                Falha(ex);
                Vagas(forma, Parte.Texto, v);
            }
        }

        // Preenchimento e contorno de um objeto ou de uma letra
        static void Vagas(object alvo, Parte parte, Visita v)
        {
            if (v.UsarPreench)
            {
                int tipo = -1;
                try { tipo = (int)((dynamic)alvo).Fill.Type; } catch (Exception ex) { Falha(ex); }
                if (tipo == 1) // liso
                {
                    object cor = null;
                    try { cor = ((dynamic)alvo).Fill.UniformColor; } catch (Exception ex) { Falha(ex); }
                    InfoCor info = LerCor(cor);
                    if (info != null && v.Ver(info, cor, parte))
                        Contar(AplicarCor(delegate { return (object)((dynamic)alvo).Fill.UniformColor; }, v,
                            delegate { ((dynamic)alvo).Fill.ApplyUniformFill(v.Nova); },
                            delegate { ((dynamic)((dynamic)alvo).Fill.UniformColor).CopyAssign(v.Nova); }), v);
                }
                else if (tipo == 2) // degrade: cor inicial no indice 0 e final em Count + 1
                {
                    Parte parteDegrade = parte == Parte.Preench ? Parte.Degrade : parte;
                    int n = -1;
                    try { n = (int)((dynamic)alvo).Fill.Fountain.Colors.Count; } catch (Exception ex) { Falha(ex); }
                    for (int i = 0; n >= 0 && i <= n + 1; i++)
                    {
                        int indice = i;
                        object cor;
                        try { cor = ((dynamic)alvo).Fill.Fountain.Colors.Item(indice).Color; } catch { break; }
                        InfoCor info = LerCor(cor);
                        if (info != null && v.Ver(info, cor, parteDegrade))
                            Contar(AplicarCor(delegate { return (object)((dynamic)alvo).Fill.Fountain.Colors.Item(indice).Color; }, v,
                                delegate { ((dynamic)((dynamic)alvo).Fill.Fountain.Colors.Item(indice).Color).CopyAssign(v.Nova); }), v);
                    }
                }
            }
            if (v.UsarContorno)
            {
                object cor = null;
                bool tem = false;
                try
                {
                    dynamic o = ((dynamic)alvo).Outline;
                    tem = (int)o.Type != 0 && (double)o.Width > 0; // 0 = sem contorno
                    if (tem) cor = o.Color;
                }
                catch (Exception ex) { Falha(ex); }
                if (tem)
                {
                    InfoCor info = LerCor(cor);
                    if (info != null && v.Ver(info, cor, parte == Parte.Texto ? Parte.Texto : Parte.Contorno))
                        Contar(AplicarCor(delegate { return (object)((dynamic)alvo).Outline.Color; }, v,
                            delegate { ((dynamic)((dynamic)alvo).Outline.Color).CopyAssign(v.Nova); },
                            delegate { ((dynamic)alvo).Outline.SetProperties(Color: v.Nova); }), v);
                }
            }
        }

        static void Contar(bool trocou, Visita v)
        {
            if (!trocou) { v.Recusadas++; return; }
            v.Trocas++;
            if (!ReferenceEquals(v.UltimaTrocada, v.FormaAtual)) { v.UltimaTrocada = v.FormaAtual; v.ObjetosTrocados++; }
        }

        // Troca a cor e confere lendo de novo do objeto. Primeiro pelos valores (CMYKAssign, RGBAssign...),
        // depois pelas alternativas (aplicar a cor nova como objeto).
        static bool AplicarCor(Func<object> pegar, Visita v, params Action[] alternativas)
        {
            object cor = null;
            try { cor = pegar(); } catch (Exception ex) { Falha(ex); }
            if (cor != null && AtribuirValores(cor, v.NovaInfo) && Confere(pegar, v)) return true;
            foreach (Action alternativa in alternativas)
            {
                try { alternativa(); } catch (Exception ex) { Falha(ex); continue; }
                if (Confere(pegar, v)) return true;
            }
            if (falhas < 20)
            {
                falhas++;
                InfoCor depois = null;
                try { depois = LerCor(pegar()); } catch { }
                Registro.Diagnostico("Troca recusada: ficou " + (depois != null ? depois.Chave : "?") + ", esperado " + v.NovaInfo.Chave);
            }
            return false;
        }

        static bool AtribuirValores(object cor, InfoCor nova)
        {
            if (nova == null || nova.Numeros == null) return false;
            dynamic c = cor;
            int[] x = nova.Numeros;
            try
            {
                switch (nova.Modelo)
                {
                    case "CMYK": c.CMYKAssign(x[0], x[1], x[2], x[3]); return true;
                    case "CMY": c.CMYAssign(x[0], x[1], x[2]); return true;
                    case "RGB": c.RGBAssign(x[0], x[1], x[2]); return true;
                    case "Cinza": c.GrayAssign(x[0]); return true;
                }
            }
            catch (Exception ex) { Falha(ex); }
            return false;
        }

        static bool Confere(Func<object> pegar, Visita v)
        {
            object cor = null;
            try { cor = pegar(); } catch { return false; }
            InfoCor i = LerCor(cor);
            return i != null && i.Chave == v.NovaInfo.Chave;
        }

        // ---------------- cores ----------------

        internal static InfoCor LerCor(object cor)
        {
            if (cor == null) return null;
            dynamic c = cor;
            InfoCor i = new InfoCor();
            try { i.Tipo = (int)c.Type; } catch (Exception ex) { Falha(ex); return null; }
            i.Modelo = Modelo(cor, i.Tipo);
            try
            {
                switch (i.Modelo)
                {
                    case "CMYK": i.Numeros = new int[] { (int)c.CMYKCyan, (int)c.CMYKMagenta, (int)c.CMYKYellow, (int)c.CMYKBlack }; break;
                    case "CMY": i.Numeros = new int[] { (int)c.CMYCyan, (int)c.CMYMagenta, (int)c.CMYYellow }; break;
                    case "RGB": i.Numeros = new int[] { (int)c.RGBRed, (int)c.RGBGreen, (int)c.RGBBlue }; break;
                    case "Cinza": i.Numeros = new int[] { (int)c.Gray }; break;
                }
            }
            catch (Exception ex) { Falha(ex); i.Numeros = null; }
            if (i.Numeros != null)
            {
                string[] partes = new string[i.Numeros.Length];
                for (int k = 0; k < partes.Length; k++) partes[k] = i.Numeros[k].ToString(CultureInfo.InvariantCulture);
                i.Descricao = i.Modelo + " " + String.Join("/", partes);
            }
            else i.Descricao = Nomear(cor, i.Modelo);
            return i;
        }

        static string Nomear(object cor, string modelo)
        {
            dynamic c = cor;
            string nome = "";
            try { nome = Convert.ToString(c.Name); } catch { }
            if (nome == null) nome = "";
            if (modelo == "Pantone" || modelo == "Especial")
            {
                try { int tinta = (int)c.Tint; if (tinta >= 0 && tinta < 100) nome += " " + tinta + "%"; } catch { }
                if (nome.Length > 0) return nome;
            }
            string valores = null;
            try
            {
                if (modelo == "HSB") valores = (int)c.HSBHue + "/" + (int)c.HSBSaturation + "/" + (int)c.HSBBrightness;
                else if (modelo == "HLS") valores = (int)c.HLSHue + "/" + (int)c.HLSLightness + "/" + (int)c.HLSSaturation;
                else if (modelo == "Lab")
                    valores = ((double)c.LabLuminance).ToString("0.#", CultureInfo.InvariantCulture) + "/" +
                              ((double)c.LabComponentA).ToString("0.#", CultureInfo.InvariantCulture) + "/" +
                              ((double)c.LabComponentB).ToString("0.#", CultureInfo.InvariantCulture);
            }
            catch { valores = null; }
            if (valores == null) valores = TextoDaCor(cor);
            string r = modelo;
            if (!String.IsNullOrEmpty(valores)) r += " " + valores;
            if (nome.Length > 0) r += " · " + nome;
            return r;
        }

        static CorUsada NovaEntrada(InfoCor info, object cor, object app)
        {
            CorUsada u = new CorUsada { Chave = info.Chave, Modelo = info.Modelo, Descricao = info.Descricao };
            dynamic a = app;
            // copia, para abrir a janela de cores com ela
            try
            {
                dynamic copia = a.CreateCMYKColor(0, 0, 0, 0);
                copia.CopyAssign(cor);
                u.Amostra = (object)copia;
                InfoCor conferida = LerCor((object)copia);
                if (conferida == null || conferida.Chave != info.Chave)
                    Registro.Diagnostico("Copia diferente: " + info.Chave + " -> " + (conferida != null ? conferida.Chave : "?"));
            }
            catch (Exception ex) { Falha(ex); }
            // amostra na tela
            int[] x = info.Numeros;
            if (x != null && info.Modelo == "CMYK")
            {
                u.R = ParaByte(255.0 * (1 - x[0] / 100.0) * (1 - x[3] / 100.0));
                u.G = ParaByte(255.0 * (1 - x[1] / 100.0) * (1 - x[3] / 100.0));
                u.B = ParaByte(255.0 * (1 - x[2] / 100.0) * (1 - x[3] / 100.0));
            }
            else if (x != null && info.Modelo == "RGB") { u.R = ParaByte(x[0]); u.G = ParaByte(x[1]); u.B = ParaByte(x[2]); }
            else if (x != null && info.Modelo == "CMY")
            {
                u.R = ParaByte(255.0 * (1 - x[0] / 100.0)); u.G = ParaByte(255.0 * (1 - x[1] / 100.0)); u.B = ParaByte(255.0 * (1 - x[2] / 100.0));
            }
            else if (x != null && info.Modelo == "Cinza") { u.R = u.G = u.B = ParaByte(x[0]); }
            else
            {
                try
                {
                    dynamic tela = a.CreateCMYKColor(0, 0, 0, 0);
                    tela.CopyAssign(cor);
                    tela.ConvertToRGB();
                    u.R = ParaByte((int)tela.RGBRed); u.G = ParaByte((int)tela.RGBGreen); u.B = ParaByte((int)tela.RGBBlue);
                }
                catch (Exception ex) { Falha(ex); }
            }
            return u;
        }

        static byte ParaByte(double v) { return (byte)Math.Max(0, Math.Min(255, Math.Round(v))); }

        // Modelo de cor pelo tipo do Corel; na primeira vez de cada tipo confere tambem pelo texto da cor
        // (quando o Corel fornece) e anota em diagnostico.log.
        static string Modelo(object cor, int tipo)
        {
            string m;
            if (modelos.TryGetValue(tipo, out m)) return m;
            m = PorNumero(tipo);
            string texto = TextoDaCor(cor);
            string pelo = PorTexto(texto);
            if (pelo != null) m = pelo;
            modelos[tipo] = m;
            Registro.Diagnostico("Tipo de cor " + tipo + " = " + m + (texto != null ? " · " + texto : ""));
            return m;
        }

        static string PorNumero(int tipo)
        {
            switch (tipo)
            {
                case 1: return "Pantone";
                case 2: return "CMYK";
                case 4: return "CMY";
                case 5: return "RGB";
                case 6: return "HSB";
                case 7: return "HLS";
                case 8: return "Preto e branco";
                case 9: return "Cinza";
                case 11: return "YIQ";
                case 12: return "Lab";
                case 20: return "Registro";
                case 25: return "Especial";
                default: return "Cor";
            }
        }

        static string PorTexto(string texto)
        {
            if (String.IsNullOrEmpty(texto)) return null;
            string t = texto.TrimStart().ToUpperInvariant();
            string[] prefixos = { "CMYK", "CMY", "RGB", "HSB", "HLS", "LAB", "YIQ", "GRAY", "GREY", "BW", "REGISTRATION", "PANTONE", "SPOT", "USERINK" };
            string[] nomes = { "CMYK", "CMY", "RGB", "HSB", "HLS", "Lab", "YIQ", "Cinza", "Cinza", "Preto e branco", "Registro", "Pantone", "Especial", "Especial" };
            for (int i = 0; i < prefixos.Length; i++)
                if (t.StartsWith(prefixos[i], StringComparison.Ordinal)) return nomes[i];
            return null;
        }

        static string TextoDaCor(object cor)
        {
            if (textoCor == 0 || cor == null) return null;
            try
            {
                string s = cor.GetType().InvokeMember("ToString", BindingFlags.InvokeMethod, null, cor, null, CultureInfo.InvariantCulture) as string;
                if (String.IsNullOrEmpty(s) || s.StartsWith("System.", StringComparison.Ordinal)) { textoCor = 0; return null; }
                textoCor = 1; return s;
            }
            catch { textoCor = 0; return null; }
        }

        static bool Logico(object alvo, string nome, bool padrao)
        {
            try
            {
                return Convert.ToBoolean(alvo.GetType().InvokeMember(nome, BindingFlags.GetProperty, null, alvo, null, CultureInfo.InvariantCulture));
            }
            catch { return padrao; }
        }

        static void Falha(Exception ex) { if (falhas++ < 20) Registro.Erro(ex); }
    }

    // ======================================================================
    // Painel
    // ======================================================================
    public sealed class DockerUI : UserControl
    {
        readonly Corel corel;
        readonly DispatcherTimer timer;
        readonly ComboBox alcance = new ComboBox();
        readonly CheckBox usarPreench = new CheckBox(), usarContorno = new CheckBox();
        readonly TextBlock resumo = new TextBlock();
        readonly TextBlock status = new TextBlock();
        readonly StackPanel lista = new StackPanel();
        readonly Button converter = new Button(); // v1.4: converter RGB para CMYK
        Leitura leitura;
        bool ativo;
        int geracao;
        string chaveVista;
        string selecaoChave; // v1.1: cliques seguidos em Selecionar passam pelos grupos
        int selecaoProximo;
        // v1.3: imagens fora do CMYK, com Anterior/Próxima
        readonly StackPanel secaoImagens = new StackPanel();
        readonly TextBlock tituloImagens = new TextBlock();
        readonly TextBlock posicaoImagens = new TextBlock();
        readonly StackPanel listaImagens = new StackPanel();
        readonly TextBlock tituloCores = new TextBlock();
        readonly List<Button> linhasImagens = new List<Button>();
        Button maisImagens;
        int posicaoImagem = -1;
        const int BlocoImagens = 100;
        static readonly Brush Fundo = Pincel(67, 67, 67);
        static readonly Brush Texto = Pincel(235, 235, 235);
        static readonly Brush Suave = Pincel(190, 190, 190);
        static readonly Brush Alerta = Pincel(255, 184, 77);
        static readonly Brush Linha = Pincel(79, 79, 79);
        static readonly Brush LinhaAtual = Pincel(52, 96, 130); // v1.3: item selecionado na lista de imagens

        public DockerUI(object app)
        {
            corel = new Corel(app);
            FontFamily = new FontFamily("Segoe UI");
            FontSize = 12;
            Background = Fundo;
            Foreground = Texto;
            MinWidth = 240;
            MinHeight = 300;
            ConstruirInterface();
            timer = new DispatcherTimer(DispatcherPriority.Background, Dispatcher);
            timer.Interval = TimeSpan.FromMilliseconds(1500); // v1.2: confere com menos frequencia
            timer.Tick += AoTick;
            Loaded += delegate { Ativar(); };
            Unloaded += delegate { Desativar(); };
            IsVisibleChanged += delegate { if (IsVisible && IsLoaded) Ativar(); else Desativar(); };
            PreviewKeyDown += delegate(object s, KeyEventArgs e)
            {
                if (e.Key == Key.F5) { LerCores(); e.Handled = true; }
            };
        }

        static Brush Pincel(byte r, byte g, byte b)
        {
            SolidColorBrush p = new SolidColorBrush(Color.FromRgb(r, g, b));
            p.Freeze(); return p;
        }

        static TextBlock Rotulo(string texto, Brush cor)
        {
            return new TextBlock { Text = texto, Foreground = cor, TextWrapping = TextWrapping.Wrap, VerticalAlignment = VerticalAlignment.Center };
        }

        // v1.3: botao cinza (linha da lista), igual nos temas claro e escuro do Windows
        static ControlTemplate ModeloLinha()
        {
            FrameworkElementFactory borda = new FrameworkElementFactory(typeof(Border), "borda");
            borda.SetValue(Border.BackgroundProperty, new TemplateBindingExtension(Button.BackgroundProperty));
            borda.SetValue(Border.BorderBrushProperty, new TemplateBindingExtension(Button.BorderBrushProperty));
            borda.SetValue(Border.BorderThicknessProperty, new TemplateBindingExtension(Button.BorderThicknessProperty));
            borda.SetValue(Border.PaddingProperty, new TemplateBindingExtension(Button.PaddingProperty));
            FrameworkElementFactory apresentador = new FrameworkElementFactory(typeof(ContentPresenter));
            apresentador.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            borda.AppendChild(apresentador);
            ControlTemplate modelo = new ControlTemplate(typeof(Button)); modelo.VisualTree = borda;
            Trigger hover = new Trigger { Property = Button.IsMouseOverProperty, Value = true };
            hover.Setters.Add(new Setter(Border.BackgroundProperty, Pincel(97, 97, 97), "borda")); modelo.Triggers.Add(hover);
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
            raiz.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            raiz.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            Content = raiz;

            StackPanel topo = new StackPanel();
            ScrollViewer rolavel = new ScrollViewer { Content = topo,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto, HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled };
            raiz.Children.Add(rolavel);

            TextBlock titulo = Rotulo("Trocar cores · v1.4", Texto);
            titulo.FontWeight = FontWeights.SemiBold;
            titulo.Margin = new Thickness(10, 9, 8, 5);
            topo.Children.Add(titulo);
            TextBlock dica = Rotulo("Cores usadas na arte, inclusive dentro de grupos e PowerClips. Selecionar: marca no Corel os objetos com a cor (página atual). Trocar: escolhe a nova cor e troca em todos os lugares. Imagens RGB aparecem numa lista própria: clique ou use Anterior/Próxima para ir até elas. Converter RGB para CMYK: converte tudo de uma vez (página atual ou todas).", Suave);
            dica.FontSize = 11; dica.Margin = new Thickness(10, 0, 8, 8);
            topo.Children.Add(dica);

            Grid linha = new Grid { Margin = new Thickness(10, 0, 10, 2) };
            linha.ColumnDefinitions.Add(new ColumnDefinition());
            linha.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(85) });
            alcance.Items.Add("Página atual");
            alcance.Items.Add("Todas as páginas");
            alcance.SelectedIndex = 0; alcance.Margin = new Thickness(0, 0, 7, 0);
            alcance.ToolTip = "Onde ler e trocar as cores. Selecionar funciona sempre na página atual.";
            alcance.SelectionChanged += delegate { LerCores(); };
            linha.Children.Add(alcance);
            Button ler = new Button { Content = "Ler cores", Height = 27, ToolTip = "Lê de novo as cores da arte (F5)" };
            ler.Click += delegate { LerCores(); };
            Grid.SetColumn(ler, 1); linha.Children.Add(ler);
            topo.Children.Add(linha);

            StackPanel opcoes = new StackPanel { Orientation = Orientation.Horizontal };
            ConfigurarCaixa(usarPreench, "Preenchimento"); usarPreench.IsChecked = true;
            ConfigurarCaixa(usarContorno, "Contorno"); usarContorno.IsChecked = true;
            usarPreench.Click += delegate { LerCores(); };
            usarContorno.Click += delegate { LerCores(); };
            opcoes.Children.Add(usarPreench); opcoes.Children.Add(usarContorno);
            topo.Children.Add(opcoes);

            // v1.4: converte tudo o que está fora do CMYK (cores e imagens), na página atual ou em todas
            converter.Content = "Converter RGB para CMYK"; converter.Height = 27;
            converter.Margin = new Thickness(10, 4, 10, 6);
            converter.ToolTip = "Converte para CMYK as cores RGB (preenchimento, contorno, degradê e texto) e as imagens RGB, " +
                "na página atual ou em todas as páginas, conforme a opção acima. Pantone, cores especiais, registro e cinza não mudam.";
            converter.Visibility = Visibility.Collapsed;
            converter.Click += delegate { ConverterRgb(); };
            topo.Children.Add(converter);

            resumo.Text = "—"; resumo.Foreground = Suave; resumo.TextWrapping = TextWrapping.Wrap;
            resumo.Margin = new Thickness(10, 1, 8, 6);
            topo.Children.Add(resumo);

            // v1.3: imagens fora do CMYK (RGB, Lab, paleta): lista e Anterior/Próxima com zoom
            secaoImagens.Margin = new Thickness(10, 0, 10, 10);
            secaoImagens.Visibility = Visibility.Collapsed;
            tituloImagens.Foreground = Alerta; tituloImagens.FontWeight = FontWeights.SemiBold; tituloImagens.TextWrapping = TextWrapping.Wrap;
            tituloImagens.Margin = new Thickness(0, 0, 0, 4);
            tituloImagens.ToolTip = "Imagens em RGB, Lab ou com paleta: no Corel, converta em Bitmaps > Modo > Cores CMYK.";
            secaoImagens.Children.Add(tituloImagens);
            Grid navegar = new Grid { Margin = new Thickness(0, 0, 0, 4) };
            navegar.ColumnDefinitions.Add(new ColumnDefinition());
            navegar.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            navegar.ColumnDefinitions.Add(new ColumnDefinition());
            Button anterior = new Button { Content = "‹ Anterior", Height = 25, FontSize = 11,
                ToolTip = "Seleciona a imagem anterior e dá zoom nela no Corel" };
            anterior.Click += delegate { IrParaImagem(posicaoImagem - 1); };
            navegar.Children.Add(anterior);
            posicaoImagens.Foreground = Texto; posicaoImagens.FontSize = 11; posicaoImagens.MinWidth = 56;
            posicaoImagens.TextAlignment = TextAlignment.Center; posicaoImagens.VerticalAlignment = VerticalAlignment.Center;
            posicaoImagens.Margin = new Thickness(8, 0, 8, 0);
            Grid.SetColumn(posicaoImagens, 1); navegar.Children.Add(posicaoImagens);
            Button proxima = new Button { Content = "Próxima ›", Height = 25, FontSize = 11,
                ToolTip = "Seleciona a próxima imagem e dá zoom nela no Corel" };
            proxima.Click += delegate { IrParaImagem(posicaoImagem + 1); };
            Grid.SetColumn(proxima, 2); navegar.Children.Add(proxima);
            secaoImagens.Children.Add(navegar);
            secaoImagens.Children.Add(listaImagens);
            topo.Children.Add(secaoImagens);

            tituloCores.Text = "Cores"; tituloCores.Foreground = Texto; tituloCores.FontWeight = FontWeights.SemiBold;
            tituloCores.Margin = new Thickness(10, 0, 8, 4); tituloCores.Visibility = Visibility.Collapsed;
            topo.Children.Add(tituloCores);
            lista.Margin = new Thickness(10, 0, 10, 8);
            topo.Children.Add(lista);

            StackPanel rodape = new StackPanel { Margin = new Thickness(10, 0, 10, 9) };
            status.TextWrapping = TextWrapping.Wrap; status.Foreground = Suave; status.FontSize = 11;
            status.Margin = new Thickness(0, 6, 0, 4);
            rodape.Children.Add(status);
            // assinatura
            TextBlock assinatura = Rotulo("by Bruno Faqui", Suave);
            assinatura.FontSize = 11; assinatura.FontStyle = FontStyles.Italic; assinatura.Margin = new Thickness(0, 6, 0, 0);
            rodape.Children.Add(assinatura);
            Grid.SetRow(rodape, 1); raiz.Children.Add(rodape);
        }

        void Ativar()
        {
            if (ativo || !IsVisible) return;
            ativo = true;
            timer.Start();
            // v1.2: nao le sozinho; so quando pedir (Ler cores/F5)
            if (leitura == null)
            {
                resumo.Text = "Clique em Ler cores (F5) para listar as cores da arte.";
                resumo.Foreground = Suave;
            }
        }

        void Desativar()
        {
            if (!ativo) return;
            ativo = false;
            timer.Stop();
        }

        // v1.2: troca de documento (ou de pagina, na leitura da pagina atual): a lista fica velha, mas nao le
        // sozinho (ler a arte inteira pesa); avisa para clicar em Ler cores
        void AoTick(object sender, EventArgs e)
        {
            if (!ativo) return;
            try
            {
                string chave = ChaveDaVista();
                if (chave == chaveVista) return;
                chaveVista = chave;
                if (leitura != null)
                {
                    leitura = null;
                    MostrarCores(); // v1.3: limpa tambem a lista de imagens
                    resumo.Text = "A página ou o arquivo mudou. Clique em Ler cores (F5).";
                    resumo.Foreground = Suave;
                }
            }
            catch (Exception ex) { Registro.Erro(ex); }
        }

        string ChaveDaVista()
        {
            string chave = corel.ChaveAtual();
            if (chave == null) return "";
            if (alcance.SelectedIndex == 1)
            {
                int dois = chave.IndexOf(':');
                if (dois > 0) chave = chave.Substring(0, dois); // todas as paginas: so o documento importa
            }
            return chave;
        }

        void LerCores()
        {
            if (!ativo) return;
            int minha = ++geracao;
            selecaoChave = null; selecaoProximo = 0;
            resumo.Text = "Lendo as cores…"; resumo.Foreground = Suave;
            Dispatcher.BeginInvoke(new Action(delegate
            {
                if (!ativo || minha != geracao) return;
                try
                {
                    chaveVista = ChaveDaVista();
                    object doc = corel.Documento();
                    if (doc == null) { leitura = null; MostrarCores(); resumo.Text = "Abra uma arte no CorelDRAW."; return; }
                    bool todas = alcance.SelectedIndex == 1;
                    Leitura l = Motor.Ler(corel.Aplicacao, doc, todas, usarPreench.IsChecked == true, usarContorno.IsChecked == true);
                    l.ChaveCorel = corel.ChaveAtual();
                    leitura = l;
                    MostrarCores();
                }
                catch (Exception ex)
                {
                    Registro.Erro(ex);
                    leitura = null; MostrarCores();
                    resumo.Text = "Não foi possível ler as cores: " + Registro.Explicar(ex);
                }
            }), DispatcherPriority.Background);
        }

        void MostrarCores()
        {
            lista.Children.Clear();
            Leitura l = leitura;
            MostrarImagens(l); // v1.3
            resumo.Foreground = Suave;
            converter.Visibility = Visibility.Collapsed;
            if (l == null) { resumo.Text = "—"; return; }
            int foraTotal = l.ImagensForaTotal;
            foreach (CorUsada c in l.Cores) if (c.ForaDoCmyk) foraTotal++;
            if (foraTotal > 0)
            {
                converter.Content = l.TodasPaginas ? "Converter RGB para CMYK (todas as páginas)" : "Converter RGB para CMYK (esta página)";
                converter.Visibility = Visibility.Visible;
            }
            List<string> notas = new List<string>();
            if (l.Imagens > 0)
                notas.Add(l.Imagens + (l.Imagens == 1 ? " imagem" : " imagens") + (l.ImagensForaTotal > 0
                    ? " (" + l.ImagensForaTotal + " fora do CMYK, na lista abaixo)" : " (cores nos pixels, não entram)"));
            if (l.PowerClips > 0) notas.Add("inclui " + l.PowerClips + (l.PowerClips == 1 ? " PowerClip" : " PowerClips"));
            if (l.ComEfeito > 0) notas.Add(l.ComEfeito + " mistura/contour/extrusão não lidos");
            if (l.ForaDaPagina > 0) notas.Add(l.ForaDaPagina + " fora da página (ignorados)");
            if (l.Parcial) notas.Add("leitura parcial: arte muito grande");
            string extra = notas.Count > 0 ? "\n" + String.Join(" · ", notas.ToArray()) : "";
            if (l.Cores.Count == 0)
            {
                resumo.Text = "Nenhuma cor encontrada" + (l.TodasPaginas ? " no arquivo." : " nesta página.") + extra;
                return;
            }
            List<CorUsada> ordem = new List<CorUsada>(l.Cores);
            ordem.Sort(delegate(CorUsada x, CorUsada y) { return y.Total.CompareTo(x.Total); });
            int fora = 0;
            foreach (CorUsada u in ordem) if (u.ForaDoCmyk) fora++;
            resumo.Text = ordem.Count + (ordem.Count == 1 ? " cor" : " cores") + " em " + l.Objetos + " objetos" +
                (fora > 0 ? " · " + fora + " fora do CMYK" : "") + extra;
            int mostrar = Math.Min(ordem.Count, 80);
            for (int i = 0; i < mostrar; i++) lista.Children.Add(LinhaCor(ordem[i]));
            if (ordem.Count > mostrar)
            {
                TextBlock mais = Rotulo("… e mais " + (ordem.Count - mostrar) + " cores pouco usadas.", Suave);
                mais.FontSize = 11; mais.Margin = new Thickness(0, 3, 0, 0);
                lista.Children.Add(mais);
            }
        }

        // ---------------- v1.3: imagens fora do CMYK ----------------

        void MostrarImagens(Leitura l)
        {
            listaImagens.Children.Clear(); linhasImagens.Clear(); maisImagens = null; posicaoImagem = -1;
            if (l == null || l.ImagensFora.Count == 0)
            {
                secaoImagens.Visibility = Visibility.Collapsed;
                tituloCores.Visibility = Visibility.Collapsed;
                return;
            }
            int n = l.ImagensFora.Count;
            tituloImagens.Text = "Imagens fora do CMYK: " + l.ImagensForaTotal +
                (l.ImagensForaTotal > n ? " (a lista mostra as primeiras " + n + ")" : "");
            secaoImagens.Visibility = Visibility.Visible;
            tituloCores.Visibility = l.Cores.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
            AtualizarPosicaoImagem();
            MaisImagens();
        }

        // lista em blocos de 100 (com muitas imagens o painel nao trava); "Mostrar mais" acrescenta o proximo bloco
        void MaisImagens()
        {
            Leitura l = leitura;
            if (l == null) return;
            if (maisImagens != null) { listaImagens.Children.Remove(maisImagens); maisImagens = null; }
            int de = linhasImagens.Count, ate = Math.Min(l.ImagensFora.Count, de + BlocoImagens);
            for (int i = de; i < ate; i++)
            {
                int indice = i;
                ImagemAchada img = l.ImagensFora[i];
                Grid conteudo = new Grid();
                conteudo.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                conteudo.ColumnDefinitions.Add(new ColumnDefinition());
                TextBlock numero = Rotulo((i + 1) + ".", Suave);
                numero.FontSize = 11; numero.Margin = new Thickness(0, 0, 6, 0); numero.VerticalAlignment = VerticalAlignment.Top;
                conteudo.Children.Add(numero);
                string onde = !l.TodasPaginas ? "" : img.Pagina == null ? " · página mestre" : " · pág. " + img.NumeroPagina;
                TextBlock texto = Rotulo(img.Descricao + onde, Texto);
                texto.FontSize = 11;
                Grid.SetColumn(texto, 1); conteudo.Children.Add(texto);
                Button linha = new Button { Content = conteudo, Foreground = Texto,
                    Background = indice == posicaoImagem ? LinhaAtual : Linha,
                    BorderBrush = Pincel(94, 94, 94), BorderThickness = new Thickness(0, 0, 0, 1),
                    HorizontalContentAlignment = HorizontalAlignment.Stretch, Padding = new Thickness(6, 3, 6, 3),
                    Cursor = Cursors.Hand, ToolTip = "Clique para selecionar e dar zoom nesta imagem no Corel" };
                linha.Template = ModeloLinha();
                linha.Click += delegate { IrParaImagem(indice); };
                linhasImagens.Add(linha);
                listaImagens.Children.Add(linha);
            }
            int resto = l.ImagensFora.Count - linhasImagens.Count;
            if (resto > 0)
            {
                TextBlock texto = Rotulo("Mostrar mais " + Math.Min(resto, BlocoImagens) + "  (faltam " + resto + ")", Texto);
                texto.FontSize = 11; texto.TextAlignment = TextAlignment.Center;
                maisImagens = new Button { Content = texto, Foreground = Texto, Background = Pincel(64, 64, 64),
                    BorderBrush = Pincel(94, 94, 94), BorderThickness = new Thickness(1),
                    Padding = new Thickness(6, 4, 6, 4), Margin = new Thickness(0, 4, 0, 0), Cursor = Cursors.Hand,
                    ToolTip = "Também dá para passar imagem por imagem com Anterior / Próxima" };
                maisImagens.Template = ModeloLinha();
                maisImagens.Click += delegate { MaisImagens(); };
                listaImagens.Children.Add(maisImagens);
            }
        }

        void AtualizarPosicaoImagem()
        {
            Leitura l = leitura;
            int n = l == null ? 0 : l.ImagensFora.Count;
            posicaoImagens.Text = posicaoImagem < 0 ? (n == 1 ? "1 imagem" : n + " imagens") : (posicaoImagem + 1) + " de " + n;
        }

        // seleciona a imagem da lista e da zoom; da ultima volta a primeira
        void IrParaImagem(int indice)
        {
            Leitura l = leitura;
            if (l == null || l.ImagensFora.Count == 0) return;
            int n = l.ImagensFora.Count;
            if (indice < 0) indice = n - 1;
            if (indice >= n) indice = 0;
            int anterior = posicaoImagem;
            posicaoImagem = indice;
            if (anterior >= 0 && anterior < linhasImagens.Count) linhasImagens[anterior].Background = Linha;
            if (indice < linhasImagens.Count) linhasImagens[indice].Background = LinhaAtual;
            AtualizarPosicaoImagem();
            try
            {
                object doc = corel.Documento();
                if (doc == null) return;
                if (!l.TodasPaginas && l.ChaveCorel != corel.ChaveAtual())
                {
                    status.Text = "A página mudou. Clique em Ler cores (F5).";
                    return;
                }
                ImagemAchada img = l.ImagensFora[indice];
                selecaoChave = null; selecaoProximo = 0;
                status.Text = (indice + 1) + " de " + n + ": " + img.Descricao + "\n" + Motor.SelecionarImagem(corel.Aplicacao, doc, img);
            }
            catch (Exception ex) { Registro.Erro(ex); status.Text = "Não foi possível selecionar: " + Registro.Explicar(ex); }
        }

        UIElement LinhaCor(CorUsada u)
        {
            Grid g = new Grid { Margin = new Thickness(0, 0, 0, 1), Background = Linha };
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(32) });
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            Border amostra = new Border { Width = 22, Height = 22, Margin = new Thickness(5, 6, 5, 6),
                Background = Pincel(u.R, u.G, u.B), BorderBrush = Pincel(35, 35, 35), BorderThickness = new Thickness(1),
                VerticalAlignment = VerticalAlignment.Top };
            g.Children.Add(amostra);
            StackPanel textos = new StackPanel { Margin = new Thickness(0, 4, 4, 5) };
            textos.Children.Add(Rotulo(u.Descricao + (u.ForaDoCmyk ? "  ·  fora do CMYK" : ""), u.ForaDoCmyk ? Alerta : Texto));
            TextBlock usos = Rotulo(Usos(u), Suave); usos.FontSize = 11;
            textos.Children.Add(usos);
            StackPanel botoes = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 4, 0, 0) };
            Button sel = new Button { Content = "Selecionar", Height = 23, Padding = new Thickness(9, 0, 9, 0),
                Margin = new Thickness(0, 0, 6, 0), FontSize = 11,
                ToolTip = "Seleciona no Corel os objetos desta página que usam esta cor. Dentro de PowerClip, entra nele. Clique de novo para o próximo lugar." };
            sel.Click += delegate { Selecionar(u); };
            Button troca = new Button { Content = "Trocar…", Height = 23, Padding = new Thickness(9, 0, 9, 0), FontSize = 11,
                ToolTip = "Escolhe a nova cor e troca esta em todos os lugares (um Ctrl+Z desfaz)" };
            troca.Click += delegate { Trocar(u); };
            botoes.Children.Add(sel); botoes.Children.Add(troca);
            textos.Children.Add(botoes);
            Grid.SetColumn(textos, 1); g.Children.Add(textos);
            return g;
        }

        static string Usos(CorUsada u)
        {
            List<string> p = new List<string>();
            if (u.Preench > 0) p.Add("preench. " + u.Preench);
            if (u.Degrade > 0) p.Add("degradê " + u.Degrade);
            if (u.Contorno > 0) p.Add("contorno " + u.Contorno);
            if (u.Texto > 0) p.Add("texto " + u.Texto);
            return String.Join(" · ", p.ToArray());
        }

        void Selecionar(CorUsada u)
        {
            try
            {
                object doc = corel.Documento();
                if (doc == null) return;
                int pedido = selecaoChave == u.Chave ? selecaoProximo : 0;
                int grupos;
                status.Text = Motor.Selecionar(corel.Aplicacao, doc, u.Chave, usarPreench.IsChecked == true,
                    usarContorno.IsChecked == true, pedido, out grupos);
                selecaoChave = u.Chave;
                selecaoProximo = pedido + 1;
            }
            catch (Exception ex) { Registro.Erro(ex); status.Text = "Não foi possível selecionar: " + Registro.Explicar(ex); }
        }

        // v1.4: converte para CMYK as cores e imagens fora do CMYK, com confirmação
        void ConverterRgb()
        {
            try
            {
                object doc = corel.Documento();
                if (doc == null) return;
                Leitura l = leitura;
                if (l != null && !l.TodasPaginas && l.ChaveCorel != corel.ChaveAtual())
                {
                    status.Text = "A página mudou. Clique em Ler cores antes de converter.";
                    return;
                }
                bool todas = alcance.SelectedIndex == 1;
                int cores = 0, imagens = 0;
                if (l != null)
                {
                    foreach (CorUsada u in l.Cores) if (u.ForaDoCmyk) cores++;
                    imagens = l.ImagensForaTotal;
                }
                string pergunta = "Converter para CMYK " + (todas ? "em TODAS as páginas" : "nesta página") + ":\n\n" +
                    "• " + cores + (cores == 1 ? " cor fora do CMYK (RGB, Lab...)" : " cores fora do CMYK (RGB, Lab...)") + "\n" +
                    "• " + imagens + (imagens == 1 ? " imagem RGB/Lab/paleta" : " imagens RGB/Lab/paleta") + "\n\n" +
                    "Preenchimento, contorno, degradê e texto. Pantone, cores especiais, registro e cinza não mudam.\n" +
                    "A conversão usa o gerenciamento de cores do Corel. Ctrl+Z desfaz tudo de uma vez.";
                if (MessageBox.Show(pergunta, "Converter RGB para CMYK", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                    return;
                status.Text = "Convertendo…";
                Mouse.OverrideCursor = Cursors.Wait;
                try { status.Text = Motor.ConverterParaCmyk(corel.Aplicacao, doc, todas); }
                finally { Mouse.OverrideCursor = null; }
                LerCores();
            }
            catch (Exception ex) { Registro.Erro(ex); status.Text = "Não foi possível converter: " + Registro.Explicar(ex); }
        }

        void Trocar(CorUsada u)
        {
            try
            {
                object doc = corel.Documento();
                if (doc == null) return;
                Leitura l = leitura;
                if (l != null && !l.TodasPaginas && l.ChaveCorel != corel.ChaveAtual())
                {
                    status.Text = "A página mudou. Clique em Ler cores antes de trocar.";
                    return;
                }
                dynamic app = corel.Aplicacao;
                dynamic nova = app.CreateCMYKColor(0, 0, 0, 0);
                if (u.Amostra != null) { try { nova.CopyAssign(u.Amostra); } catch (Exception ex) { Registro.Erro(ex); } }
                if (!(bool)nova.UserAssignEx()) return; // janela de cores do Corel; cancelou
                bool todas = alcance.SelectedIndex == 1;
                status.Text = Motor.Trocar(corel.Aplicacao, doc, u.Chave, (object)nova, todas,
                    usarPreench.IsChecked == true, usarContorno.IsChecked == true);
                LerCores();
            }
            catch (Exception ex) { Registro.Erro(ex); status.Text = "Não foi possível trocar: " + Registro.Explicar(ex); }
        }
    }
}
