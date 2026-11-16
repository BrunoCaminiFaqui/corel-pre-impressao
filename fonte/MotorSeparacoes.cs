// Preserva as chapas CMYK e especiais em PDF e processa com Ghostscript.
// O Ghostscript e uma instalacao externa; nao e redistribuido neste pacote.
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
#if !MOTOR_APENAS
using Microsoft.Win32;
using System.Windows.Media;
using System.Windows.Media.Imaging;
#endif

namespace SeparacoesCorel
{
    internal static class ExportacaoPdf
    {
        internal static void Exportar(object documento, string arquivo)
        {
            dynamic doc = documento;
            object opcoes = doc.PDFSettings;
            // Todos os valores originais sao lidos antes da primeira mudanca.
            // Sao restaurados inclusive se PublishToPDF falhar.
            string[] nomes = { "pdfVersion", "PublishRange", "ColorMode", "OutputSpotColorsAs", "Overprints",
                "ComplexFillsAsBitmaps", "TextAsCurves", "IncludeBleed", "CropMarks",
                "RegistrationMarks", "DensitometerScales", "FileInformation", "DownsampleColor", "DownsampleGray", "DownsampleMono", "EncryptType" };
            object[] valores = { 8, 1, 1, 0, true, false, true, false, false, false, false, false, false, false, false, 0 };
            object[] anteriores = new object[nomes.Length];
            for (int i = 0; i < nomes.Length; i++) anteriores[i] = Ler(opcoes, nomes[i]);
            Exception erro = null;
            List<Exception> erros = new List<Exception>();
            int alterados = 0;
            try
            {
                for (int i = 0; i < nomes.Length; i++)
                {
                    alterados = i + 1;
                    Gravar(opcoes, nomes[i], valores[i]);
                }
                doc.PublishToPDF(arquivo);
            }
            catch (Exception ex) { erro = ex; }
            finally
            {
                for (int i = 0; i < alterados; i++)
                    try { Gravar(opcoes, nomes[i], anteriores[i]); }
                    catch (Exception ex) { erros.Add(new Exception("Não foi possível restaurar a opção PDF " + nomes[i], ex)); }
            }
            if (erros.Count > 0)
            {
                if (erro != null) erros.Insert(0, erro);
                throw new AggregateException("Revise as configurações de publicação em PDF: uma opção não pôde ser restaurada.", erros);
            }
            if (erro != null) throw new InvalidOperationException("Falha ao gerar o PDF de separações: " + erro.Message, erro);
        }
        static object Ler(object alvo, string nome)
        { return alvo.GetType().InvokeMember(nome, BindingFlags.GetProperty, null, alvo, null, CultureInfo.InvariantCulture); }
        static void Gravar(object alvo, string nome, object valor)
        { alvo.GetType().InvokeMember(nome, BindingFlags.SetProperty, null, alvo, new object[] { valor }, CultureInfo.InvariantCulture); }
    }

    internal sealed class CorDaChapa
    {
        internal string Nome;
        internal byte R, G, B;
    }

    internal static class MotorSeparacoes
    {
        internal static Dictionary<string, CorDaChapa> LerCores(string log)
        {
            Dictionary<string, CorDaChapa> cores = new Dictionary<string, CorDaChapa>(StringComparer.Ordinal);
            string padrao = "%%SeparationColor:\\s*\"(.*?)\"\\s+100%\\s+ink\\s*=\\s*(\\d+)\\s+(\\d+)\\s+(\\d+)\\s+(\\d+)\\s+CMYK";
            foreach (Match m in Regex.Matches(log, padrao))
            {
                double c = Int32.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture) / 32760.0;
                double mag = Int32.Parse(m.Groups[3].Value, CultureInfo.InvariantCulture) / 32760.0;
                double y = Int32.Parse(m.Groups[4].Value, CultureInfo.InvariantCulture) / 32760.0;
                double k = Int32.Parse(m.Groups[5].Value, CultureInfo.InvariantCulture) / 32760.0;
                cores[m.Groups[1].Value] = new CorDaChapa { Nome = m.Groups[1].Value,
                    R = ByteCor((1 - c) * (1 - k)), G = ByteCor((1 - mag) * (1 - k)), B = ByteCor((1 - y) * (1 - k)) };
            }
            return cores;
        }
        static byte ByteCor(double v) { return (byte)Math.Round(Math.Max(0, Math.Min(1, v)) * 255); }

        internal static string Executar(string executavel, string pdf, string diretorio, int dpi)
        { return Executar(executavel, pdf, diretorio, dpi, 1); }

        // v2.2: fator > 1 renderiza em resolucao maior e reduz no proprio Ghostscript
        // (DownScaleFactor). Suaviza bordas, textos e fotos (anti-serrilhado).
        internal static string Executar(string executavel, string pdf, string diretorio, int dpi, int fator)
        {
            if (fator < 1) fator = 1;
            if (fator > 8) fator = 8;
            ProcessStartInfo info = new ProcessStartInfo { FileName = executavel, UseShellExecute = false,
                CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true,
                WorkingDirectory = diretorio, StandardOutputEncoding = Encoding.UTF8, StandardErrorEncoding = Encoding.UTF8 };
            info.Arguments = "-dSAFER -dBATCH -dNOPAUSE -sDEVICE=tiffsep -dPrintSpotCMYK=true " +
                "-dMaxSpots=60 -dFirstPage=1 -dLastPage=1 -dUseCropBox -sCompression=none -dOverprint=/enable " +
                // v2.2: ignora perfis ICC embutidos no PDF. O CMYK passa direto, com os mesmos
                // numeros do Corel (preto 70% fica somente no K), e cinza vai somente para o K.
                "-dOverrideICC=true -dDeviceGrayToK=true " +
                "-r" + (dpi * fator).ToString(CultureInfo.InvariantCulture) +
                (fator > 1 ? " -dDownScaleFactor=" + fator.ToString(CultureInfo.InvariantCulture) : "") +
                " -sOutputFile=" + Aspas(Path.Combine(diretorio, "placas.tif")) + " " + Aspas(pdf);
            using (Process processo = Process.Start(info))
            {
                Task<string> saida = Task.Factory.StartNew(delegate { return processo.StandardOutput.ReadToEnd(); });
                Task<string> erro = Task.Factory.StartNew(delegate { return processo.StandardError.ReadToEnd(); });
                if (!processo.WaitForExit(120000))
                {
                    try { processo.Kill(); processo.WaitForExit(); } catch { }
                    throw new TimeoutException("A geração das separações excedeu 2 minutos. Reduza a resolução e tente novamente.");
                }
                string log = saida.Result + "\n" + erro.Result;
                if (processo.ExitCode != 0) throw new InvalidOperationException("O motor de separações não conseguiu ler o PDF.\n" + log);
                if (Regex.IsMatch(log, "too many.*(spot|color)|maximum.*(spot|color).*exceed", RegexOptions.IgnoreCase))
                    throw new InvalidOperationException("A página excedeu o limite de tintas especiais do motor. A prévia não foi exibida.");
                return log;
            }
        }

        // v2.7: renderiza so um retangulo da pagina com o tamanho exato em pixels da tela.
        // x0/y0 em pontos (1/72 pol), origem no canto inferior esquerdo da pagina.
        // metodo: 0 PageOffset, 1 BeginPage + translate, 2 Install + translate (sinal = direcao do
        // deslocamento), 3 PDF intermediario recortado com CropBox. O painel testa qual funciona
        // neste Ghostscript (MotorRegiao.Sondar). O PostScript e montado so com numeros formatados aqui.
        internal static string ExecutarRegiao(string executavel, string pdf, string diretorio, double dpi, int fator,
                                              double x0, double y0, int larguraPx, int alturaPx, int metodo, int sinal)
        {
            if (fator < 1) fator = 1;
            if (fator > 8) fator = 8;
            if (larguraPx < 1 || alturaPx < 1 || dpi <= 0) throw new ArgumentException("Área inválida.");
            double larguraPt = larguraPx * 72.0 / dpi, alturaPt = alturaPx * 72.0 / dpi;
            string entrada = pdf, deslocamento = "", log = "";
            if (metodo == 3)
            {
                entrada = Path.Combine(diretorio, "recorte.pdf");
                log += Rodar(executavel, diretorio,
                    "-dSAFER -dBATCH -dNOPAUSE -sDEVICE=pdfwrite -dFirstPage=1 -dLastPage=1 " +
                    "-sColorConversionStrategy=LeaveColorUnchanged -dAutoRotatePages=/None " +
                    "-dDownsampleColorImages=false -dDownsampleGrayImages=false -dDownsampleMonoImages=false " +
                    "-dAutoFilterColorImages=false -dAutoFilterGrayImages=false " +
                    "-sOutputFile=" + Aspas(entrada) +
                    " -c \"[/CropBox [" + Num(x0) + " " + Num(y0) + " " + Num(x0 + larguraPt) + " " + Num(y0 + alturaPt) +
                    "] /PAGES pdfmark\" -f " + Aspas(pdf));
            }
            else
            {
                double ox = (sinal == 2 || sinal == 3) ? x0 : -x0;
                double oy = (sinal == 1 || sinal == 3) ? y0 : -y0;
                string d = Num(ox) + " " + Num(oy);
                if (metodo == 0) deslocamento = " -c \"<</PageOffset [" + d + "]>> setpagedevice\"";
                else if (metodo == 1) deslocamento = " -c \"<</BeginPage {pop " + d + " translate}>> setpagedevice\"";
                else deslocamento = " -c \"<</Install {" + d + " translate}>> setpagedevice\"";
            }
            log += Rodar(executavel, diretorio,
                "-dSAFER -dBATCH -dNOPAUSE -sDEVICE=tiffsep -dPrintSpotCMYK=true " +
                "-dMaxSpots=60 -dFirstPage=1 -dLastPage=1 -dUseCropBox -sCompression=none -dOverprint=/enable " +
                "-dOverrideICC=true -dDeviceGrayToK=true -dFIXEDMEDIA " +
                "-g" + (larguraPx * fator).ToString(CultureInfo.InvariantCulture) + "x" +
                (alturaPx * fator).ToString(CultureInfo.InvariantCulture) +
                " -r" + Num(dpi * fator) +
                (fator > 1 ? " -dDownScaleFactor=" + fator.ToString(CultureInfo.InvariantCulture) : "") +
                " -sOutputFile=" + Aspas(Path.Combine(diretorio, "placas.tif")) + deslocamento + " -f " + Aspas(entrada));
            return log;
        }

        static string Rodar(string executavel, string diretorio, string argumentos)
        {
            ProcessStartInfo info = new ProcessStartInfo { FileName = executavel, UseShellExecute = false,
                CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true,
                WorkingDirectory = diretorio, StandardOutputEncoding = Encoding.UTF8, StandardErrorEncoding = Encoding.UTF8 };
            info.Arguments = argumentos;
            using (Process processo = Process.Start(info))
            {
                Task<string> saida = Task.Factory.StartNew(delegate { return processo.StandardOutput.ReadToEnd(); });
                Task<string> erro = Task.Factory.StartNew(delegate { return processo.StandardError.ReadToEnd(); });
                if (!processo.WaitForExit(120000))
                {
                    try { processo.Kill(); processo.WaitForExit(); } catch { }
                    throw new TimeoutException("A geração da área visível excedeu 2 minutos.");
                }
                string log = saida.Result + "\n" + erro.Result;
                if (processo.ExitCode != 0)
                    throw new InvalidOperationException("Ghostscript terminou com código " + processo.ExitCode + ".\n" + argumentos + "\n" + log);
                return log;
            }
        }

        static string Num(double valor)
        {
            if (double.IsNaN(valor) || double.IsInfinity(valor)) throw new ArgumentException("Medida inválida.");
            return valor.ToString("0.###", CultureInfo.InvariantCulture);
        }

        // Sem shell. Caminhos sao argumentos, nunca comandos ou trechos PostScript.
        internal static string Aspas(string valor)
        {
            if (valor.IndexOf('"') >= 0 || valor.IndexOf('\r') >= 0 || valor.IndexOf('\n') >= 0)
                throw new ArgumentException("Caminho inválido.");
            return "\"" + valor + "\"";
        }

        internal static void Limpar(string pasta)
        {
            if (String.IsNullOrEmpty(pasta)) return;
            // Somente diretorios temporarios gerados por este complemento.
            string nome = Path.GetFileName(pasta);
            if (!nome.StartsWith("Imprell_Separacoes_", StringComparison.Ordinal)) return;
            if (!String.Equals(Path.GetFullPath(Path.GetDirectoryName(pasta)).TrimEnd(Path.DirectorySeparatorChar),
                               Path.GetFullPath(Path.GetTempPath()).TrimEnd(Path.DirectorySeparatorChar), StringComparison.OrdinalIgnoreCase)) return;
            try { Directory.Delete(pasta, true); } catch { }
        }

#if !MOTOR_APENAS
        internal static string LocalizarExecutavel()
        {
            List<string> candidatos = new List<string>();
            string config = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Imprell", "SeparacoesCorel", "ghostscript.txt");
            if (File.Exists(config)) candidatos.Add(File.ReadAllText(config).Trim().Trim('"'));
            foreach (RegistryView view in new RegistryView[] { RegistryView.Registry64, RegistryView.Registry32 })
            {
                using (RegistryKey baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view))
                using (RegistryKey gs = baseKey.OpenSubKey(@"SOFTWARE\GPL Ghostscript"))
                {
                    if (gs == null) continue;
                    string[] versoes = gs.GetSubKeyNames(); Array.Sort(versoes, ComparerVersoes.Instance); Array.Reverse(versoes);
                    foreach (string v in versoes)
                    using (RegistryKey versao = gs.OpenSubKey(v))
                    {
                        string dll = Convert.ToString(versao.GetValue("GS_DLL"));
                        if (String.IsNullOrEmpty(dll)) continue;
                        string pasta = Path.GetDirectoryName(dll);
                        candidatos.Add(Path.Combine(pasta, "gswin64c.exe"));
                        candidatos.Add(Path.Combine(pasta, "gswin32c.exe"));
                    }
                }
            }
            foreach (string raiz in new string[] { Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                                                   Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86) })
            {
                if (String.IsNullOrEmpty(raiz)) continue;
                string pasta = Path.Combine(raiz, "gs");
                if (!Directory.Exists(pasta)) continue;
                string[] dirs = Directory.GetDirectories(pasta); Array.Sort(dirs, StringComparer.OrdinalIgnoreCase); Array.Reverse(dirs);
                foreach (string dir in dirs)
                {
                    candidatos.Add(Path.Combine(dir, "bin", "gswin64c.exe"));
                    candidatos.Add(Path.Combine(dir, "bin", "gswin32c.exe"));
                }
            }
            foreach (string candidato in candidatos)
            {
                string nome = Path.GetFileName(candidato);
                if ((nome.Equals("gswin64c.exe", StringComparison.OrdinalIgnoreCase) || nome.Equals("gswin32c.exe", StringComparison.OrdinalIgnoreCase)) && File.Exists(candidato))
                    return candidato;
            }
            throw new InvalidOperationException("Para separar CMYK e Pantone, instale o Ghostscript para Windows (64 bits) uma única vez: ghostscript.com/releases/gsdnld.html. Depois clique em Atualizar. Nenhuma tinta será convertida silenciosamente.");
        }
        sealed class ComparerVersoes : IComparer<string>
        {
            internal static readonly ComparerVersoes Instance = new ComparerVersoes();
            public int Compare(string a, string b)
            {
                Version x, y;
                if (Version.TryParse(a, out x) && Version.TryParse(b, out y)) return x.CompareTo(y);
                return String.CompareOrdinal(a, b);
            }
        }
#endif
    }

#if !MOTOR_APENAS
    internal static class LeitorSeparacoes
    {
        internal static ImagemCmyk Ler(Captura captura)
        {
            string log = MotorSeparacoes.Executar(captura.Motor, captura.Arquivo, captura.Diretorio, captura.Dpi, captura.Fator);
            return LerPasta(captura.Diretorio, log);
        }

        // v2.4: le as chapas ja geradas numa pasta (pagina inteira ou area visivel).
        internal static ImagemCmyk LerPasta(string diretorio, string log)
        {
            Dictionary<string, CorDaChapa> cores = MotorSeparacoes.LerCores(log);
            string[] processos = { "Cyan", "Magenta", "Yellow", "Black" };
            int w = 0, h = 0;
            byte[][] dados = new byte[4][];
            HashSet<string> usados = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < 4; i++)
            {
                string arquivo = LocalizarChapa(diretorio, processos[i]);
                dados[i] = LerCinza(arquivo, ref w, ref h);
                usados.Add(arquivo);
            }
            if (cores.Count > 60 || (long)w * h * (4 + cores.Count) > 320L * 1024 * 1024)
                throw new InvalidOperationException("Muitas chapas para esta resolução. Reduza a resolução e atualize.");
            byte[] cmyk = new byte[checked(w * h * 4)];
            for (int i = 0; i < w * h; i++)
                for (int canal = 0; canal < 4; canal++) cmyk[i * 4 + canal] = dados[canal][i];
            List<TintaEspecial> especiais = new List<TintaEspecial>();
            foreach (KeyValuePair<string, CorDaChapa> cor in cores)
            {
                if (Array.IndexOf(processos, cor.Key) >= 0) continue;
                string arquivo = LocalizarChapa(diretorio, cor.Key);
                if (!usados.Add(arquivo)) throw new InvalidOperationException("Duas tintas especiais resultaram no mesmo nome de arquivo: " + cor.Key);
                especiais.Add(new TintaEspecial { Nome = cor.Key, R = cor.Value.R, G = cor.Value.G, B = cor.Value.B,
                    Cobertura = LerCinza(arquivo, ref w, ref h) });
            }
            // Nao omitir uma tinta se o formato de metadados mudar numa versao do motor.
            foreach (string arquivo in Directory.GetFiles(diretorio, "placas*.tif"))
                if (!Path.GetFileName(arquivo).Equals("placas.tif", StringComparison.OrdinalIgnoreCase) && !usados.Contains(arquivo))
                    throw new InvalidOperationException("Uma chapa não pôde ser identificada: " + Path.GetFileName(arquivo) + ". A prévia foi interrompida para não omitir tinta.");
            especiais.Sort(delegate(TintaEspecial a, TintaEspecial b) { return String.CompareOrdinal(a.Nome, b.Nome); });
            return new ImagemCmyk(w, h, cmyk, especiais.ToArray());
        }

        static string LocalizarChapa(string pasta, string nome)
        {
            // Nomes Pantone usuais sao preservados. Nomes especiais que o motor
            // sanitizou sao rejeitados em vez de atribuir a chapa errada.
            if (nome.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
                throw new InvalidOperationException("O nome da tinta contém caracteres incompatíveis com o motor: " + nome);
            string a = Path.Combine(pasta, "placas(" + nome + ").tif");
            if (File.Exists(a)) return a;
            a = Path.Combine(pasta, "placas." + nome + ".tif");
            if (File.Exists(a)) return a;
            throw new InvalidOperationException("Não foi encontrada a chapa " + nome + ". A prévia não será mostrada incompleta.");
        }

        static byte[] LerCinza(string arquivo, ref int w, ref int h)
        {
            if (new FileInfo(arquivo).Length > 64L * 1024 * 1024) throw new InvalidOperationException("Chapa grande demais. Reduza a resolução.");
            using (FileStream fluxo = File.OpenRead(arquivo))
            {
                BitmapSource img = BitmapDecoder.Create(fluxo, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad).Frames[0];
                if ((long)img.PixelWidth * img.PixelHeight > ImagemCmyk.MaxPixels) throw new InvalidOperationException("Prévia grande demais. Reduza a resolução.");
                if (w == 0) { w = img.PixelWidth; h = img.PixelHeight; }
                if (w != img.PixelWidth || h != img.PixelHeight) throw new InvalidOperationException("As chapas têm dimensões diferentes.");
                if (img.Format != PixelFormats.Gray8) img = new FormatConvertedBitmap(img, PixelFormats.Gray8, null, 0);
                byte[] dados = new byte[checked(w * h)]; img.CopyPixels(dados, w, 0);
                for (int i = 0; i < dados.Length; i++) dados[i] = (byte)(255 - dados[i]);
                return dados;
            }
        }
    }

    // v2.7: detalhe da area visivel na resolucao da tela, gerado a partir do PDF ja exportado.
    // Na primeira vez, testa num PDF pequeno qual jeito de recortar a pagina funciona neste
    // Ghostscript e guarda o resultado. Tudo fica registrado em diagnostico.log.
    internal static class MotorRegiao
    {
        static volatile int metodo = -2; // -2 ainda nao testado; -1 nenhum funciona; 0 a 3 metodo em uso
        static int sinal;
        static string motorTestado;
        static readonly object trava = new object();

        static string Pasta
        {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Imprell", "SeparacoesCorel"); }
        }

        internal static bool Disponivel() { return metodo != -1; }

        internal static void Diagnostico(string texto)
        {
            try
            {
                Directory.CreateDirectory(Pasta);
                string arquivo = Path.Combine(Pasta, "diagnostico.log");
                if (File.Exists(arquivo) && new FileInfo(arquivo).Length > 512 * 1024) File.Delete(arquivo);
                File.AppendAllText(arquivo, DateTime.Now.ToString("s") + " " + texto + Environment.NewLine);
            }
            catch { }
        }

        static string NovaPasta()
        {
            string pasta = Path.Combine(Path.GetTempPath(), "Imprell_Separacoes_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(pasta);
            return pasta;
        }

        static bool GarantirMetodo(string executavel)
        {
            lock (trava)
            {
                if (motorTestado == executavel && metodo != -2) return metodo >= 0;
                int m, s;
                if (LerConfig(executavel, out m, out s))
                {
                    sinal = s; motorTestado = executavel; metodo = m;
                    return true;
                }
                Diagnostico("Testando como este Ghostscript gera só uma área da página: " + executavel);
                if (Sondar(executavel, out m, out s))
                {
                    sinal = s; motorTestado = executavel; metodo = m;
                    GravarConfig(executavel, m, s);
                    Diagnostico("Funcionou o método " + m + " (sinal " + s + "). A área visível será gerada assim.");
                    return true;
                }
                motorTestado = executavel; metodo = -1;
                Diagnostico("Nenhum método funcionou. A prévia usa só a página inteira.");
                return false;
            }
        }

        static bool LerConfig(string executavel, out int m, out int s)
        {
            m = -1; s = 0;
            try
            {
                string arquivo = Path.Combine(Pasta, "regiao-v27.txt");
                if (!File.Exists(arquivo)) return false;
                string[] linhas = File.ReadAllLines(arquivo);
                if (linhas.Length < 2 || !String.Equals(linhas[1].Trim(), executavel, StringComparison.OrdinalIgnoreCase)) return false;
                string[] partes = linhas[0].Trim().Split(' ');
                if (partes.Length != 2) return false;
                if (!Int32.TryParse(partes[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out m) ||
                    !Int32.TryParse(partes[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out s)) return false;
                return m >= 0 && m <= 3 && s >= 0 && s <= 3;
            }
            catch { return false; }
        }

        static void GravarConfig(string executavel, int m, int s)
        {
            try
            {
                Directory.CreateDirectory(Pasta);
                File.WriteAllText(Path.Combine(Pasta, "regiao-v27.txt"),
                    m.ToString(CultureInfo.InvariantCulture) + " " + s.ToString(CultureInfo.InvariantCulture) + Environment.NewLine + executavel);
            }
            catch { }
        }

        // PDF de 200 x 200 pt com um quadrado 100% K em x 50-70, y 120-140 (origem embaixo).
        // A area pedida e x 40-80, y 110-150: o quadrado tem que cair nos pixels 10 a 30 (72 dpi).
        static bool Sondar(string executavel, out int metodoOk, out int sinalOk)
        {
            metodoOk = -1; sinalOk = 0;
            string pasta = NovaPasta();
            try
            {
                string pdf = Path.Combine(pasta, "teste.pdf");
                CriarPdfTeste(pdf);
                for (int m = 0; m <= 3; m++)
                {
                    int sinais = m == 3 ? 1 : 4;
                    for (int s = 0; s < sinais; s++)
                    {
                        string saida = NovaPasta();
                        try
                        {
                            string log = MotorSeparacoes.ExecutarRegiao(executavel, pdf, saida, 72, 1, 40, 110, 40, 40, m, s);
                            ImagemCmyk img = LeitorSeparacoes.LerPasta(saida, log);
                            int dentro = PretoEm(img, 20, 20);
                            int fora = Math.Max(Math.Max(PretoEm(img, 3, 3), PretoEm(img, 36, 3)), Math.Max(PretoEm(img, 3, 36), PretoEm(img, 36, 36)));
                            Diagnostico("Método " + m + " sinal " + s + ": " + img.Largura + "x" + img.Altura + " px, preto dentro " + dentro + ", fora " + fora);
                            if (img.Largura == 40 && img.Altura == 40 && dentro >= 200 && fora <= 40)
                            {
                                metodoOk = m; sinalOk = s;
                                return true;
                            }
                        }
                        catch (Exception ex) { Diagnostico("Método " + m + " sinal " + s + ": erro " + ex.Message); }
                        finally { MotorSeparacoes.Limpar(saida); }
                    }
                }
                return false;
            }
            finally { MotorSeparacoes.Limpar(pasta); }
        }

        static int PretoEm(ImagemCmyk img, int x, int y)
        {
            if (x < 0 || y < 0 || x >= img.Largura || y >= img.Altura) return 255;
            return img.Pixels[(y * img.Largura + x) * 4 + 3];
        }

        static void CriarPdfTeste(string arquivo)
        {
            string conteudo = "0 0 0 1 k 50 120 20 20 re f\n";
            string[] objetos = {
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /Resources << >> >>",
                "<< /Length " + conteudo.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n" + conteudo + "endstream" };
            StringBuilder pdf = new StringBuilder("%PDF-1.4\n");
            List<int> posicoes = new List<int>();
            for (int i = 0; i < objetos.Length; i++)
            {
                posicoes.Add(pdf.Length);
                pdf.Append((i + 1).ToString(CultureInfo.InvariantCulture) + " 0 obj\n" + objetos[i] + "\nendobj\n");
            }
            int xref = pdf.Length;
            pdf.Append("xref\n0 " + (objetos.Length + 1).ToString(CultureInfo.InvariantCulture) + "\n0000000000 65535 f \n");
            foreach (int p in posicoes) pdf.Append(p.ToString("D10", CultureInfo.InvariantCulture) + " 00000 n \n");
            pdf.Append("trailer\n<< /Size " + (objetos.Length + 1).ToString(CultureInfo.InvariantCulture) + " /Root 1 0 R >>\nstartxref\n" +
                       xref.ToString(CultureInfo.InvariantCulture) + "\n%%EOF\n");
            File.WriteAllText(arquivo, pdf.ToString(), Encoding.ASCII);
        }

        internal static ImagemCmyk Renderizar(string executavel, string pdf, ImagemCmyk pagina, double dpi, int fator,
            double x0, double y0, int larguraPx, int alturaPx, double u0, double v0, double u1, double v1)
        {
            if (!GarantirMetodo(executavel)) return null;
            string pasta = NovaPasta();
            try
            {
                string log = MotorSeparacoes.ExecutarRegiao(executavel, pdf, pasta, dpi, fator, x0, y0, larguraPx, alturaPx, metodo, sinal);
                ImagemCmyk regiao = LeitorSeparacoes.LerPasta(pasta, log);
                if (regiao.Especiais.Length != pagina.Especiais.Length) return null;
                bool informativo;
                double diferenca = Diferenca(pagina, regiao, u0, v0, u1, v1, out informativo);
                if (informativo && diferenca > 25)
                {
                    Diagnostico("Área visível diferente da página inteira (diferença " +
                                diferenca.ToString("0.0", CultureInfo.InvariantCulture) + "); descartada.");
                    return null;
                }
                return regiao;
            }
            catch (Exception ex) { Corel.Registrar(ex); Diagnostico("Erro ao gerar a área visível: " + ex.Message); return null; }
            finally { MotorSeparacoes.Limpar(pasta); }
        }

        // Media da diferenca por canal (0 a 255) entre cada pixel da pagina inteira e a media
        // da area correspondente no detalhe. informativo = o trecho tem desenho para comparar.
        internal static double Diferenca(ImagemCmyk pagina, ImagemCmyk regiao, double u0, double v0, double u1, double v1,
                                         out bool informativo)
        {
            informativo = false;
            if (u1 <= u0 || v1 <= v0) return 0;
            int bx0 = Math.Max(0, (int)Math.Ceiling(u0 * pagina.Largura));
            int bx1 = Math.Min(pagina.Largura, (int)Math.Floor(u1 * pagina.Largura));
            int by0 = Math.Max(0, (int)Math.Ceiling(v0 * pagina.Altura));
            int by1 = Math.Min(pagina.Altura, (int)Math.Floor(v1 * pagina.Altura));
            if (bx1 - bx0 < 2 || by1 - by0 < 2) return 0;
            double sx = regiao.Largura / ((u1 - u0) * pagina.Largura);
            double sy = regiao.Altura / ((v1 - v0) * pagina.Altura);
            long total = (long)(bx1 - bx0) * (by1 - by0);
            int passo = (int)Math.Max(1, Math.Ceiling(Math.Sqrt(total / 20000.0)));
            int sub = (int)Math.Max(1, Math.Ceiling(Math.Max(sx, sy) / 8.0));
            byte[] p = pagina.Pixels, r = regiao.Pixels;
            double soma = 0;
            long n = 0, comTinta = 0;
            for (int by = by0; by < by1; by += passo)
            {
                for (int bx = bx0; bx < bx1; bx += passo)
                {
                    double rx = ((double)bx / pagina.Largura - u0) / (u1 - u0) * regiao.Largura;
                    double ry = ((double)by / pagina.Altura - v0) / (v1 - v0) * regiao.Altura;
                    int ix0 = Math.Max(0, (int)Math.Floor(rx)), iy0 = Math.Max(0, (int)Math.Floor(ry));
                    int ix1 = Math.Min(regiao.Largura, (int)Math.Ceiling(rx + sx));
                    int iy1 = Math.Min(regiao.Altura, (int)Math.Ceiling(ry + sy));
                    if (ix1 <= ix0 || iy1 <= iy0) continue;
                    long a0 = 0, a1 = 0, a2 = 0, a3 = 0;
                    int m = 0;
                    for (int y = iy0; y < iy1; y += sub)
                    {
                        for (int x = ix0; x < ix1; x += sub)
                        {
                            int i = (y * regiao.Largura + x) * 4;
                            a0 += r[i]; a1 += r[i + 1]; a2 += r[i + 2]; a3 += r[i + 3]; m++;
                        }
                    }
                    if (m == 0) continue;
                    int j = (by * pagina.Largura + bx) * 4;
                    soma += (Math.Abs((double)a0 / m - p[j]) + Math.Abs((double)a1 / m - p[j + 1]) +
                             Math.Abs((double)a2 / m - p[j + 2]) + Math.Abs((double)a3 / m - p[j + 3])) / 4.0;
                    n++;
                    if (p[j] + p[j + 1] + p[j + 2] + p[j + 3] > 60) comTinta++;
                }
            }
            if (n == 0) return 0;
            double fracao = (double)comTinta / n;
            informativo = fracao >= 0.05 && fracao <= 0.95;
            return soma / n;
        }
    }
#endif
}
