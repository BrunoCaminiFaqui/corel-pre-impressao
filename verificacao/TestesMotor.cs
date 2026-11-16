using System;
using System.Linq;
using System.IO;
using SeparacoesCorel;

internal static class TestesMotor
{
    static int testes;
    static void Exigir(bool ok, string nome)
    { testes++; if (!ok) throw new Exception("FALHOU: " + nome); }
    static byte[] Pixel(int c, int m, int y, int k)
    { return new byte[] { (byte)c, (byte)m, (byte)y, (byte)k }; }
    static bool[] Todas() { return new bool[] { true, true, true, true }; }
    static void Cor(byte[] p, int r, int g, int b, string nome)
    { Exigir(p[2] == r && p[1] == g && p[0] == b && p[3] == 255, nome); }
    static void Invalido(Action acao, string nome)
    {
        bool rejeitado = false;
        try { acao(); } catch (ArgumentException) { rejeitado = true; }
        Exigir(rejeitado, nome);
    }
    static void Main(string[] args)
    {
        Cor(Renderizador.Desenhar(new ImagemCmyk(1, 1, Pixel(0, 0, 0, 0)), Todas(), false, false, 300), 255, 255, 255, "papel branco");
        Cor(Renderizador.Desenhar(new ImagemCmyk(1, 1, Pixel(255, 0, 0, 0)), Todas(), false, false, 300), 0, 255, 255, "ciano");
        Cor(Renderizador.Desenhar(new ImagemCmyk(1, 1, Pixel(0, 255, 0, 0)), Todas(), false, false, 300), 255, 0, 255, "magenta");
        Cor(Renderizador.Desenhar(new ImagemCmyk(1, 1, Pixel(0, 0, 255, 0)), Todas(), false, false, 300), 255, 255, 0, "amarelo");
        Cor(Renderizador.Desenhar(new ImagemCmyk(1, 1, Pixel(0, 0, 0, 255)), Todas(), false, false, 300), 0, 0, 0, "preto");
        for (int mask = 0; mask < 16; mask++)
        {
            bool[] v = new bool[4];
            for (int i = 0; i < 4; i++) v[i] = (mask & (1 << i)) != 0;
            byte[] p = Renderizador.Desenhar(new ImagemCmyk(1, 1, Pixel(100, 150, 200, 50)), v, false, false, 300);
            int k = v[3] ? 50 : 0;
            Cor(p, (v[0] ? 155 : 255) * (255 - k) / 255, (v[1] ? 105 : 255) * (255 - k) / 255,
                (v[2] ? 55 : 255) * (255 - k) / 255, "combinacao de chapas " + mask);
        }
        for (int chapa = 0; chapa < 4; chapa++)
        {
            bool[] v = new bool[4]; v[chapa] = true;
            Cor(Renderizador.Desenhar(new ImagemCmyk(1, 1, Pixel(128, 128, 128, 128)), v, true, false, 300),
                127, 127, 127, "chapa cinza " + chapa);
        }
        ImagemCmyk limite = new ImagemCmyk(1, 1, Pixel(255, 255, 255, 0));
        Cor(Renderizador.Desenhar(limite, Todas(), false, true, 300), 0, 0, 0, "limite exato nao destaca");
        ImagemCmyk acima = new ImagemCmyk(1, 1, Pixel(255, 255, 255, 1));
        Cor(Renderizador.Desenhar(acima, new bool[4], false, true, 300), 0, 230, 0, "carga usa chapas ocultas");
        int[] amostra = acima.Amostra(0, 0);
        Exigir(amostra[0] == 100 && amostra[4] == 300, "leitura arredondada");
        Exigir(acima.Amostra(-1, 0) == null && acima.Amostra(0, 1) == null, "amostra fora da imagem");
        ImagemCmyk dois = new ImagemCmyk(2, 1, new byte[] { 255, 0, 0, 0, 0, 0, 0, 255 });
        Exigir(dois.Amostra(1, 0)[3] == 100 && dois.Amostra(1, 0)[0] == 0, "ordem dos pixels");
        int x, y;
        Exigir(Renderizador.LocalizarPixel(100, 100, 200, 200, 400, 200, out x, out y) && x == 200 && y == 100, "amostra com margens");
        Exigir(!Renderizador.LocalizarPixel(100, 49, 200, 200, 400, 200, out x, out y), "margem superior");
        Exigir(!Renderizador.LocalizarPixel(200, 100, 200, 200, 400, 200, out x, out y), "borda direita");
        Exigir(!Renderizador.LocalizarPixel(0, 0, 0, 100, 400, 200, out x, out y), "painel sem largura");
        Invalido(delegate { new ImagemCmyk(0, 1, new byte[0]); }, "dimensao zero");
        Invalido(delegate { new ImagemCmyk(int.MaxValue, int.MaxValue, new byte[0]); }, "overflow");
        Invalido(delegate { new ImagemCmyk(1, 1, new byte[3]); }, "buffer incompleto");
        Invalido(delegate { new ImagemCmyk(2001, 2000, new byte[0]); }, "limite de memoria");
        Invalido(delegate { Renderizador.Desenhar(acima, new bool[3], false, false, 300); }, "chapas incompletas");
        Invalido(delegate { Renderizador.Desenhar(acima, Todas(), false, false, 401); }, "limite invalido");
        TestarEspeciais();
        TestarPdf();
        if (args.Length == 3)
        {
            Directory.CreateDirectory(args[2]);
            string log = MotorSeparacoes.Executar(args[0], Path.GetFullPath(args[1]), Path.GetFullPath(args[2]), 72);
            var cores = MotorSeparacoes.LerCores(log);
            Exigir(cores.Count == 2 && cores.ContainsKey("PANTONE 186 C") && cores.ContainsKey("PANTONE 871 C"), "motor externo gerou duas tintas nomeadas");
            Exigir(File.Exists(Path.Combine(args[2], "placas(PANTONE 186 C).tif")) && File.Exists(Path.Combine(args[2], "placas(PANTONE 871 C).tif")), "arquivos de chapas do processo externo");
        }
        Console.WriteLine("PASSOU: " + testes + " verificacoes CMYK, especiais, exportacao e motor externo (quando solicitado).");
    }
    static void TestarEspeciais()
    {
        var vermelho = new TintaEspecial { Nome = "PANTONE 186 C", R = 255, G = 0, B = 102, Cobertura = new byte[] { 255 } };
        var dourado = new TintaEspecial { Nome = "PANTONE 871 C", R = 184, G = 161, B = 46, Cobertura = new byte[] { 255 } };
        var tinta = new ImagemCmyk(1, 1, Pixel(0, 0, 0, 0), new[] { vermelho, dourado });
        Cor(Renderizador.Desenhar(tinta, Todas(), false, false, 300, new[] { false, false }), 255, 255, 255, "desligar ambos os Pantones apaga apenas suas tintas");
        Cor(Renderizador.Desenhar(tinta, new bool[4], false, false, 300, new[] { true, false }), 255, 0, 102, "isolar Pantone vermelho");
        Cor(Renderizador.Desenhar(tinta, new bool[4], false, false, 300, new[] { false, true }), 184, 161, 46, "isolar Pantone dourado");
        Cor(Renderizador.Desenhar(tinta, Todas(), true, false, 300, new[] { true, false }), 0, 0, 0, "chapa Pantone em preto");
        Exigir(tinta.Amostra(0, 0)[4] == 0, "Pantone nao contamina o CMYK");
        var cianoSpot = new ImagemCmyk(1, 1, Pixel(255, 0, 0, 0), new[] { vermelho });
        Cor(Renderizador.Desenhar(cianoSpot, Todas(), false, false, 300, new[] { false }), 0, 255, 255, "ocultar Pantone preserva ciano");
        vermelho.Cobertura = new byte[] { 128 };
        Cor(Renderizador.Desenhar(tinta, new bool[4], true, false, 300, new[] { true, false }), 127, 127, 127, "reticula especial 50 por cento");
        vermelho.Cobertura = new byte[] { 0 };
        Cor(Renderizador.Desenhar(tinta, new bool[4], false, false, 300, new[] { true, false }), 255, 255, 255, "especial sem cobertura");
        Invalido(delegate { Renderizador.Desenhar(tinta, Todas(), false, false, 300, new bool[1]); }, "visibilidades especiais incompletas");
        Invalido(delegate { new ImagemCmyk(2, 1, new byte[8], new[] { vermelho }); }, "chapa especial de tamanho errado");
        string log = "GPL Ghostscript\n%%SeparationColor: \"PANTONE 186 C\" 100% ink = 0 32760 19656 0 CMYK\n%%SeparationColor: \"PANTONE 871 C\" 100% ink = 6552 9828 26208 3276 CMYK";
        var cores = MotorSeparacoes.LerCores(log);
        Exigir(cores.Count == 2 && cores["PANTONE 186 C"].R == 255 && cores["PANTONE 186 C"].B == 102, "leitura de metadados reais Ghostscript");
        Exigir(MotorSeparacoes.LerCores("sem especiais").Count == 0, "pagina apenas CMYK");
        Invalido(delegate { MotorSeparacoes.Aspas("arquivo\"incorreto"); }, "argumentos com aspas rejeitados");
    }

    static void TestarPdf()
    {
        foreach (bool falhar in new[] { false, true })
        {
            var doc = new DocumentoPdfFalso { Falhar = falhar };
            var props = typeof(OpcoesPdfFalsas).GetProperties();
            foreach (var prop in props) prop.SetValue(doc.PDFSettings, prop.PropertyType == typeof(bool) ? (object)true : 2, null);
            object[] antes = props.Select(p => p.GetValue(doc.PDFSettings, null)).ToArray();
            bool falhou = false;
            try { ExportacaoPdf.Exportar(doc, "apenas-teste.pdf"); }
            catch (InvalidOperationException) { falhou = true; }
            Exigir(falhou == falhar && doc.Publicado, "exportacao propaga falha " + falhar);
            Exigir(props.Select(p => p.GetValue(doc.PDFSettings, null)).SequenceEqual(antes), "todas as opcoes restauradas apos exportacao " + falhar);
            Exigir(doc.PreservouEspeciais && doc.PaginaAtual, "PDF preserva especiais e pagina atual " + falhar);
        }
    }

}

public sealed class DocumentoPdfFalso
{
    public OpcoesPdfFalsas PDFSettings { get; set; }
    public bool Falhar, Publicado, PreservouEspeciais, PaginaAtual;
    public DocumentoPdfFalso() { PDFSettings = new OpcoesPdfFalsas(); }
    public void PublishToPDF(string caminho)
    {
        Publicado = true;
        PreservouEspeciais = PDFSettings.OutputSpotColorsAs == 0 && PDFSettings.Overprints && !PDFSettings.ComplexFillsAsBitmaps;
        PaginaAtual = PDFSettings.PublishRange == 1;
        if (Falhar) throw new IOException("Falha de exportacao injetada");
    }
}
public sealed class OpcoesPdfFalsas
{
    public int pdfVersion { get; set; }
    public int PublishRange { get; set; }
    public int ColorMode { get; set; }
    public int OutputSpotColorsAs { get; set; }
    public bool Overprints { get; set; }
    public bool ComplexFillsAsBitmaps { get; set; }
    public bool TextAsCurves { get; set; }
    public bool IncludeBleed { get; set; }
    public bool CropMarks { get; set; }
    public bool RegistrationMarks { get; set; }
    public bool DensitometerScales { get; set; }
    public bool FileInformation { get; set; }
    public bool DownsampleColor { get; set; }
    public bool DownsampleGray { get; set; }
    public bool DownsampleMono { get; set; }
    public int EncryptType { get; set; }
}
