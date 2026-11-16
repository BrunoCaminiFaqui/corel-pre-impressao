# Ferramentas de pré-impressão para CorelDRAW

Painéis e macros para conferir e preparar artes para impressão offset no CorelDRAW.

by **Bruno Faqui**

## O que tem aqui

### Painel "Separações CMYK + Pantone"
Docker (janela de encaixe) que mostra a prévia das separações da página:

- liga e desliga Ciano, Magenta, Amarelo, Preto e cada cor especial (Pantone, verniz, branco...);
- Shift + clique isola uma chapa;
- leitura das porcentagens de tinta onde o mouse passa;
- **Conferência de cores**: lista o que está em RGB (objetos e imagens), a cor de registro e o preto/cinza
  que não é só K. Tem filtros por tipo e os botões Anterior / Próximo, que selecionam cada item e dão zoom nele,
  inclusive dentro de PowerClip.

### Painel "Trocar cores"
Docker que substitui o Localizar e substituir do Corel:

- lista as cores usadas na arte, inclusive dentro de grupos e PowerClips;
- **Selecionar** marca no Corel os objetos com a cor;
- **Trocar** troca a cor em todos os lugares, num passo só de desfazer;
- lista as imagens fora do CMYK (RGB, Lab, paleta) com Anterior / Próxima e zoom;
- **Converter RGB para CMYK**: converte de uma vez as cores e imagens RGB, na página atual ou em todas as páginas.

### Macros (VBA)
- `macros/Sangria.bas` – **sangria pelo contorno da faca**: cria o contorno externo da navalha com a medida da
  sangria, separa (Ctrl+K) e solda na moldura do PowerClip, sem deformar e sem tirar o conteúdo do PowerClip.
  Funciona em todas as páginas ou nas escolhidas.
- `macros/TrocarCor.bas` – troca uma cor por outra sem usar o Localizar e substituir.

## Requisitos

- Windows e CorelDRAW 64 bits (preparado para o CorelDRAW 2021; os painéis usam recursos do X7 em diante);
- .NET Framework 4 (já vem no Windows);
- [Ghostscript 64 bits](https://www.ghostscript.com/releases/gsdnld.html), edição AGPL – usado pelo painel de
  separações para gerar a prévia. Não acompanha este repositório.

## Instalar os painéis

1. Instale o Ghostscript (uma vez só).
2. Baixe o repositório (Code > Download ZIP) e extraia **tudo** para uma pasta.
3. Salve as artes e feche o CorelDRAW.
4. Botão direito em `Instalar.cmd` > **Executar como administrador**.
   O script compila os dois painéis a partir da pasta `fonte/` (com o compilador do .NET que já vem no Windows)
   e instala no Corel. Se a compilação falhar, nada é instalado.
5. No Corel: **Janela > Janelas de encaixe > Separações CMYK + Pantone** e **Trocar cores**.

Para remover: `Desinstalar.cmd` como administrador.
Registro da instalação: `Ver_registro.cmd`.

## Instalar as macros

1. No Corel, abra o editor VBA (**Alt+F11**).
2. Em um projeto GMS: **Arquivo > Importar arquivo...** e escolha o `.bas`.
3. Rode pela lista de macros (**Alt+F8**): `SangrarPeloContorno` ou `TrocarCor`.

## Estrutura

| Pasta / arquivo | Conteúdo |
|---|---|
| `fonte/` | Código C# dos painéis (`Separacoes.cs`, `MotorSeparacoes.cs`, `TrocarCores.cs`) |
| `painel/`, `painel_cores/` | Arquivos de interface do Corel (`.xslt`, `.addon`) de cada painel |
| `macros/` | Macros VBA para importar no Corel |
| `verificacao/` | Testes do motor de separação |
| `Instalar.cmd`, `Instalar.ps1`, `InstalarCores.ps1`, `Recompilar.*` | Instalação e compilação |
| `LEIA_PRIMEIRO.txt` | Manual detalhado do painel de separações |
| `HISTORICO.txt` | O que mudou em cada versão |

## Observações

- A prévia é uma simulação para conferência; não substitui a prova de cor nem o RIP.
- Os registros ficam em `%LOCALAPPDATA%\Imprell\` (`SeparacoesCorel` e `TrocarCores`).
