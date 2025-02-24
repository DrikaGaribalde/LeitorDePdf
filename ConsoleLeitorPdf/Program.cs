// See https://aka.ms/new-console-template for more information
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using System;
using System.IO;
using System.Text.RegularExpressions;

class Program
{
    static void Main(string[] args)
    {
        // Caminho do arquivo PDF
        string caminhoPDF = @"C:\Cursos\pdf\relatorio.pdf";

        string caminhoCSV = @"C:\Cursos\pdf\relatorio_saida.csv";

        // Verificar se o arquivo existe
        if (File.Exists(caminhoPDF))
        {
            try
            {
                string textoCompleto = "";
                // Abrir o PDF
                using (PdfReader pdfReader = new PdfReader(caminhoPDF))
                using (PdfDocument pdfDocument = new PdfDocument(pdfReader))
                {
                   

                    // Ler todas as páginas do PDF
                    for (int i = 1; i <= pdfDocument.GetNumberOfPages(); i++)
                    {
                        // Extrair o texto de cada página
                        textoCompleto += PdfTextExtractor.GetTextFromPage(pdfDocument.GetPage(i));
                    }
                }
                // Extrair as informações-chave da nota fiscal usando regex
                string chaveAcesso = Regex.Match(textoCompleto, @"Chave de Acesso da NFS-e[:\s]+([\d]+)").Groups[1].Value;
                string dataHoraEmissao = Regex.Match(textoCompleto, @"Data e Hora de Emissão[:\s]+([\d/]+ [\d.]+)").Groups[1].Value;
                string competenciaNFS = Regex.Match(textoCompleto, @"Competência da NFS-e[:\s]+([\d/]+)").Groups[1].Value;
                string localPrestacao = Regex.Match(textoCompleto, @"Local da Prestação[:\s]+([A-Z\s-]+)").Groups[1].Value;
                string paisPrestacao = Regex.Match(textoCompleto, @"País de Prestação[:\s]+([A-Za-z]+)").Groups[1].Value;
                string dataHoraDPS = Regex.Match(textoCompleto, @"Data e Hora da emissão da DPS[:\s]+([\d/]+ [\d:]+)").Groups[1].Value;
                string numeroDPS = Regex.Match(textoCompleto, @"Número da DPS[:\s]+(\d+)").Groups[1].Value;
                string serieDPS = Regex.Match(textoCompleto, @"Série da DPS[:\s]+(\d+)").Groups[1].Value;
                string razaoSocial = Regex.Match(textoCompleto, @"Razão Social[:\s]+([A-Z\s]+)").Groups[1].Value;
                string cnpj = Regex.Match(textoCompleto, @"CNPJ/CPF/NIF[:\s]+([\d./-]+)").Groups[1].Value;
                string endereco = Regex.Match(textoCompleto, @"Endereço[:\s]+([A-Z0-9\s-]+)").Groups[1].Value;
                string email = Regex.Match(textoCompleto, @"E-mail[:\s]+([a-zA-Z0-9@.]+)").Groups[1].Value;

                // Verificar se as informações foram encontradas
                if (string.IsNullOrEmpty(chaveAcesso) || string.IsNullOrEmpty(dataHoraEmissao) || string.IsNullOrEmpty(cnpj))
                {
                    Console.WriteLine("Não foi possível extrair todas as informações.");
                    return;
                }

                // Gerar o CSV com chave-valor
                using (StreamWriter writer = new StreamWriter(caminhoCSV))
                {
                    writer.WriteLine("Chave,Valor");
                    writer.WriteLine($"Chave de Acesso da NFS-e,{chaveAcesso}");
                    writer.WriteLine($"Data e Hora de Emissão,{dataHoraEmissao}");
                    writer.WriteLine($"Competência da NFS-e,{competenciaNFS}");
                    writer.WriteLine($"Local da Prestação,{localPrestacao}");
                    writer.WriteLine($"País de Prestação,{paisPrestacao}");
                    writer.WriteLine($"Data e Hora da emissão da DPS,{dataHoraDPS}");
                    writer.WriteLine($"Número da DPS,{numeroDPS}");
                    writer.WriteLine($"Série da DPS,{serieDPS}");
                    writer.WriteLine($"Razão Social,{razaoSocial}");
                    writer.WriteLine($"CNPJ/CPF/NIF,{cnpj}");
                    writer.WriteLine($"Endereço,{endereco}");
                    writer.WriteLine($"E-mail,{email}");
                }

                Console.WriteLine($"CSV gerado com sucesso em: {caminhoCSV}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Erro ao processar o PDF: " + ex.Message);
            }
        }
        else
        {
            Console.WriteLine("O arquivo PDF não foi encontrado.");
        }
    }
}

