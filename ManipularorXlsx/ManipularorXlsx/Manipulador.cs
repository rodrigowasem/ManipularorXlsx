using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using ClosedXML.Excel;
using System.Linq;

namespace ManipularorXlsx
{
    public class Manipulador
    {
        private List<Empresa> Empresas;

        public Manipulador()
        {
            Empresas = new List<Empresa>();
        }

        public bool CarregaDadosXlsx(string diretorio)
        {
            try
            {
                var cultureInfo = new CultureInfo("pt-BR");
                var numberFormatInfo = (NumberFormatInfo)cultureInfo.NumberFormat.Clone();
                numberFormatInfo.CurrencySymbol = "";

                // Abre arquivo Excel Excel Existente
                var wb = new XLWorkbook(diretorio);

                // seta com qual planilha do aquivo vamos trabalhar
                var planilha = wb.Worksheet(1);

                // Inicia a leitura do arquivo na linha 2
                var linha = 2;

                while (true)
                {
                    var nome = planilha.Cell("A" + linha.ToString()).Value.ToString();

                    if (string.IsNullOrEmpty(nome)) break;

                    Empresas.Add(
                        new Empresa
                        {
                            Nome = nome,
                            Codigo = planilha.Cell("B" + linha.ToString()).Value.ToString(),
                            Capital = Decimal.Parse(planilha.Cell("F" + linha.ToString()).Value.ToString())
                        }
                   );

                    linha++;

                }

                return true;
            }
            catch (Exception e)
            {
                //throw new Exception("Erro ao salvar o arquivo com os 100 maiores capitais!", e);
                Console.WriteLine("Erro ao carregar arquivo .xlsx!");
                return false;
            }
        }

        public void ExibirDadosDoArquivo() 
        {

            // Monta as colunas e seus respectivos títulos
            Console.WriteLine("".PadRight(80, '-'));
            Console.WriteLine("Nome".PadRight(35) + "Código".PadRight(15) + "Capital".PadRight(30));
            Console.WriteLine("".PadRight(80, '-'));

            foreach (var empresa in Empresas)
            {
                
                Console.Write(empresa.Nome.PadRight(35));
                Console.Write(empresa.Codigo.PadRight(15));
                Console.WriteLine(empresa.Capital.ToString().PadRight(30));

            }

            Console.WriteLine("".PadRight(80, '-'));
            Console.WriteLine("Dados carregados com sucesso!!!".PadRight('-'));

        }

        public void ExibirCemMaioresCapitais()
        {
            var cultureInfo = new CultureInfo("pt-BR");
            var numberFormatInfo = (NumberFormatInfo)cultureInfo.NumberFormat.Clone();
            numberFormatInfo.CurrencySymbol = "";

            // Monta as colunas e seus respectivos títulos
            Console.WriteLine("100 Maiores Capitais Sociais da B3");
            Console.WriteLine("".PadRight(90, '-'));
            Console.WriteLine("Posição".PadRight(10) + "Nome".PadRight(35) + "Código".PadRight(15) + "Capital".PadRight(30));
            Console.WriteLine("".PadRight(90, '-'));

            var cemMaiores = Empresas.OrderByDescending(emp => emp.Capital).Take(100);

            int linha = 1;

            foreach (var empresa in cemMaiores)
            {

                Console.Write((linha).ToString().PadRight(10));
                Console.Write(empresa.Nome.PadRight(35));
                Console.Write(empresa.Codigo.PadRight(15));
                Console.WriteLine(empresa.Capital.ToString("#,0.00", numberFormatInfo).PadRight(30));

                linha++;
            }           

            Console.WriteLine("".PadRight(90, '-'));
            Console.WriteLine("Dados carregados com sucesso!!!".PadRight('-'));
            Console.WriteLine("Deseja salvar essa consulta?");

            if (Console.ReadLine().ToUpper() == "S")
            {
                GeraXlsx(cemMaiores);
            }
            else
            {
                return;
            }
        }

        public void GeraXlsx(IEnumerable<Empresa> empresas) 
        {
            var cultureInfo = new CultureInfo("pt-BR");
            var numberFormatInfo = (NumberFormatInfo)cultureInfo.NumberFormat.Clone();
            numberFormatInfo.CurrencySymbol = "";

            try
            {
                var wbNovo = new XLWorkbook();

                var ws = wbNovo.Worksheets.Add("100 Maiores");

                var linhaNovo = 2;

                ws.Cell("A1").Value = "Nome";
                ws.Cell("B1").Value = "Codigo";
                ws.Cell("C1").Value = "Capital";

                foreach (var empresa in empresas)
                {
                    ws.Cell("A" + linhaNovo).Value = empresa.Nome;
                    ws.Cell("B" + linhaNovo).Value = empresa.Codigo;
                    ws.Cell("C" + linhaNovo).Value = empresa.Capital.ToString("#,0.00", numberFormatInfo);

                    linhaNovo++;
                }

                wbNovo.SaveAs("testerodrigo.xlsx");
                Console.WriteLine("O arquivo foi gerado com sucesso!");
            }
            catch (Exception e)
            {
                //throw new Exception("Erro ao salvar o arquivo com os 100 maiores capitais!", e);
                Console.WriteLine("Erro ao salvar o arquivo com os 100 maiores capitais!");
            }
        }
    }
}
