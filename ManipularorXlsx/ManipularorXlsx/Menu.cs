using System;
using System.Collections.Generic;
using System.Text;

namespace ManipularorXlsx
{
    public class Menu
    {
        // Opções possíveis 
        public const int ExibirDadosDoArquivo = 1;
        public const int ExibirCemMaioresCapitais = 2;
        public const int Sair = 3;
        private static Dictionary<int, string> Opcoes { get; set; }

        // Construtor estático do menu
        static Menu()
        {
            Opcoes = new Dictionary<int, string>()
            {
                {ExibirDadosDoArquivo, "EXIBIR DADOS DO ARQUIVO XLSX"},
                {ExibirCemMaioresCapitais, "EXIBIR OS 100 MAIORES CAPITAIS DA B3"},
                {Sair, "SAIR" }
            };
        }

        public static int ExibeMenuPrincipal()
        {

            int escolha;
            do
            {
                Console.WriteLine("\nInforme qual operação deseja realizar...");
                foreach (var opcao in Opcoes)
                {
                    System.Console.WriteLine($"{opcao.Key} - {opcao.Value}");
                }

                Int32.TryParse(Console.ReadLine(), out escolha);
            } while (!Opcoes.ContainsKey(escolha));

            return escolha;
        }
    }
}
