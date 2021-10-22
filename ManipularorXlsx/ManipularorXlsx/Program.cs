using System;

namespace ManipularorXlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            Manipulador manipulador = new Manipulador();
            int apcaoMenuPrincipal = 0;

            Console.Write("\nInforme o caminho do arquivo que deseja exibir os dados:");
            var diretorio = Console.ReadLine();
            while (!manipulador.CarregaDadosXlsx(diretorio))
            {
                Console.Write("\nErro ao carregar arquivo! Verifique o diretório informado:");
                diretorio = Console.ReadLine();
            }

            do
            {
                apcaoMenuPrincipal = Menu.ExibeMenuPrincipal();

                switch (apcaoMenuPrincipal)
                {
                    case Menu.ExibirDadosDoArquivo:
                        manipulador.ExibirDadosDoArquivo();
                        break;
                    case Menu.ExibirCemMaioresCapitais:
                        manipulador.ExibirCemMaioresCapitais();
                        break;
                    case 3:
                        Environment.Exit(0);
                        break;
                    default:
                        Console.WriteLine("\nOpção inválida!");
                        break;
                }

            } while (apcaoMenuPrincipal != 6);
        }
    }
}
