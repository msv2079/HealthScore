using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HealthScoreConsoleApp
{
	class Program
	{
		static void Main(string[] args)
		{
			Console.WriteLine("Processando...");

			try
			{
				ProcessarArquivo();
				Console.Clear();
				Console.WriteLine("Arquivo Processado com sucesso!");
			}
			catch (Exception ex)
			{
				Console.Clear();
				Console.WriteLine("Hum... parece que deu erro, veja se a menagem abaixo ajuda em algo");
				Console.WriteLine(ex.Message);
			}

			Console.ReadKey();
		}

		private static void ProcessarArquivo()
		{
			var nomeArquivo = RetornarArquivo();

			var logica = new ArquivoLogica();
			var validarArquivo = logica.Validar(nomeArquivo);

			if (validarArquivo.Item1)
				logica.Processar(nomeArquivo);
			else
				throw new Exception(validarArquivo.Item2);
		}

		private static string RetornarArquivo()
		{
			var arquivo = ConfigurationManager.AppSettings["CaminhoCompletoArquivo"];

			return arquivo;
		}
	}
}
