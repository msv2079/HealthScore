using HealthScoreConsoleApp;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace HealthScoreTests
{
	[TestClass]
	public class ArquivoTests
	{
		[TestMethod]
		public void TestarArquivoValido()
		{
			var arquivo = new ArquivoLogica();

			var resultado = arquivo.Validar(@"D:\BKP\VER\TesteTecnico\Model.xlsx");

			Assert.IsTrue(resultado.Item1);
		}

		[TestMethod]
		public void TestarArquivoInexistente()
		{
			var arquivo = new ArquivoLogica();

			var resultado = arquivo.Validar(@"D:\BKP\VER\TesteTecnico\Modelo.xlsx");

			Assert.IsFalse(resultado.Item1);
		}

		[TestMethod]
		public void TestarExtensaoInvalida()
		{
			var arquivo = new ArquivoLogica();

			var resultado = arquivo.Validar(@"D:\BKP\VER\TesteTecnico\Model.xls");

			Assert.IsFalse(resultado.Item1);
		}
	}
}
