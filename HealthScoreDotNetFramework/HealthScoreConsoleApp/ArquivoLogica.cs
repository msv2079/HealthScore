using IronXL;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace HealthScoreConsoleApp
{
	public class ArquivoLogica
	{
		public Tuple<bool, string> Validar(string arquivo)
		{
			if (string.IsNullOrWhiteSpace(arquivo))
				return new Tuple<bool, string>(false, "Informe o caminho completo da planilha no arquivo de configurações App.config");

			if (!File.Exists(arquivo))
				return new Tuple<bool, string>(false, "Planilha não localizada! Informe o caminho completo da planilha no arquivo de configurações App.config");

			if (!".xlsx".Equals(new FileInfo(arquivo).Extension, StringComparison.CurrentCultureIgnoreCase))
				return new Tuple<bool, string>(false, "Só sei trabalhar com arquivos .xlsx");


			var planilhaValida = false;

			var workbook = WorkBook.Load(arquivo);

			var workSheets = workbook.WorkSheets;

			if (workSheets.Count == 3 && workSheets.Any(x => x.Name.Contains("Settings")) && workSheets.Any(x => x.Name.Contains("Events")) && workSheets.Any(x => x.Name.Contains("Outputs")))
				planilhaValida = true;

			workbook.Close();

			if (!planilhaValida)
				throw new Exception("Arquivo não possui as planilhas corretas!");

			return new Tuple<bool, string>(true, "");
		}

		public void Processar(string arquivo)
		{
			var workbook = WorkBook.Load(arquivo);

			var planilhaSettings = workbook.WorkSheets.First(x => x.Name == "Settings");

			var strPerfis = Convert.ToString(planilhaSettings["B2"].Value);
			var strBaseHealthScore = Convert.ToString(planilhaSettings["B3"].Value);

			if (!int.TryParse(strPerfis, out var intPerfis))
				throw new Exception("Número de Perfis Inválido");

			if (!int.TryParse(strBaseHealthScore, out var intBaseHealthScore))
				throw new Exception("Número de Perfis Inválido");

			var listaEventos = RetornarEvents(workbook);
			var listaOutput = new List<OutputModel>();

			var listaPerfis = Enumerable.Range(1, intPerfis).ToList();
			var healthScoreFinal = 0;

			while (listaPerfis.Count > 0)
			{
				//Tive que adicionar esse sleep pois sem ele todos os números aleatórios eram gerados identicos
				Thread.Sleep(100);

				var proximoPerfil = false;

				var perfilId = listaPerfis.First();

				var numeroAleatorio = new Random().Next(0, 10);

				var healthScoreInicial = intBaseHealthScore + numeroAleatorio;

				/*
					Aqui foi utilizado a quantidade de eventos da lista, pois assim ficará dinamico
					Se tiver 50 eventos não precisará mudar o maior valor de Random, 
					nem correrá o risco de sortear o número 5 e ter apenas 3 eventos na planilha o que retornaria null
				*/

				var idAleatorio = new Random().Next(1, listaEventos.Count + 1);
				EventoModel evento = listaEventos.FirstOrDefault(x => x.Id == idAleatorio);

				if ("Death".Equals(evento.NameEvent, StringComparison.CurrentCultureIgnoreCase))
					proximoPerfil = true;

				var heathCalculo = healthScoreFinal == 0 ? healthScoreInicial : healthScoreFinal;

				healthScoreFinal = heathCalculo + evento.HealthScoreDiscount;

				if (healthScoreFinal < 1)
					proximoPerfil = true;

				var perfil = listaOutput.FirstOrDefault(x => x.ProfileID == perfilId);

				if (perfil == null)
				{
					listaOutput.Add(new OutputModel()
					{
						ProfileID = perfilId,
						Eventos = new List<EventoOutputModel>(),
						HealthScore = 0
					});

					perfil = listaOutput.FirstOrDefault(x => x.ProfileID == perfilId);

					foreach (var item in listaEventos)
					{
						perfil.Eventos.Add(new EventoOutputModel
						{
							NomeEvento = item.NameEvent,
							Ocorrencias = 0
						});
					}
				}

				var eventoPerfil = perfil.Eventos.First(x => x.NomeEvento == evento.NameEvent);
				eventoPerfil.Ocorrencias = eventoPerfil.Ocorrencias + 1;
				perfil.HealthScore = healthScoreFinal;

				if (proximoPerfil)
				{
					healthScoreFinal = 0;
					listaPerfis.RemoveAt(0);
				}
			}

			GravarOutput(workbook, listaOutput);

			workbook.Save();

			workbook.Close();
		}

		private void GravarOutput(WorkBook workbook, List<OutputModel> listaOutput)
		{
			var outputWorkSheet = workbook.WorkSheets.First(x => x.Name == "Outputs");

			for (int i = 0; i < listaOutput.Count; i++)
			{
				var celulaInicialEvento = 67;

				var item = listaOutput[i];

				outputWorkSheet[$"A{i + 3}"].Value = item.ProfileID;
				outputWorkSheet[$"B{i + 3}"].Value = item.HealthScore;

				foreach (var evento in item.Eventos)
				{
					if (i == 0)
					{
						outputWorkSheet[$"{(char)celulaInicialEvento}2"].Value = evento.NomeEvento;
					}

					outputWorkSheet[$"{(char)celulaInicialEvento}{i + 3}"].Value = evento.Ocorrencias;
					celulaInicialEvento++;
				}
			}
		}

		private List<EventoModel> RetornarEvents(WorkBook workbook)
		{
			var listaEventos = new List<EventoModel>();

			var planilhaEvents = workbook.WorkSheets.First(x => x.Name == "Events");
			var contador = 3;

			while (true)
			{
				var strId = Convert.ToString(planilhaEvents[$"A{contador}"].Value);
				var strName = Convert.ToString(planilhaEvents[$"B{contador}"].Value);
				var strHealthScoreDiscount = Convert.ToString(planilhaEvents[$"C{contador}"].Value);

				if (string.IsNullOrWhiteSpace(strId) || string.IsNullOrWhiteSpace(strName) || string.IsNullOrWhiteSpace(strHealthScoreDiscount))
					break;

				if (!int.TryParse(strId, out var intId) || !int.TryParse(strHealthScoreDiscount, out var intHealthScoreDiscount))
					break;

				listaEventos.Add(new EventoModel()
				{
					Id = intId,
					NameEvent = strName,
					HealthScoreDiscount = intHealthScoreDiscount
				});

				contador++;
			}

			return listaEventos;
		}
	}
}
