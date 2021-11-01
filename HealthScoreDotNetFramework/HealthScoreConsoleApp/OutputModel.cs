using System;
using System.Collections.Generic;

namespace HealthScoreConsoleApp
{
	internal class OutputModel
	{
		public int ProfileID { get; internal set; }

		public int HealthScore { get; internal set; }

		public List<EventoOutputModel> Eventos { get; internal set; }
	}
}
