using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AtosCapital
{
    class Nota
    {
        public string nmArquivo { get; set; }
        public string dtNota { get; set; }
        public float vlMercadoria { get; set; }
        public float vlContabil { get; set; }

        public float vlAtos { get; set; }
        public string descricao { get; set; }

        public static List<Nota> notasAgrupadas(List<Nota> notaNaoAgrupadas)
        {

            List<Nota> notas = new List<Nota>();

            notas = notaNaoAgrupadas.GroupBy(e => new { e.nmArquivo, e.dtNota })
                .Select(s => new Nota

                {
                    dtNota = s.Key.dtNota,
                    nmArquivo = s.Key.nmArquivo,
                    vlMercadoria = s.Sum(a => a.vlMercadoria),
                    vlContabil = s.Sum(a => a.vlContabil),
                }
                ).ToList();

            return notas;

        }

    }
}
