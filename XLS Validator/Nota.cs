using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AtosCapital
{
    public class Nota
    {
        public string dsFantasia { get; set; }
        public string nmCNPJ { get; set; }
        public string nmArquivo { get; set; }
        public string dtNota { get; set; }
        public float vlMercadoria { get; set; }
        public float vlContabil { get; set; }
        public string cdChave { get; set; }
        public float vlAtos { get; set; }
        public string descricao { get; set; }

        public static List<Nota> notasAgrupadas(List<Nota> notaNaoAgrupadas)
        {

            List<Nota> notas = new List<Nota>();

            notas = notaNaoAgrupadas.GroupBy(e => new { e.nmCNPJ, e.dtNota, e.dsFantasia, e.nmArquivo})
                .Select(s => new Nota
                {
                    dsFantasia = s.Key.dsFantasia,
                    nmCNPJ = s.Key.nmCNPJ,
                    dtNota = s.Key.dtNota,
                    nmArquivo = s.Key.nmArquivo,
                    vlMercadoria = s.Sum(a => a.vlMercadoria),
                    vlContabil = s.Sum(a => a.vlContabil),
                }).ToList();

            return notas;

        }

        public static List<Nota> notasAgrupadasPorChave(List<Nota> notasNaoAgrupadas, DateTime data)
        {
            List<Nota> notas = new List<Nota>();
            notas = notasNaoAgrupadas.FindAll(n => n.dtNota == data.ToString("yyyyMMdd"));
            notas = notas.GroupBy(g => g.cdChave)
                .Select(s => new Nota
                {
                    cdChave = s.Key,
                    nmArquivo = s.First().nmArquivo,
                    dsFantasia = s.First().dsFantasia,
                    dtNota = s.First().dtNota,
                    nmCNPJ = s.First().nmCNPJ,
                    vlMercadoria = s.Sum(a => a.vlMercadoria),
                    vlContabil = s.Sum(a => a.vlContabil),
                }).ToList();

            return notas;
        }

    }
}
