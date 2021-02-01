using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AtosCapital
{

    class Filiais
    {
        public string ds_fantasia { get; set; }
        public string nu_cnpj { get; set; }

        public List<Filiais> listarFiliais()
        {
            using (BancoDados _dbAtos = new BancoDados(true))
            {
                List<Filiais> listaFiliais = BancoDadosUtils.ConvertDataTableToDataType<Filiais>(_dbAtos.SqlQueryDataTable(@"SELECT * FROM cliente.empresa (NOLOCK)
                                                                                                                             WHERE id_grupo = 776 AND sg_uf = 'SP';")
                                                                                                                                                                    ).ToList();
                foreach(Filiais f in listaFiliais)
                {
                    f.ds_fantasia = RemoveDiacritics(f.ds_fantasia);
                }
                return listaFiliais;
            }

        }

        static string RemoveDiacritics(string text)
        {
            return string.Concat(
                text.Normalize(NormalizationForm.FormD)
                .Where(ch => CharUnicodeInfo.GetUnicodeCategory(ch) !=
                                              UnicodeCategory.NonSpacingMark)
              ).Normalize(NormalizationForm.FormC);
        }

    }
}
