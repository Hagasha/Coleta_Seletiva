using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace Coleta_Seletiva.Models
{
    internal class Residuo
    {
        public string Etiqueta { get; set; }
        public string Tipo { get; set; }
        public string Descricao { get; set; }
        public int IdentificacaoCf { get; set; }
        public int IdentificacaoNcf { get; set; }
        public int SeparacaoCf { get; set; }
        public int SeparacaoNcf { get; set; }

    }
}
