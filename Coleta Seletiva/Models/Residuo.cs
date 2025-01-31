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
        internal double meta;

        public string Etiqueta { get; set; }
        public string Tipo { get; set; }
        public Image Imagem { get; set; }
        public string Equipe { get; set; }
        public string Descricao { get; set; }
        public int IdentificacaoCf { get; set; }
        public int IdentificacaoNcf { get; set; }
        public int SeparacaoCf { get; set; }
        public int SeparacaoNcf { get; set; }
        public int Ano { get; set; }
        public int Mes { get; set; }
        public int Semana { get; set; }
        public double Eficiencia { get; set; }

    }
}
