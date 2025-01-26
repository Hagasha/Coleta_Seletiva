using Coleta_Seletiva.Models;
using static QuestPDF.Helpers.Colors;
using static System.Runtime.InteropServices.JavaScript.JSType;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using System.IO;
using OfficeOpenXml;
using System.Xml.Linq;
using System.Windows.Forms;
using Microsoft.VisualBasic.ApplicationServices;
using QuestPDF;
using OfficeOpenXml.Style;

namespace Coleta_Seletiva
{
    public partial class Archive : Form
    {
        public Archive()
        {
            InitializeComponent();
            lista_verificacao.CellValidating += lista_verificacao_CellValidating;
            lista_verificacao.RowPostPaint += lista_verificacao_RowPostPaint;
        }

        private void Archive_Load(object sender, EventArgs e)
        {
            var residuos = new List<Residuo>
            {
                new Residuo
                {
                    Etiqueta = "RESÍDUOS DE PAPEL",
                    Tipo = "Papel",
                    Descricao = "DEVEM SER DESCARTADOS CAIXAS DE PAPELÃO, PAPÉIS DE ESCRITÓRIO, JORNAIS E REVISTAS."
                },
                new Residuo
                {
                    Etiqueta = "RESÍDUOS DE PLÁSTICO",
                    Tipo = "Plástico",
                    Descricao = "DEVEM SER DESCARTADOS FITILHOS, GARRAFAS PLÁSTICAS, SACOS PLÁSTICOS, TUBOS DE PVC, PASTAS, DIVISÓRIAS PLÁSTICAS, BATOQUES/TAMPÕES NÃO CONTAMINADOS, POTES DE ALIMENTOS HIGIENIZADOS E COPOS DESCARTÁVEIS HIGIENIZADOS."
                },
                new Residuo
                {
                    Etiqueta = "RESÍDUOS NÃO RECICLÁVEIS",
                    Tipo = "Não Reciclável",
                    Descricao = "DEVEM SER DESCARTADOS FITA ADESIVA, ETIQUETA ADESIVA, EMBALAGENS ALUMINIZADAS (COMO AS DE BISCOITO, CHOCOLATE E BARRINHAS DE CEREAL), PORCELANAS/CERAMICAS, PAPEL HIGIENICO, ABSORVENTE INTIMO, FIO DENTAL, TUBOS DE PASTA DE DENTE E POTES DE ALIMENTOS NÃO HIGIENIZADOS."
                },
                new Residuo
                {
                    Etiqueta = "RESÍDUOS METÁLICOS",
                    Tipo = "Metal",
                    Descricao = "DEVEM SER DESCARTADOS LATAS DE ALUMÍNIO, LATAS DE AÇO, PEÇAS METÁLICAS, ARAMES, PREGOS, PARAFUSOS, CHAPAS METÁLICAS E EMBALAGENS METÁLICAS LIMPAS."
                },
                new Residuo
                {
                    Etiqueta = "RESÍDUOS CONTAMINADOS",
                    Tipo = "Contaminado",
                    Descricao = "DEVEM SER DESCARTADOS PANOS USADOS, ESTOPAS USADAS, PLÁSTICO E PAPEL COM ÓLEO OU GRAXA, BATOQUES/TAMPÕES CONTAMINADOS, TUBOS DE COLA, PINCÉIS USADOS, MARCADORES INDUSTRIAIS, LATAS DE SPRAY, LATAS DE TINTA E EPI'S."
                },
                new Residuo
                {
                    Etiqueta = "RESÍDUOS ORGÂNICOS",
                    Tipo = "Orgânico",
                    Descricao = "DEVEM SER DESCARTADOS RESTOS DE ALIMENTOS, CASCAS DE FRUTAS E LEGUMES, BORRA DE CAFÉ, FOLHAS E GALHOS DE PLANTAS."
                },
                new Residuo
                {
                    Etiqueta = "RESÍDUOS DE VIDRO",
                    Tipo = "Vidro",
                    Descricao = "DEVEM SER DESCARTADOS COPOS DE VIDRO, GARRAFAS DE VIDRO, RECIPIENTES DE VIDRO, CACOS DE VIDRO E FRASCOS DE VIDRO."

                },
                new Residuo{
                    Etiqueta = "RESÍDUOS DE MADEIRA",
                    Tipo = "Madeira",
                    Descricao = "DEVEM SER DESCARTADOS PALETES, MÓVEIS QUEBRADOS, SERRAGEM E APARAS DE MADEIRA."
                }
            };

            lista_verificacao.AutoGenerateColumns = false; // Importante!
            lista_verificacao.DataSource = residuos;

            // Adicione as colunas manualmente
            lista_verificacao.Columns.Clear();

            // Colunas fixas (não editáveis)
            lista_verificacao.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Etiqueta",
                HeaderText = "Etiqueta",
                ReadOnly = true
            });

            lista_verificacao.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Tipo",
                HeaderText = "Tipo",
                ReadOnly = true
            });

            lista_verificacao.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Descricao",
                HeaderText = "Descrição",
                ReadOnly = true
            });

            // Colunas editáveis (Identificação)
            lista_verificacao.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "IdentificacaoCf", // Ligado à propriedade do modelo
                Name = "IdentificacaoCf", // Nome da coluna no DataGridView
                HeaderText = "Identificação Conforme",
                ReadOnly = false
            });

            lista_verificacao.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "IdentificacaoNcf",
                Name = "IdentificacaoNcf",
                HeaderText = "Identificação Não Conforme",
                ReadOnly = false
            });

            // Colunas editáveis (Separação)
            lista_verificacao.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "SeparacaoCf",
                Name = "SeparacaoCf",
                HeaderText = "Separação Conforme",
                ReadOnly = false
            });

            lista_verificacao.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "SeparacaoNcf",
                Name = "SeparacaoNcf",
                HeaderText = "Separação Não Conforme",
                ReadOnly = false
            });
            foreach (DataGridViewColumn column in lista_verificacao.Columns)
            {
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                column.DefaultCellStyle.ForeColor = System.Drawing.Color.Black;
            }
        }

        private void GerarPDF(List<Residuo> residuos)
        {
           
        }

        private void GerarExcel(List<Residuo> residuos)
        {
            try
            {
                var caminhoExcel = @"C:\Users\Rayan\Downloads\Coleta_Seletiva.xlsx";
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                FileInfo arquivoExcel = new FileInfo(caminhoExcel);
                ExcelWorksheet ws;

                using (var pkg = new ExcelPackage(arquivoExcel))
                {
                    // Verifica se a planilha já existe
                    if (pkg.Workbook.Worksheets.Any(s => s.Name == "Resíduos"))
                    {
                        ws = pkg.Workbook.Worksheets["Resíduos"];
                    }
                    else
                    {
                        ws = pkg.Workbook.Worksheets.Add("Resíduos");
                        // Cabeçalho inicial
                        ws.Cells["A1"].Value = "Data";
                        ws.Cells["B1"].Value = "Avaliador";
                        ws.Cells["C1"].Value = "Turno";
                        ws.Cells["D1"].Value = "Ilha";
                        ws.Cells["E1"].Value = "Equipe";
                        ws.Cells["F1"].Value = "Tipo Resíduo";
                        ws.Cells["G1"].Value = "Total (Ident)";
                        ws.Cells["H1"].Value = "Qtd Conforme (Ident)";
                        ws.Cells["I1"].Value = "Qtd Não Conforme (Ident)";
                        ws.Cells["J1"].Value = "% Identificação";
                        ws.Cells["K1"].Value = "Total (Sep)";
                        ws.Cells["L1"].Value = "Qtd Conforme (Sep)";
                        ws.Cells["M1"].Value = "Qtd Não Conforme (Sep)";
                        ws.Cells["N1"].Value = "% Separação";

                        // Formatar cabeçalho
                        using (var range = ws.Cells["A1:N1"])
                        {
                            range.Style.Font.Bold = true;
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            range.Style.Fill.BackgroundColor.SetColor(Color.Gray);
                        }
                    }

                    // Encontrar última linha preenchida
                    int linhaAtual = ws.Dimension?.End.Row + 1 ?? 2;

                    foreach (var residuo in residuos)
                    {
                        // Calcular totais e porcentagens
                        int totalIdent = residuo.IdentificacaoCf + residuo.IdentificacaoNcf;
                        double porcentIdent = totalIdent > 0 ?
                            Math.Round((double)residuo.IdentificacaoCf / totalIdent, 4) : 0;

                        int totalSep = residuo.SeparacaoCf + residuo.SeparacaoNcf;
                        double porcentSep = totalSep > 0 ?
                            Math.Round((double)residuo.SeparacaoCf / totalSep, 4) : 0; 


                        // Preencher dados
                        ws.Cells[$"A{linhaAtual}"].Value = txt_data.Text;
                        ws.Cells[$"B{linhaAtual}"].Value = cb_avaliador.Text;
                        ws.Cells[$"C{linhaAtual}"].Value = cb_turno.Text;
                        ws.Cells[$"D{linhaAtual}"].Value = cb_ilha.Text;
                        ws.Cells[$"E{linhaAtual}"].Value = cb_equipe.Text;
                        ws.Cells[$"F{linhaAtual}"].Value = residuo.Tipo;
                        ws.Cells[$"G{linhaAtual}"].Value = totalIdent;
                        ws.Cells[$"H{linhaAtual}"].Value = residuo.IdentificacaoCf;
                        ws.Cells[$"I{linhaAtual}"].Value = residuo.IdentificacaoNcf;
                        ws.Cells[$"J{linhaAtual}"].Value = porcentIdent;
                        ws.Cells[$"K{linhaAtual}"].Value = totalSep;
                        ws.Cells[$"L{linhaAtual}"].Value = residuo.SeparacaoCf;
                        ws.Cells[$"M{linhaAtual}"].Value = residuo.SeparacaoNcf;
                        ws.Cells[$"N{linhaAtual}"].Value = porcentSep;

                        // Formatar porcentagens
                        ws.Cells[$"J{linhaAtual}"].Style.Numberformat.Format = "0.00%";
                        ws.Cells[$"N{linhaAtual}"].Style.Numberformat.Format = "0.00%";

                        linhaAtual++;
                    }

                    // Autoajustar colunas
                    ws.Cells[ws.Dimension.Address].AutoFitColumns();
                    pkg.Save();
                }

                MessageBox.Show("Dados adicionados com sucesso!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao gerar Excel: {ex.Message}");
            }
        }

        private void btn_salvar_Click(object sender, EventArgs e)
        {
            bool dadosValidos = true;

            foreach (DataGridViewRow row in lista_verificacao.Rows)
            {
                if (row.IsNewRow) continue; // Ignorar linha em branco

                int identificacaoCf = Convert.ToInt32(row.Cells["IdentificacaoCf"].Value ?? 0);
                int identificacaoNcf = Convert.ToInt32(row.Cells["IdentificacaoNcf"].Value ?? 0);
                int separacaoCf = Convert.ToInt32(row.Cells["SeparacaoCf"].Value ?? 0);
                int separacaoNcf = Convert.ToInt32(row.Cells["SeparacaoNcf"].Value ?? 0);

                int totalIdentificacao = identificacaoCf + identificacaoNcf;
                int totalSeparacao = separacaoCf + separacaoNcf;

                if (totalIdentificacao != totalSeparacao)
                {
                    dadosValidos = false;
                    row.ErrorText = "Soma de Identificação ≠ Separação!";
                }
            }

            if (dadosValidos)
            {
                var residuos = (List<Residuo>)lista_verificacao.DataSource;
                GerarExcel(residuos);
                MessageBox.Show("Arquivo gerado com sucesso!");
            }
            else
            {
                MessageBox.Show("Corrija os dados destacados antes de salvar!");
            }

        }

        private void lista_verificacao_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (lista_verificacao.CurrentCell.ColumnIndex >= 3) // Colunas 3 a 6 são numéricas
            {
                TextBox textBox = e.Control as TextBox;
                if (textBox != null)
                {
                    textBox.KeyPress += (s, ev) =>
                    {
                        if (!char.IsControl(ev.KeyChar) && !char.IsDigit(ev.KeyChar))
                        {
                            ev.Handled = true;
                        }
                    };
                }
            }
        }
        private void lista_verificacao_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            // Verifica se a célula editada está nas colunas de quantidade (3 a 6)
            if (e.ColumnIndex >= 3 && e.ColumnIndex <= 6)
            {
                DataGridViewRow row = lista_verificacao.Rows[e.RowIndex];

                // Obter valores (trate valores nulos)
                int identificacaoCf = ObterValorInteiro(row.Cells["IdentificacaoCf"].Value);
                int identificacaoNcf = ObterValorInteiro(row.Cells["IdentificacaoNcf"].Value);
                int separacaoCf = ObterValorInteiro(row.Cells["SeparacaoCf"].Value);
                int separacaoNcf = ObterValorInteiro(row.Cells["SeparacaoNcf"].Value);

                // Calcular totais
                int totalIdentificacao = identificacaoCf + identificacaoNcf;
                int totalSeparacao = separacaoCf + separacaoNcf;

                // Validar
                if (totalIdentificacao != totalSeparacao)
                {
                    row.ErrorText = "Erro: Total de Identificação ≠ Separação!";
                }
                else
                {
                    row.ErrorText = string.Empty; // Limpa o erro
                }
            }
        }

        private int ObterValorInteiro(object valor)
        {
            if (valor == null || string.IsNullOrEmpty(valor.ToString()))
                return 0;

            int resultado;
            return int.TryParse(valor.ToString(), out resultado) ? resultado : 0;
        }


        private void lista_verificacao_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            DataGridViewRow row = lista_verificacao.Rows[e.RowIndex];

            // Destacar linha se houver erro
            if (!string.IsNullOrEmpty(row.ErrorText))
            {
                row.DefaultCellStyle.BackColor = Color.LightCoral;
            }
            else
            {
                row.DefaultCellStyle.BackColor = Color.White;
            }
        }
    }
}
