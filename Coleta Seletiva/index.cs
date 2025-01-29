using Coleta_Seletiva.Models;
using static System.Runtime.InteropServices.JavaScript.JSType;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using System.IO;
using OfficeOpenXml;
using System.Xml.Linq;
using System.Windows.Forms;
using Microsoft.VisualBasic.ApplicationServices;
using OfficeOpenXml.Style;
using System.Reflection;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Table;
using OfficeOpenXml.Drawing;
using System.Reflection.Emit;

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

        private Image CarregarImagemDeRecurso(string nomeImagem)
        {
            try
            {
                // Substitua "Coleta_Seletiva" pelo namespace do seu projeto
                string resourceName = $"Coleta_Seletiva.Images.{nomeImagem}";
                using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
                {
                    if (stream != null)
                        return new Bitmap(stream);
                }
                return null;
            }
            catch
            {
                return null;
            }
        }


        private void Archive_Load(object sender, EventArgs e)
        {
            var residuos = new List<Residuo>
            {
                new Residuo
                {
                    Etiqueta = "RESÍDUOS DE PAPEL",
                    Tipo = "Papel",
                    Imagem = CarregarImagemDeRecurso("Papel.png"),
                    Descricao = "DEVEM SER DESCARTADOS CAIXAS DE PAPELÃO, PAPÉIS DE ESCRITÓRIO, JORNAIS E REVISTAS."
                },
                new Residuo
                {
                    Etiqueta = "RESÍDUOS DE PLÁSTICO",
                    Tipo = "Plástico",
                    Imagem = CarregarImagemDeRecurso("Plastico.png"),
                    Descricao = "DEVEM SER DESCARTADOS FITILHOS, GARRAFAS PLÁSTICAS, SACOS PLÁSTICOS, TUBOS DE PVC, PASTAS, DIVISÓRIAS PLÁSTICAS, BATOQUES/TAMPÕES NÃO CONTAMINADOS, POTES DE ALIMENTOS HIGIENIZADOS E COPOS DESCARTÁVEIS HIGIENIZADOS."
                },
                new Residuo
                {
                    Etiqueta = "RESÍDUOS NÃO RECICLÁVEIS",
                    Tipo = "Não Reciclável",
                    Imagem = CarregarImagemDeRecurso("NaoReciclavel.png"),
                    Descricao = "DEVEM SER DESCARTADOS FITA ADESIVA, ETIQUETA ADESIVA, EMBALAGENS ALUMINIZADAS (COMO AS DE BISCOITO, CHOCOLATE E BARRINHAS DE CEREAL), PORCELANAS/CERAMICAS, PAPEL HIGIENICO, ABSORVENTE INTIMO, FIO DENTAL, TUBOS DE PASTA DE DENTE E POTES DE ALIMENTOS NÃO HIGIENIZADOS."
                },
                new Residuo
                {
                    Etiqueta = "RESÍDUOS METÁLICOS",
                    Tipo = "Metal",
                    Imagem = CarregarImagemDeRecurso("Metal.png"),
                    Descricao = "DEVEM SER DESCARTADOS LATAS DE ALUMÍNIO, LATAS DE AÇO, PEÇAS METÁLICAS, ARAMES, PREGOS, PARAFUSOS, CHAPAS METÁLICAS E EMBALAGENS METÁLICAS LIMPAS."
                },
                new Residuo
                {
                    Etiqueta = "RESÍDUOS CONTAMINADOS",
                    Tipo = "Contaminado",
                    Imagem = CarregarImagemDeRecurso("Contaminado.png"),
                    Descricao = "DEVEM SER DESCARTADOS PANOS USADOS, ESTOPAS USADAS, PLÁSTICO E PAPEL COM ÓLEO OU GRAXA, BATOQUES/TAMPÕES CONTAMINADOS, TUBOS DE COLA, PINCÉIS USADOS, MARCADORES INDUSTRIAIS, LATAS DE SPRAY, LATAS DE TINTA E EPI'S."
                },
                new Residuo
                {
                    Etiqueta = "RESÍDUOS ORGÂNICOS",
                    Tipo = "Orgânico",
                    Imagem = CarregarImagemDeRecurso("Organico.png"),
                    Descricao = "DEVEM SER DESCARTADOS RESTOS DE ALIMENTOS, CASCAS DE FRUTAS E LEGUMES, BORRA DE CAFÉ, FOLHAS E GALHOS DE PLANTAS."
                },
                new Residuo
                {
                    Etiqueta = "RESÍDUOS DE VIDRO",
                    Tipo = "Vidro",
                    Imagem = CarregarImagemDeRecurso("Vidro.png"),
                    Descricao = "DEVEM SER DESCARTADOS COPOS DE VIDRO, GARRAFAS DE VIDRO, RECIPIENTES DE VIDRO, CACOS DE VIDRO E FRASCOS DE VIDRO."

                },
                new Residuo{
                    Etiqueta = "RESÍDUOS DE MADEIRA",
                    Tipo = "Madeira",
                    Imagem = CarregarImagemDeRecurso("Madeira.png"),
                    Descricao = "DEVEM SER DESCARTADOS PALETES, MÓVEIS QUEBRADOS, SERRAGEM E APARAS DE MADEIRA."
                }
            };

            lista_verificacao.AutoGenerateColumns = false;
            lista_verificacao.DataSource = residuos;

            lista_verificacao.Columns.Clear();

            // Colunas fixas (não editáveis)
            lista_verificacao.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Etiqueta",
                HeaderText = "Etiqueta",
                ReadOnly = true,
                FillWeight = 220
            });

            lista_verificacao.Columns.Add(new DataGridViewImageColumn
            {
                DataPropertyName = "Imagem",
                HeaderText = "Tipo",
                ReadOnly = true,
                ImageLayout = DataGridViewImageCellLayout.Zoom
            });

            lista_verificacao.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Descricao",
                HeaderText = "Descrição",
                ReadOnly = true,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    WrapMode = DataGridViewTriState.True,
                },
                FillWeight = 800
            });

            // Colunas editáveis (Identificação)
            lista_verificacao.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "IdentificacaoCf",
                Name = "IdentificacaoCf",
                HeaderText = "Identificação Conforme",
                ReadOnly = false,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    Alignment = DataGridViewContentAlignment.MiddleCenter
                }
            });

            lista_verificacao.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "IdentificacaoNcf",
                Name = "IdentificacaoNcf",
                HeaderText = "Identificação Não Conforme",
                ReadOnly = false,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    Alignment = DataGridViewContentAlignment.MiddleCenter // Centraliza o conteúdo
                }
            });

            // Colunas editáveis (Separação)
            lista_verificacao.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "SeparacaoCf",
                Name = "SeparacaoCf",
                HeaderText = "Separação Conforme",
                ReadOnly = false,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    Alignment = DataGridViewContentAlignment.MiddleCenter // Centraliza o conteúdo
                }
            });

            lista_verificacao.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "SeparacaoNcf",
                Name = "SeparacaoNcf",
                HeaderText = "Separação Não Conforme",
                ReadOnly = false,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    Alignment = DataGridViewContentAlignment.MiddleCenter // Centraliza o conteúdo
                }
            });

            lista_verificacao.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            lista_verificacao.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            foreach (DataGridViewColumn column in lista_verificacao.Columns)
            {
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                column.DefaultCellStyle.ForeColor = System.Drawing.Color.Black;
            };
        }

        private void GerarNome(List<Residuo> residuos, string caminho)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                FileInfo arquivoExcel = new FileInfo(caminho);
                ExcelWorksheet ws;
                using (var pkg = new ExcelPackage(arquivoExcel))
                {
                    if (pkg.Workbook.Worksheets.Any(s => s.Name == "Respostas"))
                    {
                        ws = pkg.Workbook.Worksheets["Respostas"];
                    }
                    else
                    {
                        ws = pkg.Workbook.Worksheets.Add("Respostas");
                        ws.Cells["A1"].Value = "Data";
                        ws.Cells["B1"].Value = "Nome";
                        ws.Cells["C1"].Value = "Turno";
                        ws.Cells["D1"].Value = "Módulo";
                    }

                    int lastRow = ws.Dimension?.End.Row ?? 1;
                    int linhaAtual = lastRow + 1;

                    foreach (var residuo in residuos)
                    {
                        ws.Cells[$"A{linhaAtual}"].Value = DateTime.Now;
                        ws.Cells[$"A{linhaAtual}"].Style.Numberformat.Format = "dd/MM/yyyy HH:mm:ss";
                        ws.Cells[$"B{linhaAtual}"].Value = txt_nome.Text;
                        ws.Cells[$"C{linhaAtual}"].Value = cb_turno.Text;
                        ws.Cells[$"D{linhaAtual}"].Value = cb_modulo.Text;
                        linhaAtual++;
                    }

                    // Obter a tabela existente
                    var table = ws.Tables.FirstOrDefault(t => t.Name == "TabelaRespostas");
                    if (table != null)
                    {
                        // Excluir a tabela antiga
                        ws.Tables.Delete(table);

                        // Criar uma nova tabela com o intervalo expandido
                        var tableRange = ws.Cells[$"A1:D{linhaAtual - 1}"];
                        table = ws.Tables.Add(tableRange, "TabelaRespostas");
                        table.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;
                    }
                    else
                    {
                        // Criar nova tabela se não houver
                        var tableRange = ws.Cells[$"A1:D{linhaAtual - 1}"];
                        table = ws.Tables.Add(tableRange, "TabelaRespostas");
                        table.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;
                    }

                    ws.Cells[ws.Dimension.Address].AutoFitColumns();
                    pkg.Save();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao gerar planilha de respostas: {ex.Message}");
            }
        }

        private void GerarExcel(List<Residuo> residuos, string caminhoExcel)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                FileInfo arquivoExcel = new FileInfo(caminhoExcel);
                ExcelWorksheet ws;
                ExcelWorksheet wsGrafico;

                using (var pkg = new ExcelPackage(arquivoExcel))
                {
                    // Worksheet Resíduos
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
                        ws.Cells["O1"].Value = "Meta";
                    }

                    int linhaAtual = ws.Dimension?.End.Row + 1 ?? 2;

                    foreach (var residuo in residuos)
                    {
                        // Calcular totais e porcentagens
                        int totalIdent = residuo.IdentificacaoCf + residuo.IdentificacaoNcf;
                        double porcentIdent = totalIdent > 0 ?
                            Math.Round((double)residuo.IdentificacaoCf / totalIdent * 100, 2) : 0;

                        int totalSep = residuo.SeparacaoCf + residuo.SeparacaoNcf;
                        double porcentSep = totalSep > 0 ?
                            Math.Round((double)residuo.SeparacaoCf / totalSep * 100, 2) : 0;

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
                        ws.Cells[$"J{linhaAtual}"].Value = porcentIdent / 100;
                        ws.Cells[$"K{linhaAtual}"].Value = totalSep;
                        ws.Cells[$"L{linhaAtual}"].Value = residuo.SeparacaoCf;
                        ws.Cells[$"M{linhaAtual}"].Value = residuo.SeparacaoNcf;
                        ws.Cells[$"N{linhaAtual}"].Value = porcentSep / 100;
                        ws.Cells[$"O{linhaAtual}"].Value = 0.8; // Meta de 80%

                        ws.Cells[$"J{linhaAtual}"].Style.Numberformat.Format = "0.00%";
                        ws.Cells[$"N{linhaAtual}"].Style.Numberformat.Format = "0.00%";
                        ws.Cells[$"O{linhaAtual}"].Style.Numberformat.Format = "0.00%";

                        linhaAtual++;
                    }

                    // Obter a tabela existente
                    var table = ws.Tables.FirstOrDefault(t => t.Name == "TabelaResiduos");
                    if (table != null)
                    {
                        // Excluir a tabela antiga
                        ws.Tables.Delete(table);

                        // Criar uma nova tabela com o intervalo expandido
                        var tableRange = ws.Cells[$"A1:O{linhaAtual - 1}"];
                        table = ws.Tables.Add(tableRange, "TabelaResiduos");
                        table.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;
                    }
                    else
                    {
                        // Criar nova tabela se não houver
                        var tableRange = ws.Cells[$"A1:O{linhaAtual - 1}"];
                        table = ws.Tables.Add(tableRange, "TabelaResiduos");
                        table.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;
                    }

                    ws.Cells[ws.Dimension.Address].AutoFitColumns();

                    // Agrupar e somar os resíduos por tipo
                    var residuosAgrupados = residuos.GroupBy(r => r.Tipo).Select(g => new
                    {
                        Tipo = g.Key,
                        TotalIdent = g.Sum(r => r.IdentificacaoCf + r.IdentificacaoNcf),
                        TotalSep = g.Sum(r => r.SeparacaoCf + r.SeparacaoNcf)
                    }).ToList();

                    // Create Chart worksheet
                    wsGrafico = pkg.Workbook.Worksheets.Any(s => s.Name == "Gráfico") ?
                        pkg.Workbook.Worksheets["Gráfico"] :
                        pkg.Workbook.Worksheets.Add("Gráfico");

                    // Clear existing charts
                    wsGrafico.Drawings.Clear();

                    // Adicionar dados agrupados ao gráfico
                    linhaAtual = 1;
                    wsGrafico.Cells[$"A{linhaAtual}"].Value = "Tipo Resíduo";
                    wsGrafico.Cells[$"B{linhaAtual}"].Value = "Total Ident";
                    wsGrafico.Cells[$"C{linhaAtual}"].Value = "Total Sep";
                    linhaAtual++;

                    foreach (var grupo in residuosAgrupados)
                    {
                        wsGrafico.Cells[$"A{linhaAtual}"].Value = grupo.Tipo;
                        wsGrafico.Cells[$"B{linhaAtual}"].Value = grupo.TotalIdent;
                        wsGrafico.Cells[$"C{linhaAtual}"].Value = grupo.TotalSep;
                        linhaAtual++;
                    }

                    // Create chart
                    var chartName = "Desempenho";
                    int chartIndex = 1;
                    while (wsGrafico.Drawings.Any(d => d.Name == chartName))
                    {
                        chartName = $"Desempenho{chartIndex}";
                        chartIndex++;
                    }
                    var chart = wsGrafico.Drawings.AddChart(chartName, eChartType.ColumnClustered);

                    // Configure chart series
                    var totalIdentSeries = chart.Series.Add(wsGrafico.Cells[$"B2:B{linhaAtual - 1}"],
                        wsGrafico.Cells[$"A2:A{linhaAtual - 1}"]);
                    totalIdentSeries.Header = "Total Ident";

                    var totalSepSeries = chart.Series.Add(wsGrafico.Cells[$"C2:C{linhaAtual - 1}"],
                        wsGrafico.Cells[$"A2:A{linhaAtual - 1}"]);
                    totalSepSeries.Header = "Total Sep";

                    // Configure chart
                    chart.SetPosition(1, 0, 1, 0);
                    chart.SetSize(800, 400);
                    chart.Title.Text = "Desempenho por Tipo de Resíduo";
                    chart.XAxis.Title.Text = "Tipo de Resíduo";
                    chart.YAxis.Title.Text = "Quantidade";

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
                if (row.IsNewRow) continue;

                int identificacaoCf = ObterValorInteiro(row.Cells["IdentificacaoCf"].Value);
                int identificacaoNcf = ObterValorInteiro(row.Cells["IdentificacaoNcf"].Value);
                int separacaoCf = ObterValorInteiro(row.Cells["SeparacaoCf"].Value);
                int separacaoNcf = ObterValorInteiro(row.Cells["SeparacaoNcf"].Value);

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
                using (SaveFileDialog saveDialog = new SaveFileDialog())
                {
                    saveDialog.Filter = "Arquivos Excel|*.xlsx";
                    saveDialog.Title = "Salvar arquivo Excel";
                    saveDialog.FileName = "Coleta_Seletiva.xlsx";

                    if (saveDialog.ShowDialog() == DialogResult.OK)
                    {
                        var residuos = (List<Residuo>)lista_verificacao.DataSource;
                        GerarExcel(residuos, saveDialog.FileName);
                        GerarNome(residuos, saveDialog.FileName);

                        // Limpar o DataGrid após salvar
                        lista_verificacao.DataSource = null;
                        lista_verificacao.Rows.Clear();

                        txt_nome.Text = string.Empty;
                        txt_data.Text = string.Empty;
                        cb_turno.SelectedIndex = -1;
                        cb_modulo.SelectedIndex = -1;
                        cb_ilha.SelectedIndex = -1;
                        cb_equipe.SelectedIndex = -1;
                        cb_avaliador.SelectedIndex = -1;

                        // Recarregar dados iniciais se necessário
                        Archive_Load(this, EventArgs.Empty);
                    }
                }
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
        private void LimparDataGrid()
        {
            lista_verificacao.DataSource = null;
            lista_verificacao.Rows.Clear();
            lista_verificacao.Columns.Clear();
            lista_verificacao.Refresh();
        }
    }
}