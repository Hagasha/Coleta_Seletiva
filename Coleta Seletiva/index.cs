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
using System.Drawing.Imaging;
using System.Globalization;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Drawing.Slicer;

namespace Coleta_Seletiva
{
    public partial class Archive : Form
    {
        private const string CAMINHO_BASE = @"\\VWTBRESHFSCO\Modulos\Maxion\POWER_BI\Arquivos Coleta Seletiva";
        private const string CAMINHO_SCREENSHOTS = @"\\VWTBRESHFSCO\Modulos\Maxion\POWER_BI\Arquivos Coleta Seletiva\Screenshots";
        private const string NOME_ARQUIVO_EXCEL = "Avaliacao_Coleta_Seletiva.xlsx";

        public Archive()
        {
            InitializeComponent();
            lista_verificacao.CellValidating += lista_verificacao_CellValidating;
            lista_verificacao.RowPostPaint += lista_verificacao_RowPostPaint;
        }

        private void ValidarFormatoData(TextBox textBox)
        {
            if (!string.IsNullOrEmpty(textBox.Text))
            {
                DateTime data;
                if (!DateTime.TryParseExact(textBox.Text, "dd/MM/yyyy",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out data))
                {
                    MessageBox.Show("Por favor, insira a data no formato dd/mm/yyyy");
                    textBox.Text = "";
                }
            }
        }

        private void txt_data_Leave(object sender, EventArgs e)
        {
            ValidarFormatoData(txt_data);
        }

        private void CaptureFormScreenshot()
        {
            try
            {
                using (Bitmap bitmap = new Bitmap(this.Width, this.Height))
                {
                    this.DrawToBitmap(bitmap, new Rectangle(0, 0, this.Width, this.Height));

                    // Obter o conteúdo dos controles
                    string ilha = cb_ilha.Text;
                    string nome = txt_nome.Text;
                    string data = txt_data.Text.Replace("/", "");
                    string turno = cb_turno.Text + "Turno";

                    // Criar diretório se não existir
                    if (!Directory.Exists(CAMINHO_SCREENSHOTS))
                    {
                        Directory.CreateDirectory(CAMINHO_SCREENSHOTS);
                    }

                    // Definir o caminho completo do arquivo
                    string caminhoCompleto = Path.Combine(CAMINHO_SCREENSHOTS, $"avaliacao_{ilha}_{nome}_{data}_{turno}.png");

                    // Salvar a imagem diretamente no local especificado
                    bitmap.Save(caminhoCompleto, ImageFormat.Png);
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao capturar screenshot: {ex.Message}");
            }
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
            // No Archive_Load, ao carregar os resíduos:
            var listaResiduos = new List<Residuo>
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
            lista_verificacao.DataSource = listaResiduos;

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

            if (!string.IsNullOrEmpty(txt_data.Text))
            {
                DateTime data;
                if (DateTime.TryParseExact(txt_data.Text, "dd/MM/yyyy",
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                    out data))
                {
                    foreach (var residuo in listaResiduos)
                    {
                        residuo.Ano = data.Year;
                        residuo.Mes = data.Month;
                        residuo.Semana = CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(
                            data,
                            CalendarWeekRule.FirstFourDayWeek,
                            DayOfWeek.Monday);
                    }
                }
            }
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
                        // Novo cabeçalho com ordem atualizada
                        ws.Cells["A1"].Value = "Data";
                        ws.Cells["B1"].Value = "Ano";
                        ws.Cells["C1"].Value = "Mês";
                        ws.Cells["D1"].Value = "Semana";
                        ws.Cells["E1"].Value = "Turno";
                        ws.Cells["F1"].Value = "Módulo";
                        ws.Cells["G1"].Value = "Ilha";
                        ws.Cells["H1"].Value = "Equipe";
                        ws.Cells["I1"].Value = "Coletor avaliado";
                        ws.Cells["J1"].Value = "Identificação";
                        ws.Cells["K1"].Value = "Separação";
                        ws.Cells["L1"].Value = "Eficiência";
                        ws.Cells["M1"].Value = "Meta";
                        ws.Cells["N1"].Value = "Avaliador";
                    }

                    int linhaAtual = ws.Dimension?.End.Row + 1 ?? 2;

                    foreach (var residuo in residuos)
                    {
                        // Parsear a data corretamente
                        DateTime data = DateTime.ParseExact(txt_data.Text, "dd/MM/yyyy",
                            CultureInfo.InvariantCulture);

                        // Atualizar os valores de ano, mês e semana no objeto residuo
                        residuo.Ano = data.Year;
                        residuo.Mes = data.Month;
                        residuo.Semana = CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(
                            data,
                            CalendarWeekRule.FirstFourDayWeek,
                            DayOfWeek.Monday);

                        // Cálculo da eficiência
                        double eficiencia = CalcularEficiencia(residuo);
                        double separacao = CalcularSeparacao(residuo);
                        double identificacao = CalcularIdentificacao(residuo);
                        // Preencher dados na nova ordem
                        ws.Cells[$"A{linhaAtual}"].Value = data;
                        ws.Cells[$"B{linhaAtual}"].Value = residuo.Ano;
                        ws.Cells[$"C{linhaAtual}"].Value = residuo.Mes;
                        ws.Cells[$"D{linhaAtual}"].Value = residuo.Semana;
                        ws.Cells[$"E{linhaAtual}"].Value = cb_turno.Text;
                        ws.Cells[$"F{linhaAtual}"].Value = cb_modulo.Text;
                        ws.Cells[$"G{linhaAtual}"].Value = cb_ilha.Text.Replace("Equipe", "");
                        ws.Cells[$"H{linhaAtual}"].Value = cb_equipe.Text;
                        ws.Cells[$"I{linhaAtual}"].Value = residuo.Tipo;
                        ws.Cells[$"J{linhaAtual}"].Value = identificacao;
                        ws.Cells[$"K{linhaAtual}"].Value = separacao;
                        ws.Cells[$"L{linhaAtual}"].Value = eficiencia;
                        ws.Cells[$"M{linhaAtual}"].Value = 0.8; 
                        ws.Cells[$"N{linhaAtual}"].Value = cb_avaliador.Text;

                        ws.Cells[$"J{linhaAtual}"].Style.Numberformat.Format = "0.00%";
                        ws.Cells[$"K{linhaAtual}"].Style.Numberformat.Format = "0.00%";
                        ws.Cells[$"L{linhaAtual}"].Style.Numberformat.Format = "0.00%";
                        ws.Cells[$"M{linhaAtual}"].Style.Numberformat.Format = "0.00%";

                        linhaAtual++;
                    }

                    // Configurar tabela
                    ConfigurarTabela(ws, linhaAtual);

                    // Criar gráfico atualizado
                    CriarGraficoEficiencia(pkg, residuos);

                    pkg.Save();
                }

                MessageBox.Show("Dados adicionados com sucesso!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao gerar Excel: {ex.Message}");
            }
        }

        private void ConfigurarTabela(ExcelWorksheet ws, int linhaAtual)
        {
            // Obter a tabela existente
            var table = ws.Tables.FirstOrDefault(t => t.Name == "TabelaResiduos");
            if (table != null)
            {
                // Excluir a tabela antiga
                ws.Tables.Delete(table);
            }

            // Criar uma nova tabela com o intervalo atualizado
            var tableRange = ws.Cells[$"A1:N{linhaAtual - 1}"];
            table = ws.Tables.Add(tableRange, "TabelaResiduos");
            table.TableStyle = TableStyles.Medium2;

            // Configurar formatação das colunas
            ws.Column(1).Style.Numberformat.Format = "dd/MM/yyyy";
            ws.Column(10).Style.Numberformat.Format = "0.00%";      // Eficiência
            ws.Column(11).Style.Numberformat.Format = "0.00%";      // Meta
            ws.Column(12).Style.Numberformat.Format = "0.00%";      // Eficiência
            ws.Column(13).Style.Numberformat.Format = "0.00%";

            // Auto-ajustar largura das colunas
            ws.Cells[ws.Dimension.Address].AutoFitColumns();

            // Aplicar formatação condicional para eficiência
            var eficienciaRange = ws.Cells[$"L2:L{linhaAtual - 1}"];
            var cfRule = eficienciaRange.ConditionalFormatting.AddTwoColorScale();
            cfRule.LowValue.Color = Color.FromArgb(255, 199, 206);  // Vermelho claro
            cfRule.HighValue.Color = Color.FromArgb(198, 239, 206); // Verde claro
        }
        private double CalcularEficiencia(Residuo residuo)
        {
            int totalIdent = residuo.IdentificacaoCf + residuo.IdentificacaoNcf;
            int totalSep = residuo.SeparacaoCf + residuo.SeparacaoNcf;

            if (totalIdent + totalSep == 0) return 0;

            return (double)(residuo.IdentificacaoCf + residuo.SeparacaoCf) / (totalIdent + totalSep);
        }

        private double CalcularSeparacao(Residuo residuo)
        {
            int totalIdent = residuo.IdentificacaoCf + residuo.IdentificacaoNcf;
            int totalSep = residuo.SeparacaoCf + residuo.SeparacaoNcf;

            if (totalIdent + totalSep == 0) return 0;

            return (double)(residuo.SeparacaoCf) / (totalSep);
        }

        private double CalcularIdentificacao(Residuo residuo)
        {
            int totalIdent = residuo.IdentificacaoCf + residuo.IdentificacaoNcf;
            int totalSep = residuo.SeparacaoCf + residuo.SeparacaoNcf;

            if (totalIdent + totalSep == 0) return 0;

            return (double)(residuo.IdentificacaoCf) / (totalIdent);
        }

        private void CriarGraficoEficiencia(ExcelPackage pkg, List<Residuo> residuos)
        {
            var wsResiduos = pkg.Workbook.Worksheets["Resíduos"];
            var wsGrafico = pkg.Workbook.Worksheets.Any(s => s.Name == "Gráfico") ?
                pkg.Workbook.Worksheets["Gráfico"] :
                pkg.Workbook.Worksheets.Add("Gráfico");


            var existingPivotTable = wsGrafico.PivotTables.FirstOrDefault(pt => pt.Name == "PivotEficiencia");
            if (existingPivotTable != null)
                wsGrafico.PivotTables.Delete(existingPivotTable);

            if (wsResiduos?.Dimension == null)
            {
                MessageBox.Show("Não há dados na planilha de resíduos para gerar o gráfico.");
                return;
            }

            var sourceRange = wsResiduos.Cells[wsResiduos.Dimension.Address];
            var pivotTable = wsGrafico.PivotTables.Add(wsGrafico.Cells["A3"], sourceRange, "PivotEficiencia");

            // Configuração dos campos da pivot (mantido igual)
            var semanaField = pivotTable.RowFields.Add(pivotTable.Fields["Semana"]);
            var eficienciaField = pivotTable.DataFields.Add(pivotTable.Fields["Eficiência"]);
            var metaField = pivotTable.DataFields.Add(pivotTable.Fields["Meta"]);

            semanaField.Name = "Rótulos de Linha";
            eficienciaField.Name = "Média de Eficiência";
            metaField.Name = "Média de Meta";

            eficienciaField.Function = DataFieldFunctions.Average;
            metaField.Function = DataFieldFunctions.Average;
            eficienciaField.Format = "0.00%";
            metaField.Format = "0.00%";

            pivotTable.RowGrandTotals = false;
            pivotTable.ColumnGrandTotals = false;
            semanaField.SubtotalTop = false;
            pivotTable.DataOnRows = false;
            semanaField.Sort = eSortType.Ascending;

            pivotTable.Calculate();

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
                try
                {
                    // Capturar screenshot primeiro
                    CaptureFormScreenshot();

                    // Criar diretório se não existir
                    if (!Directory.Exists(CAMINHO_BASE))
                    {
                        Directory.CreateDirectory(CAMINHO_BASE);
                    }

                    // Caminho completo do arquivo Excel
                    string caminhoExcel = Path.Combine(CAMINHO_BASE, NOME_ARQUIVO_EXCEL);

                    var residuos = (List<Residuo>)lista_verificacao.DataSource;
                    GerarExcel(residuos, caminhoExcel);
                    GerarNome(residuos, caminhoExcel);

                    // Limpar o DataGrid após salvar
                    lista_verificacao.DataSource = null;
                    lista_verificacao.Rows.Clear();

                    // Limpar campos
                    txt_nome.Text = string.Empty;
                    txt_data.Text = string.Empty;
                    cb_turno.SelectedIndex = -1;
                    cb_modulo.SelectedIndex = -1;
                    cb_ilha.SelectedIndex = -1;
                    cb_equipe.SelectedIndex = -1;
                    cb_avaliador.SelectedIndex = -1;

                    // Recarregar dados iniciais
                    Archive_Load(this, EventArgs.Empty);

                    MessageBox.Show("Dados salvos com sucesso!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Erro ao salvar os dados: {ex.Message}");
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