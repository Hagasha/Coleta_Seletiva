namespace Coleta_Seletiva
{
    partial class Archive
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            txt_data = new TextBox();
            lbl_titulo = new Label();
            lbl_dados = new Label();
            Txt_dados = new Label();
            lbl_data = new Label();
            lbl_turno = new Label();
            lbl_modulo = new Label();
            txt_modulo = new TextBox();
            lbl_ilha = new Label();
            lbl_equipe = new Label();
            lbl_avaliador = new Label();
            cb_avaliador = new ComboBox();
            cb_equipe = new ComboBox();
            cb_ilha = new ComboBox();
            label1 = new Label();
            btn_salvar = new Button();
            lista_verificacao = new DataGridView();
            cb_turno = new ComboBox();
            ((System.ComponentModel.ISupportInitialize)lista_verificacao).BeginInit();
            SuspendLayout();
            // 
            // txt_data
            // 
            txt_data.BackColor = Color.DeepSkyBlue;
            txt_data.BorderStyle = BorderStyle.None;
            txt_data.Font = new Font("Segoe UI", 12F);
            txt_data.Location = new Point(67, 145);
            txt_data.Name = "txt_data";
            txt_data.Size = new Size(129, 22);
            txt_data.TabIndex = 0;
            // 
            // lbl_titulo
            // 
            lbl_titulo.BackColor = Color.DodgerBlue;
            lbl_titulo.BorderStyle = BorderStyle.FixedSingle;
            lbl_titulo.Font = new Font("Arial", 35F, FontStyle.Bold);
            lbl_titulo.Location = new Point(-2, 0);
            lbl_titulo.Margin = new Padding(10, 0, 3, 0);
            lbl_titulo.Name = "lbl_titulo";
            lbl_titulo.Size = new Size(1188, 50);
            lbl_titulo.TabIndex = 1;
            lbl_titulo.Text = "AVALIAÇÃO DE EFICIÊNCIA DA COLETA SELETIVA";
            // 
            // lbl_dados
            // 
            lbl_dados.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            lbl_dados.BackColor = Color.DodgerBlue;
            lbl_dados.Font = new Font("Arial", 18F);
            lbl_dados.ForeColor = SystemColors.ButtonHighlight;
            lbl_dados.Location = new Point(6, 62);
            lbl_dados.Name = "lbl_dados";
            lbl_dados.Size = new Size(1160, 32);
            lbl_dados.TabIndex = 2;
            lbl_dados.Text = "DADOS DA INSPEÇÃO";
            // 
            // Txt_dados
            // 
            Txt_dados.AutoEllipsis = true;
            Txt_dados.Font = new Font("Segoe UI", 12F);
            Txt_dados.Location = new Point(6, 94);
            Txt_dados.Name = "Txt_dados";
            Txt_dados.Size = new Size(1154, 45);
            Txt_dados.TabIndex = 3;
            Txt_dados.Text = "Durante a avaliação da coleta seletiva, é fundamental verificar três aspectos principais: a identificação dos coletores, as condições físicas dos coletores e a correta separação dos resíduos.\r\n\r\n";
            // 
            // lbl_data
            // 
            lbl_data.Font = new Font("Segoe UI", 11F);
            lbl_data.ForeColor = Color.MidnightBlue;
            lbl_data.Location = new Point(12, 145);
            lbl_data.Name = "lbl_data";
            lbl_data.Size = new Size(49, 20);
            lbl_data.TabIndex = 4;
            lbl_data.Text = "DATA:";
            // 
            // lbl_turno
            // 
            lbl_turno.Font = new Font("Segoe UI", 11F);
            lbl_turno.ForeColor = Color.MidnightBlue;
            lbl_turno.Location = new Point(202, 145);
            lbl_turno.Name = "lbl_turno";
            lbl_turno.Size = new Size(61, 20);
            lbl_turno.TabIndex = 5;
            lbl_turno.Text = "TURNO:";
            // 
            // lbl_modulo
            // 
            lbl_modulo.Font = new Font("Segoe UI", 11F);
            lbl_modulo.ForeColor = Color.MidnightBlue;
            lbl_modulo.Location = new Point(404, 144);
            lbl_modulo.Name = "lbl_modulo";
            lbl_modulo.Size = new Size(74, 20);
            lbl_modulo.TabIndex = 7;
            lbl_modulo.Text = "MÓDULO:";
            // 
            // txt_modulo
            // 
            txt_modulo.BackColor = Color.DeepSkyBlue;
            txt_modulo.BorderStyle = BorderStyle.None;
            txt_modulo.Font = new Font("Segoe UI", 12F);
            txt_modulo.Location = new Point(484, 145);
            txt_modulo.Name = "txt_modulo";
            txt_modulo.Size = new Size(129, 22);
            txt_modulo.TabIndex = 8;
            // 
            // lbl_ilha
            // 
            lbl_ilha.Font = new Font("Segoe UI", 11F);
            lbl_ilha.ForeColor = Color.MidnightBlue;
            lbl_ilha.Location = new Point(619, 145);
            lbl_ilha.Name = "lbl_ilha";
            lbl_ilha.Size = new Size(51, 20);
            lbl_ilha.TabIndex = 9;
            lbl_ilha.Text = "ILHA:";
            // 
            // lbl_equipe
            // 
            lbl_equipe.Font = new Font("Segoe UI", 11F);
            lbl_equipe.ForeColor = Color.MidnightBlue;
            lbl_equipe.Location = new Point(801, 145);
            lbl_equipe.Name = "lbl_equipe";
            lbl_equipe.Size = new Size(74, 20);
            lbl_equipe.TabIndex = 11;
            lbl_equipe.Text = "EQUIPE:";
            // 
            // lbl_avaliador
            // 
            lbl_avaliador.Font = new Font("Segoe UI", 11F);
            lbl_avaliador.ForeColor = Color.MidnightBlue;
            lbl_avaliador.Location = new Point(12, 185);
            lbl_avaliador.Name = "lbl_avaliador";
            lbl_avaliador.Size = new Size(95, 20);
            lbl_avaliador.TabIndex = 15;
            lbl_avaliador.Text = "AVALIADOR:";
            // 
            // cb_avaliador
            // 
            cb_avaliador.BackColor = Color.DeepSkyBlue;
            cb_avaliador.FormattingEnabled = true;
            cb_avaliador.Items.AddRange(new object[] { "Operador", "Líder", "Supervisor/Gerente" });
            cb_avaliador.Location = new Point(103, 186);
            cb_avaliador.Name = "cb_avaliador";
            cb_avaliador.Size = new Size(129, 23);
            cb_avaliador.TabIndex = 17;
            // 
            // cb_equipe
            // 
            cb_equipe.BackColor = Color.DeepSkyBlue;
            cb_equipe.FormattingEnabled = true;
            cb_equipe.Items.AddRange(new object[] { "Buyoff", "Equipe 1 e 2", "Equipe 3 e 4", "Equipe 5", "Equipe 6", "Equipe 7" });
            cb_equipe.Location = new Point(862, 145);
            cb_equipe.Name = "cb_equipe";
            cb_equipe.Size = new Size(129, 23);
            cb_equipe.TabIndex = 18;
            // 
            // cb_ilha
            // 
            cb_ilha.BackColor = Color.DeepSkyBlue;
            cb_ilha.FormattingEnabled = true;
            cb_ilha.Items.AddRange(new object[] { "Ilha 1", "Ilha 2", "Ilha 3", "Ilha 4", "Ilha 5", "Ilha 7" });
            cb_ilha.Location = new Point(666, 145);
            cb_ilha.Name = "cb_ilha";
            cb_ilha.Size = new Size(129, 23);
            cb_ilha.TabIndex = 19;
            // 
            // label1
            // 
            label1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            label1.BackColor = Color.DodgerBlue;
            label1.Font = new Font("Arial", 18F);
            label1.ForeColor = SystemColors.ButtonHighlight;
            label1.Location = new Point(6, 220);
            label1.Name = "label1";
            label1.Size = new Size(1160, 32);
            label1.TabIndex = 20;
            label1.Text = "LISTA DE VERIFICAÇÃO";
            // 
            // btn_salvar
            // 
            btn_salvar.BackColor = Color.MidnightBlue;
            btn_salvar.BackgroundImageLayout = ImageLayout.None;
            btn_salvar.FlatStyle = FlatStyle.Flat;
            btn_salvar.Location = new Point(259, 186);
            btn_salvar.Name = "btn_salvar";
            btn_salvar.Size = new Size(139, 23);
            btn_salvar.TabIndex = 21;
            btn_salvar.Text = "Salvar";
            btn_salvar.UseVisualStyleBackColor = false;
            btn_salvar.Click += btn_salvar_Click;
            // 
            // lista_verificacao
            // 
            lista_verificacao.BackgroundColor = SystemColors.ButtonHighlight;
            lista_verificacao.BorderStyle = BorderStyle.None;
            lista_verificacao.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            lista_verificacao.GridColor = Color.MidnightBlue;
            lista_verificacao.Location = new Point(-2, 257);
            lista_verificacao.Name = "lista_verificacao";
            lista_verificacao.Size = new Size(1160, 412);
            lista_verificacao.TabIndex = 22;
            lista_verificacao.EditingControlShowing += lista_verificacao_EditingControlShowing;
            // 
            // cb_turno
            // 
            cb_turno.BackColor = Color.DeepSkyBlue;
            cb_turno.FormattingEnabled = true;
            cb_turno.Items.AddRange(new object[] { "1°", "2°" });
            cb_turno.Location = new Point(269, 142);
            cb_turno.Name = "cb_turno";
            cb_turno.Size = new Size(129, 23);
            cb_turno.TabIndex = 23;
            // 
            // Archive
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = SystemColors.ButtonHighlight;
            ClientSize = new Size(1184, 681);
            Controls.Add(cb_turno);
            Controls.Add(lista_verificacao);
            Controls.Add(btn_salvar);
            Controls.Add(label1);
            Controls.Add(cb_ilha);
            Controls.Add(cb_equipe);
            Controls.Add(cb_avaliador);
            Controls.Add(lbl_avaliador);
            Controls.Add(lbl_equipe);
            Controls.Add(lbl_ilha);
            Controls.Add(txt_modulo);
            Controls.Add(lbl_modulo);
            Controls.Add(lbl_turno);
            Controls.Add(lbl_data);
            Controls.Add(Txt_dados);
            Controls.Add(lbl_dados);
            Controls.Add(lbl_titulo);
            Controls.Add(txt_data);
            ForeColor = SystemColors.ControlLightLight;
            Name = "Archive";
            Text = "Avaliação da Coleta Seletiva";
            Load += Archive_Load;
            ((System.ComponentModel.ISupportInitialize)lista_verificacao).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private TextBox txt_data;
        private Label lbl_titulo;
        private Label lbl_dados;
        private Label Txt_dados;
        private Label lbl_data;
        private Label lbl_turno;
        private Label lbl_modulo;
        private TextBox txt_modulo;
        private Label lbl_ilha;
        private Label lbl_equipe;
        private Label lbl_avaliador;
        private ComboBox cb_avaliador;
        private ComboBox cb_equipe;
        private ComboBox cb_ilha;
        private Label label1;
        private Button btn_salvar;
        private DataGridView lista_verificacao;
        private ComboBox cb_turno;
    }
}
