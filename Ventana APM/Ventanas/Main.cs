using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Ventana_APM.MODELO.CLASES;
using Ventana_APM.MODELO.COLECCIONES;

namespace Ventana_APM
{
    public partial class Main : Form
    {
        private Button currentButton;
        private Random random = new Random();
        private int tempIndex;
        private Form activeForm;
        Coleccion_Cartera_Empresarial _Cartera_Empresarial = new Coleccion_Cartera_Empresarial();
        Coleccion_cartera coleccion_Cartera = new Coleccion_cartera();
        List<Cartera_Empresarial> cartera_Empresarials_sice = new List<Cartera_Empresarial>();
        List<Cartera> carteras_pablo;
        List<string> tipos_ = new List<string>();
        private BackgroundWorker bw_analizar;
        string aux_path = string.Empty;
        public Main()
        {
            InitializeComponent();
            //WindowState = FormWindowState.Maximized;
            this.bw_analizar = new BackgroundWorker();
            this.bw_analizar.DoWork += new DoWorkEventHandler(bw_DoWork_analizar);//este sirve para los metodos
            this.bw_analizar.ProgressChanged += new ProgressChangedEventHandler(bw_ProgressChanged_analizar);//este envia el progreso
            this.bw_analizar.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw_RunWorkerCompleted_analizar);//este cuando termina
            this.bw_analizar.WorkerReportsProgress = true;

            this.analizar_Btn.Click += new EventHandler(analizar_Btn_Click);

            button6.Visible = false;
            button1.Visible = false;
            
        }

        private void ActivateButton(object btnSender)
        {
            if (btnSender != null)
            {
                if (currentButton != (Button)btnSender)
                {
                    DisableButton();
                    Color color = SelectThemeColor();
                    currentButton = (Button)btnSender;
                    currentButton.BackColor = color;
                    currentButton.ForeColor = Color.White;
                    currentButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                    panelTitleBar.BackColor = color;
                    panelLogo.BackColor = ThemeColor.ChangeColorBrightness(color, -0.3);
                    ThemeColor.PrimaryColor = color;
                    ThemeColor.SecondaryColor = ThemeColor.ChangeColorBrightness(color, -0.3);
                   // btnCloseChildForm.Visible = true;
                }
            }
        }

        private Color SelectThemeColor()
        {
            int index = random.Next(ThemeColor.ColorList.Count);
            while (tempIndex == index)
            {
                index = random.Next(ThemeColor.ColorList.Count);
            }
            tempIndex = index;
            string color = ThemeColor.ColorList[index];
            return ColorTranslator.FromHtml(color);
        }
        //
        private void DisableButton()
        {
            foreach (Control previousBtn in panelMenu.Controls)
            {
                if (previousBtn.GetType() == typeof(Button))
                {
                    previousBtn.BackColor = Color.FromArgb(51, 51, 76);
                    previousBtn.ForeColor = Color.Gainsboro;
                    previousBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                }
            }
        }

        private void OpenChildForm(Form childForm, object btnSender)
        {
            if (activeForm != null)
                activeForm.Close();
            ActivateButton(btnSender);
            activeForm = childForm;
            childForm.TopLevel = false;
            childForm.FormBorderStyle = FormBorderStyle.None;
            childForm.Dock = DockStyle.Fill;
            this.panelDesktopPane.Controls.Clear();//asi y despues guardar reactivarlo
            this.panelDesktopPane.Controls.Add(childForm);
            this.panelDesktopPane.Tag = childForm;
            childForm.BringToFront();
            childForm.Show();
            
            // lblTitle.Text = childForm.Text;
        }
        private void btn_ventana_1_Click(object sender, EventArgs e)
        {
            if (tipos_.Count == 0)
            {
                System.Media.SystemSounds.Exclamation.Play();
                MessageBox.Show("¡Se deben cargar ambas carteras!");
            }
            else
            {
                OpenChildForm(new Ventanas.Ventana_Cargos_De_Activacion(cartera_Empresarials_sice, carteras_pablo, tipos_), sender);
            }
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (tipos_.Count == 0)
            {
                System.Media.SystemSounds.Exclamation.Play();
                MessageBox.Show("¡Se deben cargar ambas carteras!");
            }
            else
            {
                OpenChildForm(new Ventanas.Ventana_Cargos_de_Arrendamiento(cartera_Empresarials_sice, carteras_pablo, tipos_), sender);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (tipos_.Count == 0)
            {
                System.Media.SystemSounds.Exclamation.Play();
                MessageBox.Show("¡Se deben cargar ambas carteras!");
            }
            else
            {
                OpenChildForm(new Ventanas.Ajustes_Varios(cartera_Empresarials_sice, carteras_pablo, tipos_), sender);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (tipos_.Count == 0)
            {
                System.Media.SystemSounds.Exclamation.Play();
                MessageBox.Show("¡Se deben cargar ambas carteras!");
            }
            else
            {
                OpenChildForm(new Ventanas.Ventana_Compartidos(), sender);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (tipos_.Count == 0)
            {
                System.Media.SystemSounds.Exclamation.Play();
                MessageBox.Show("¡Se deben cargar ambas carteras!");
            }
            else
            {
                OpenChildForm(new Ventanas.Ventana_Gyp_Pro(), sender);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (tipos_.Count == 0)
            {
                System.Media.SystemSounds.Exclamation.Play();
                MessageBox.Show("¡Se deben cargar ambas carteras!");
            }
            else
            {
                OpenChildForm(new Ventanas.PruebaHilo(), sender);
            }
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            if (tipos_.Count == 0)
            {
                System.Media.SystemSounds.Exclamation.Play();
                MessageBox.Show("¡Se deben cargar ambas carteras!");
            }
            else
            {
                OpenChildForm(new Ventanas.Ajuste_CFM(), sender);
            }
        }

        private void analizar_Btn_Click(object sender, EventArgs e)
        {
            if (!this.bw_analizar.IsBusy)
            {
                this.bw_analizar.RunWorkerAsync();
                this.analizar_Btn.Enabled = false;
            }
        }

        private void bw_ProgressChanged_analizar(object sender, ProgressChangedEventArgs e)
        {
            //this.txt_cambiante.Text = e.ProgressPercentage.ToString() + "% complete";
            progressBar1.Value = e.ProgressPercentage;//izi pizzi funcionando el progress bar
            StringBuilder sb = new StringBuilder();
            if (e.ProgressPercentage == 10)
            {
                sb.AppendLine("Cargando datos de Cartera...");
                queestasucediendo_txt.Text = sb.ToString();
            }
            else if (e.ProgressPercentage == 25)
            {
                sb.AppendLine("Cargando datos de Cartera Empresarial...");
                queestasucediendo_txt.AppendText(Environment.NewLine + sb.ToString());
            }
            else if (e.ProgressPercentage == 40)
            {
                sb.AppendLine("Espere un momento...");
                queestasucediendo_txt.AppendText(Environment.NewLine + sb.ToString());
            }
            else if (e.ProgressPercentage == 80)
            {
                sb.AppendLine("Creado filtro...");
                queestasucediendo_txt.AppendText(Environment.NewLine + sb.ToString());
            }
            else if (e.ProgressPercentage == 100)
            {
                sb.AppendLine("Proceso Terminado");
                queestasucediendo_txt.AppendText(Environment.NewLine + sb.ToString());
            }

        }
        private void bw_RunWorkerCompleted_analizar(object sender, RunWorkerCompletedEventArgs e)
        {
            //this.txt_cambiante.Text = "The answer is: " + e.Result.ToString(); 2 asi recibo el resultado de bw_dowork

            //this.txt_cambiante.Text = "100" + "% complete";
            this.analizar_Btn.Enabled = true;//
            this.segmento_dgv.Visible = true;
            this.btn_Guardar.Visible = true;
            System.Media.SystemSounds.Exclamation.Play();
            MessageBox.Show("Proceso terminado!");
        }

        private void bw_DoWork_analizar(object sender, DoWorkEventArgs e)//ya esta andando el thread
        {
            BackgroundWorker worker = (BackgroundWorker)sender;
            int i = 10;
            worker.ReportProgress(i);
            if (cartera_empresas_txt.Text.Equals("") && cartera_txt.Text.Equals(""))
            {
                string message2 = "Ingrese las rutas de las carteras";
                MessageBox.Show(message2);
            }
            else
            {
                //esta lisito esta gestionao´pla pla plaash
                carteras_pablo = coleccion_Cartera.GenerarListado(cartera_txt.Text);
                i = 25;

                worker.ReportProgress(i);
                i = 40;

                worker.ReportProgress(i);//el hechiceroooooooooo con sus podeeresss sus.. GRAANDES PODERES!!!!
                cartera_Empresarials_sice = _Cartera_Empresarial.GenerarListado(cartera_empresas_txt.Text);

                i = 70;
                worker.ReportProgress(i);
                List<string> coleccion_empresas_segmentos = new List<string>();
                var coleccion_segmentos = cartera_Empresarials_sice.Select(x => x.SEGMENTO).Distinct().ToList();
                foreach (var item in coleccion_segmentos)
                {

                    coleccion_empresas_segmentos.Add(item.ToUpper());

                }
                //colecion_empresas_segmento.DeleteAll();
                //tasukete foland!
                //----------------------------------------------------------------------

                //--------
                i = 80;
                worker.ReportProgress(i);
                var coleccion_empresas2 = carteras_pablo.Select(x => x.SEGMENTO).Distinct().ToList();
                foreach (var item in coleccion_empresas2)
                {
                    coleccion_empresas_segmentos.Add(item.ToUpper());
                }

                i = 90;
                worker.ReportProgress(i);

                DataTable dtEmp = new DataTable();
                //

                dtEmp.Columns.Add("Tipo", typeof(string));
                dtEmp.Columns.Add("check", typeof(bool));

                //i want u in my room together laara lara lara boom boom boom bum
                //te lavaste con corazones pum pum pum
                foreach (var item in coleccion_empresas_segmentos.Distinct())
                {
                    dtEmp.Rows.Add(new Object[] { item, false });
                }

                i = 100;

                worker.ReportProgress(i);

                try
                {

                    segmento_dgv.DataSource = dtEmp;//aqui no entiendo bien que esta sucediendo aqui

                }
                catch (Exception ex)//mmmmmmmmmm
                {

                    throw;//aqui no entiendo bien que esta sucediendo aqui

                }
                
                //si se repite mas de 3 veces el proceso... ¡pla!, se va a la mierda...
                //....

            }
        }

        //......................................................................................

        private void btn_Guardar_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < segmento_dgv.RowCount - 1; i++)
            {
                if (segmento_dgv.Rows[i].Cells["Check"].Value.Equals(true))
                {
                    tipos_.Add(segmento_dgv.Rows[i].Cells["Tipo"].Value.ToString());
                }
            }

            segmento_dgv.Visible = false;
            btn_Guardar.Visible = false;
            string message = "Se agregaron los tipos";
            System.Media.SystemSounds.Exclamation.Play();
            MessageBox.Show(message);
            

        }

        private void empresarial_Btn_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = ".xlsx";
            //ofd.Filter;
            ofd.Filter = "Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (ofd.ShowDialog() == DialogResult.OK)
                aux_path = ofd.FileName;
            cartera_empresas_txt.Text = aux_path;
        }

        private void cartera_Btn_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = ".xlsx";
            ofd.Filter = "Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (ofd.ShowDialog() == DialogResult.OK)
                aux_path = ofd.FileName;
            cartera_txt.Text = aux_path;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            // OpenChildForm(new Ventanas.Corporaciones(), sender); 
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void cartera_empresas_txt_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
