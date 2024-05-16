using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Ventana_APM.MODELO.CLASES;
using Ventana_APM.MODELO.COLECCIONES;

namespace Ventana_APM.Ventanas
{
    public partial class PruebaHilo : Form
    {
        // Declare our worker thread
        private BackgroundWorker bw;
        List<Cartera> carteras_pablo;

        public PruebaHilo()
        {
            InitializeComponent();
            //----------------------todo esto es necesario para el thread del progress bar
            this.bw = new BackgroundWorker();
            this.bw.DoWork += new DoWorkEventHandler(bw_DoWork);//este sirve para los metodos
            this.bw.ProgressChanged += new ProgressChangedEventHandler(bw_ProgressChanged);//este envia el progreso
            this.bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw_RunWorkerCompleted);//este cuando termina
            this.bw.WorkerReportsProgress = true;
            //----------------------------------------------------------------------------

            this.button1.Click += new EventHandler(button1_Click);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!this.bw.IsBusy)
            {
                this.bw.RunWorkerAsync();
                this.button1.Enabled = false;
            }
        }

        private void bw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            this.txt_cambiante.Text = e.ProgressPercentage.ToString() + "% complete";
            progressBar1.Value = e.ProgressPercentage;//izi pizzi funcionando el progress bar
            
        }

        private void bw_DoWork(object sender, DoWorkEventArgs e)//ya esta andando el thread para cargar barritas :B
        {

            Coleccion_cartera coleccion_Cartera = new Coleccion_cartera();
            carteras_pablo = coleccion_Cartera.GenerarListado(@"C:\Users\Marcelo\Desktop\Ciclo 18 entrada\Cartera 08.02.2021 _PABLO.xlsx"); //esto tiene que tener un contadorsss
            BackgroundWorker worker = (BackgroundWorker)sender;

            for (int i = 0; i < 100; ++i)
            {
                // report your progres
                worker.ReportProgress(i);//asi se envian los cambios
                
                System.Threading.Thread.Sleep(100);
            }
            e.Result = "42"; //1 asi envio algo 
        }

        private void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //this.txt_cambiante.Text = "The answer is: " + e.Result.ToString(); 2 asi recibo el resultado de bw_dowork

            this.txt_cambiante.Text = "100" + "% complete";
            this.button1.Enabled = true;
        }
    }
}
