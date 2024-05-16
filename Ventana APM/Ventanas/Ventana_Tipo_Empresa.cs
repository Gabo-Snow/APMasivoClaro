using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Ventana_APM.Ventanas
{
    public partial class Ventana_Tipo_Empresa : Form
    {
        public Ventana_Tipo_Empresa()
        {


            DataTable dtEmp = new DataTable();
            // add column to datatable  

            dtEmp.Columns.Add("Tipo", typeof(string));
            
            InitializeComponent();
            List<string> vs = new List<string>();
            for (int i = 0; i < 4; i++)
            {
                vs.Add("hola");
                dtEmp.Rows.Add(new Object[] { "Hola"});
            }
            segmento_dgv.DataSource = dtEmp;
            DataGridViewCheckBoxColumn dgvCmb = new DataGridViewCheckBoxColumn();
            dgvCmb.ValueType = typeof(bool);
            dgvCmb.Name = "Chk";
            dgvCmb.HeaderText = "CheckBox";

            segmento_dgv.Columns.Add(dgvCmb);
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
    }
}
