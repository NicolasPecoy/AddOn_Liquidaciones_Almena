using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FacturacionElectronica;
using AddOn_Liquidaciones_Almena;

namespace FactElectronicaSICFE
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            SboClass oInfoAddon;
            oInfoAddon = new SboClass();
        }

        private void Form1_Paint(object sender, PaintEventArgs e)
        {
            this.Visible = false;
            this.Hide();
        }
    }
}
