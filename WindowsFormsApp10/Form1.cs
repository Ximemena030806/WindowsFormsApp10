using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using IronXL;

namespace WindowsFormsApp10
{
    public partial class Form1 : Form
    {
        string ruta = AppDomain.CurrentDomain.BaseDirectory;
        public Form1()
        {
            InitializeComponent();

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ruta += "file.xlsx";
        }

        private void btnConectar_Click(object sender, EventArgs e)
        {
            WorkBook workBook = WorkBook.Load(ruta);    
            WorkSheet workSheet = workBook.WorkSheets.First();

            int cellValue = workSheet["A2"].IntValue;

            foreach(var cell in workSheet["A1:C5"])
            {
                MessageBox.Show(cell.AddressString + "" + cell.Text);
            }
            decimal sum = workSheet["A2:A5"].Sum();
            MessageBox.Show(sum.ToString());

            decimal max = workSheet["A2:A5"].Max();
            MessageBox.Show(max.ToString());

            workSheet["A10"].Value = "8";
            workSheet.SetCellValue(1, 1, "prueba"); 

            workSheet.SaveAs(ruta);
        }
    }
}
