using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace Nomina
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void importarExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (ofdExcel.ShowDialog() == DialogResult.OK)
            {
                string archivo = ofdExcel.FileName;
                //MessageBox.Show("Archivo seleccionado: " + archivo);
                CargarExcel(archivo);

            }
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void tmrReloj_Tick(object sender, EventArgs e)
        {
            ttsHora.Text = DateTime.Now.ToString("hh:mm:ss tt");
        }

        private void CargarExcel(string path)
        {
            ExcelPackage.License.SetNonCommercialPersonal("Jose Luis Mota Espeleta");

            using (var package = new ExcelPackage(new System.IO.FileInfo(path)))
            {
                if (package.Workbook.Worksheets.Count == 0)
                {
                    MessageBox.Show("El archivo de Excel no contiene hojas de trabajo.");
                    return; //En caso que no haya hojas
                }

                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                DataTable dt = new DataTable();

                // Leer los encabezados de columna
                int colCount = worksheet.Dimension.End.Column;
                for(int i=0;i<=colCount;i++)
                {
                    dt.Columns.Add(worksheet.Cells[2, i + 1].Text);
                }

                // Leer las filas de datos
                int rowCount = worksheet.Dimension.End.Row;
                for (int i = 3; i < rowCount; i++)
                {
                    DataRow row = dt.NewRow();
                    for (int j = 1; j < colCount-1; j++)
                    {
                        row[j] = worksheet.Cells[i, j ].Text;
                    }
                    dt.Rows.Add(row);
                }


            }
        }
    }
}