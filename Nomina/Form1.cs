using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Style.ThreeD;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace Nomina
{
    public partial class Form1 : Form
    {
        public string rutaJson = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "horas.json");

        List<string> columnas;
        public Form1()
        {
            InitializeComponent();
            importarExcelToolStripMenuItem.Enabled = false;
            columnas =new List<string>();
            columnas.Add("EE");
            columnas.Add("Nombre");
            columnas.Add("Apellido");
            columnas.Add("Rg. Hrs");
            columnas.Add("OT. Hrs.");
            columnas.Add("D Hrs.");
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
            DataTable dt = new DataTable();
            ExcelPackage.License.SetNonCommercialPersonal("Jose Luis Mota Espeleta");

            using (var package = new ExcelPackage(new System.IO.FileInfo(path)))
            {
                if (package.Workbook.Worksheets.Count == 0)
                {
                    MessageBox.Show("El archivo de Excel no contiene hojas de trabajo.");
                    return; //En caso que no haya hojas
                }

                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];



                // Leer los encabezados de columna
                foreach (var col in columnas)
                {
                    dt.Columns.Add(col);
                }

                // Leer las filas de datos
                int rowCount = 0;
                for (int i = 3; i < worksheet.Dimension.End.Row; i++)
                {
                    if (worksheet.Cells[i, 1].Text == "")
                        break;
                    else
                        rowCount = rowCount + 1;
                }
                rowCount = rowCount + 3;
                for (int i = 3; i < rowCount; i++)
                {
                    DataRow row = dt.NewRow();
                    int k = 1;
                    for (int j = 1; j < dt.Columns.Count + 1; j++)
                    {
                        if (worksheet.Cells[i, k].Text == "")
                            break;
                        row[j - 1] = worksheet.Cells[i, k].Text;
                        if (j == 2)
                        {
                            string[] partes = SepararNombre(worksheet.Cells[i, k].Text);
                            row[1] = partes[1];
                            row[2] = partes[0];
                            j++;
                        }
                        k++;
                    }
                    dt.Rows.Add(row);
                }
            }

            // Mostrar los datos en el DataGridView
            dgvInformacion.Rows.Clear();
            dgvInformacion.Rows.Add(dt.Rows.Count);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < 6; j++)
                {
                    dgvInformacion.Rows[i].Cells[j].Value = dt.Rows[i][j].ToString();
                }
            }

            calcularHoras();
        }

        private string [] SepararNombre(string nombreCompleto)
        { 
            
            string[] partes = nombreCompleto.Split(',');
            if (partes.Length >= 2)
            {
                string apellido = partes[0];
                string nombre = partes[1]; 
                partes[1] = nombre.Trim();               
            }
            return partes;
        }

        private void calcularHoras()
        {
            double regular = 0;
            double overTime = 0;
            double doubles = 0;
            if (File.Exists(rutaJson))
            {
                try
                {
                    string json = File.ReadAllText(rutaJson);
                    Datos datosJson = JsonConvert.DeserializeObject<Datos>(json);

                    if (datosJson != null)
                    {
                        regular = double.Parse(datosJson.regular.ToString());
                        overTime = double.Parse(datosJson.extras.ToString());
                        doubles = double.Parse(datosJson.dobles.ToString());
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ocurrió un error al cargar los datos", "Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            //Calcular horas y total
            for (int i = 0; i < dgvInformacion.Rows.Count; i++)
            {

                //if (dgvInformacion.Rows[i].IsNewRow) continue;

                double horaRegular = double.Parse(dgvInformacion.Rows[i].Cells[3].Value.ToString());
                double horasExtra = double.Parse(dgvInformacion.Rows[i].Cells[4].Value.ToString());
                double horasDobles = double.Parse(dgvInformacion.Rows[i].Cells[5].Value.ToString());

                dgvInformacion.Rows[i].Cells[6].Value = regular * horaRegular;
                dgvInformacion.Rows[i].Cells[7].Value = overTime * horasExtra;
                dgvInformacion.Rows[i].Cells[8].Value = doubles * horasDobles;

                dgvInformacion.Rows[i].Cells[9].Value = double.Parse(dgvInformacion.Rows[i].Cells[6].Value.ToString()) +
                    double.Parse(dgvInformacion.Rows[i].Cells[7].Value.ToString()) +
                    double.Parse(dgvInformacion.Rows[i].Cells[8].Value.ToString());
            }
        }

        private void insertarValoresAHorasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 form = new Form2();
            form.Show();
            importarExcelToolStripMenuItem.Enabled = true;
            tmrActualizacion.Enabled = true;        
        }

        private void tmrActualizacion_Tick(object sender, EventArgs e)
        {
            calcularHoras();
        }
    }
}