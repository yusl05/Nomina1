using Microsoft.Win32;
using Newtonsoft.Json;
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
    public partial class Form2 : Form
    {
        public string rutaJson = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "horas.json");
        public Form2()
        {
            InitializeComponent();
            CargarDatosDelJson();
        }

        private Datos ObtenerDatosDelFormulario()
        {
            try
            {
                double regular = double.Parse(tBRegular.Text);
                double extras = double.Parse(tBHrsExtra.Text);
                double dobles = double.Parse(tBDobles.Text);

                return new Datos { regular = regular, extras = extras, dobles = dobles };
            }
            catch (FormatException ex)
            {
                MessageBox.Show("Valores inválidos" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }
        }

        private void CargarDatosDelJson()
        {
            if (File.Exists(rutaJson))
            {
                try
                {
                    string json = File.ReadAllText(rutaJson);
                    Datos datosCargados = JsonConvert.DeserializeObject<Datos>(json);

                    if (datosCargados != null)
                    {
                        tBRegular.Text = datosCargados.regular.ToString();
                        tBHrsExtra.Text = datosCargados.extras.ToString();
                        tBDobles.Text = datosCargados.dobles.ToString();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ocurrió un error al cargar los datos", "Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        public void guardarJson(Datos horas)
        {
            var caracteristicas = new JsonSerializerSettings
            {
                Formatting = Formatting.Indented,
                DateFormatHandling = DateFormatHandling.IsoDateFormat,
                NullValueHandling = NullValueHandling.Ignore,
            };
            string json = JsonConvert.SerializeObject(horas, caracteristicas);
            File.WriteAllText(rutaJson, json);
        }

        private void btnAceptar_Click(object sender, EventArgs e)
        {
            Datos datosParaGuardar = ObtenerDatosDelFormulario();
            if (datosParaGuardar != null)
            {
                guardarJson(datosParaGuardar);
                MessageBox.Show("Datos guardados exitosamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);

                this.Close();
            }
        }

    }
}
