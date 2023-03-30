using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsFormsApp1.Classes;
using System.IO;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        SWAPI objSW = new SWAPI();
        string nameArcPRT = @"C:\Users\Lucas Rodrigues\SKA AUTOMACAO DE ENGENHARIAS LTDA\Desenvolvimento Dassault - General\Documentos sobre o time\Integração de colaboradores\Treinamentos\5 - Treinamento API\Códigos fonte\SolidAPI\Peças\Garra.SLDPRT";
        string nameArcDRW = @"C:\Users\Lucas Rodrigues\SKA AUTOMACAO DE ENGENHARIAS LTDA\Desenvolvimento Dassault - General\Documentos sobre o time\Integração de colaboradores\Treinamentos\5 - Treinamento API\Códigos fonte\SolidAPI\Peças\Garra.SLDDRW";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void btnAbrirSW_Click(object sender, EventArgs e)
        {

            // Substituir (Visivel.Checked, Versão do seu SW);


            try
            {
                labelSW.Text = "Abrindo SolidWorks...";
                objSW.AbrirSolidworks(Visivel.Checked, 31);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ERRO", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                labelSW.Text = "Pronto!";
            }
        }

        private void chkinv_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void btnFecharSW_Click(object sender, EventArgs e)
        {
            try
            {
                labelSW.Text = "Fechando SolidWorks...";
                objSW.FecharSolidworks();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "ERRO", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                labelSW.Text = "Pronto!";
            }

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void btnAbrir_Click(object sender, EventArgs e)
        {
            try
            {
                objSW.AbrirArquivo(nameArcPRT, Path.GetExtension(nameArcPRT));
                MessageBox.Show(objSW.ObtemArquivoAtivo());

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "ERRO", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnFechar_Click(object sender, EventArgs e)
        {
            try
            {
                labelSW.Text = "Fechando arquivo...";
                objSW.FecharArquivo(nameArcPRT);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "ERRO", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                labelSW.Text = "Pronto!";
            }
        }

        private void btnExportar_Click(object sender, EventArgs e)
        {
            try
            {
                //runs from a drw archive
                //labelSW.Text = "Exportando para JPG..."
                //objSW.ExportarParaJPG(nameArcDRW.ToUpper().Replace((".SLDDRW", ".JPG"));

                //runs from a prt archive
                labelSW.Text = "Exportando para DWG...";
                objSW.ExportarParaDWG(nameArcDRW, nameArcPRT.ToUpper().Replace(".SLDPRT", ".DWG"));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ERRO", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                labelSW.Text = "Pronto!";
            }
        }

        private void btnBuscaProp_Click(object sender, EventArgs e)
        {
            try
            {
                labelSW.Text = "Busca o valor de uma propriedade....";
                MessageBox.Show(objSW.BuscaPropriedade("Autor"), "Valor Encontrado", MessageBoxButtons.OK);
                MessageBox.Show(objSW.BuscaPropriedade("Autor"), "Valor Encontrado", MessageBoxButtons.OK);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ERRO", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            finally
            {
                labelSW.Text = "Pronto!";
            }
        }

        private void btnCria_AtualizaProp_Click(object sender, EventArgs e)
        {
            try
            {
                labelSW.Text = "Gravando/Att o valor de uma propriedade...";
                objSW.Cria_AtualizaPropriedade("Autor", "Lucas");
                objSW.Cria_AtualizaPropriedade("Código", "MA25498632-00");
                MessageBox.Show(objSW.BuscaPropriedade("Autor"), "Valor Encontrado", MessageBoxButtons.OK);
                MessageBox.Show(objSW.BuscaPropriedade("Código"), "Valor Encontrado", MessageBoxButtons.OK);

                if (objSW.SalvaArquivo())
                {
                    MessageBox.Show("Arquivo salvo com sucesso!", "Sucesso!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ERRO", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                labelSW.Text = "Pronto!";
            }
        }
    }
}
