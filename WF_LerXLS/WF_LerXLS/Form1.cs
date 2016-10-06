using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;

using Excel = Microsoft.Office.Interop.Excel;
using System.Collections;
using System.Reflection; 

namespace WF_LerXLS
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        
        private void btn_Arquivo_Original_Click(object sender, EventArgs e)
        {
            Stream myStream = null;
            string leitura = "Falha na leitura";

            lstbox_campos.Items.Clear();
            dtgv1.DataSource = null;


            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "xls files (*.xls)|*.xls|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                lblarquivo_origem.Text = openFileDialog1.FileName.ToString();

                try
                {
                    if ((myStream = openFileDialog1.OpenFile()) != null)
                    {
                        leitura = TrazerItens(openFileDialog1.FileName, Path.GetExtension(openFileDialog1.FileName),
                            openFileDialog1.FileName.Replace(Path.GetExtension(openFileDialog1.FileName), ""));

                        if (leitura != "")
                        {
                            MessageBox.Show(leitura);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }

        public string TrazerItens(string NomeArquivo, string extensao, string nomesemextensao)
        {
            try
            {
                
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                Excel.Range range;

                int rCnt = 0;
                int cCnt = 0;

                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(NomeArquivo, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                range = xlWorkSheet.UsedRange;

                //lstbox_campos.Items.Add("Nenhuma");

                DataSet ds = new DataSet();
                foreach (Excel.Worksheet sheet in xlWorkBook.Sheets)
                {
                    DataTable dt = new DataTable(sheet.Name);
                   
                    for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
                    {
                        for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                        {
                            try
                            {
                                lstbox_campos.Items.Add((string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2);
                            }
                            catch
                            {
                                lstbox_campos.Items.Add("Não Identificado");
                            }
                        }
                        break;
                    }

                    break;
                }

                xlWorkBook.Close(true, null, null);
                xlApp.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);
                return "";
            }
            catch
            {
                return "houve uma falha";
            }
        }

        private void btn_Arquivo_Comparativo_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "xls files (*.xls)|*.xls|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                lblarquivo_comparativo.Text = openFileDialog1.FileName.ToString();
            }
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void btn_comparar_Click(object sender, EventArgs e)
        {
            Comparar();
        }

        private DataTable Comparar()
        {
            DataTable dtsaida = new DataTable();
            try
            {
                int cDuplic;
                int iItemchave = -1;
     
                int clin = 0;
                int ccol = 0;

                int icolunao = 0;
                int caux = 0;
                string serro = "";

                DateTime datavalue;
                double valor;
                string sTemErro = "";

                try
                {
                    iItemchave = lstbox_campos.SelectedIndex;
                }
                catch
                {
                    iItemchave = -1;
                }

                DataTable dtOrigem = LerExcel(lblarquivo_origem.Text);
                DataTable dtComparativo = LerExcel(lblarquivo_comparativo.Text);
                DataTable dtResultado = new DataTable();

                DataTable dtResumo = new DataTable();

                dtResultado = dtOrigem.Clone();

                DataColumn dc;
                dc = new DataColumn();
                dc.ColumnName = "Erro";
                dc.DataType = Type.GetType("System.String");
                dtResultado.Columns.Add(dc);

                dc = new DataColumn();
                dc.ColumnName = "Linha";
                dc.DataType = Type.GetType("System.Double");
                dtResultado.Columns.Add(dc);

                dc = new DataColumn();
                dc.ColumnName = "Coluna";
                dc.DataType = Type.GetType("System.Double");
                dtResultado.Columns.Add(dc);

                progressBar1.Visible = true;
                btn_comparar.Visible = false;

                for (clin = 0; clin < dtOrigem.Rows.Count; clin++)
                {
                    sTemErro = "";
                    cDuplic = 0;

                    for (ccol = 0; ccol < dtComparativo.Rows.Count; ccol++)
                    {
                        //pode validar duplicidade
                        if (iItemchave >= 0)
                        {
                            //significa q encontrou um registro igual na outra tabela
                            if (dtOrigem.Rows[clin].ItemArray[iItemchave].ToString() == dtComparativo.Rows[ccol].ItemArray[iItemchave].ToString())
                            {
                                cDuplic = cDuplic + 1;
                                if(cDuplic > 1)
                                {
                                    DataRow rnova = dtResultado.NewRow();
                                    for (caux = 0; caux < dtOrigem.Rows[clin].ItemArray.Count(); caux++)
                                    {
                                        rnova[caux] = dtOrigem.Rows[clin].ItemArray[caux];
                                    }
                                    sTemErro = "Duplicidade";
                                    rnova["Erro"] = sTemErro;
                                    rnova["Linha"] = clin;
                                    rnova["Coluna"] = caux;
                                    dtResultado.Rows.Add(rnova);
                                }
                            }
                        }

                        if (dtOrigem.Rows[clin].ItemArray[0].ToString() == dtComparativo.Rows[ccol].ItemArray[0].ToString())
                        {
                            for (icolunao = 0; icolunao < dtOrigem.Rows[clin].ItemArray.Count(); icolunao++)
                            {
                                //valida se o valor da coluna relativa é diferente
                                if (dtOrigem.Rows[clin].ItemArray[icolunao].ToString() != dtComparativo.Rows[ccol].ItemArray[icolunao].ToString())
                                {
                                    DataRow rnova = dtResultado.NewRow();
                                    for (caux = 0; caux < dtOrigem.Rows[clin].ItemArray.Count(); caux++)
                                    {
                                        rnova[caux] = dtOrigem.Rows[clin].ItemArray[caux];
                                    }
                                    sTemErro = dtOrigem.Columns[icolunao].ColumnName.ToString() + " Diferente";
                                    rnova["Erro"] = sTemErro;
                                    rnova["Linha"] = clin;
                                    rnova["Coluna"] = caux;
                                    dtResultado.Rows.Add(rnova);
                                }

                                //valida o tipo da coluna
                                if (dtOrigem.Columns[icolunao].DataType.ToString() == "System.DateTime")
                                {
                                    if (!DateTime.TryParse(dtComparativo.Rows[ccol].ItemArray[icolunao].ToString(), out datavalue)) 
                                    {
                                        DataRow rnova = dtResultado.NewRow();
                                        for (caux = 0; caux < dtOrigem.Rows[clin].ItemArray.Count(); caux++)
                                        {
                                            rnova[caux] = dtOrigem.Rows[clin].ItemArray[caux];
                                        }
                                        sTemErro = dtOrigem.Columns[icolunao].ColumnName.ToString() + " Data Inválida";
                                        rnova["Erro"] = sTemErro;
                                        rnova["Linha"] = clin;
                                        rnova["Coluna"] = caux;
                                        dtResultado.Rows.Add(rnova);
                                    }
                                }

                                //valida o tipo da coluna
                                if (dtOrigem.Columns[icolunao].DataType.ToString() == "System.Double")
                                {
                                    if (!double.TryParse(dtComparativo.Rows[ccol].ItemArray[icolunao].ToString(), out valor))  
                                    {
                                        DataRow rnova = dtResultado.NewRow();
                                        for (caux = 0; caux < dtOrigem.Rows[clin].ItemArray.Count(); caux++)
                                        {
                                            rnova[caux] = dtOrigem.Rows[clin].ItemArray[caux];
                                        }
                                        sTemErro = dtOrigem.Columns[icolunao].ColumnName.ToString() + " Valor Inválido";
                                        rnova["Erro"] = sTemErro;
                                        rnova["Linha"] = clin;
                                        rnova["Coluna"] = caux;
                                        dtResultado.Rows.Add(rnova);
                                    }
                                }
                            }
                        }
                    }
                }

                dtgv1.DataSource = dtResultado;

                DataColumn dcRes;
                dcRes = new DataColumn();
                dcRes.ColumnName = "Tipo_de_Erro";
                dcRes.DataType = Type.GetType("System.String");
                dtResumo.Columns.Add(dcRes);

                dcRes = new DataColumn();
                dcRes.ColumnName = "Total";
                dcRes.DataType = Type.GetType("System.Double");
                dtResumo.Columns.Add(dcRes);

                for (clin = 0; clin < dtResultado.Rows.Count; clin++)
                {
                    serro = "";
                    cDuplic = 0;
                    serro = dtResultado.Rows[clin]["Erro"].ToString();
                    DataRow drResumo = dtResumo.NewRow();

                    for (ccol = 0; ccol < dtResultado.Rows.Count; ccol++)
                    {
                        if (dtResultado.Rows[ccol]["Erro"].ToString().Trim() == serro.Trim())
                        {
                            cDuplic = cDuplic + 1;
                        }
                    }

                    drResumo[0] = serro;
                    drResumo[1] = cDuplic;
                    dtResumo.Rows.Add(drResumo);
                }
                dtgv2.DataSource = dtResumo;

                dtsaida = dtResultado.Clone();
                dtsaida = dtResultado.Copy();
            }
            catch (Exception ex)
            {
            }
            progressBar1.Visible = false;
            btn_comparar.Visible = true;
            return dtsaida;
        }

        public DataTable LerExcel(string NomeArquivo)
        {
            DataTable dt = new DataTable();

            try
            {
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                Excel.Range range;

                int rCnt = 0;
                int cCnt = 0;

                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(NomeArquivo, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                range = xlWorkSheet.UsedRange;

                DataSet ds = new DataSet();
                foreach (Excel.Worksheet sheet in xlWorkBook.Sheets)
                {
                    dt = new DataTable(sheet.Name);
                    DataTable dtErro = new DataTable(sheet.Name);

                    progressBar1.Maximum = range.Rows.Count;
                    progressBar1.Step = (100 / range.Rows.Count);
                    progressBar1.Value = 0;
                    progressBar1.Visible = true;
                    btn_comparar.Visible = false;

                    for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
                    {
                        progressBar1.PerformStep();

                        for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                        {

                            DataColumn myDataColumn;

                            myDataColumn = new DataColumn();
                            try
                            {
                                Type typex = xlWorkSheet.Cells[2, cCnt].Value.GetType();
                                myDataColumn.ColumnName = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                                myDataColumn.DataType = Type.GetType(typex.ToString());
                            }
                            catch
                            {
                                myDataColumn.ColumnName = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                                myDataColumn.DataType = Type.GetType("System.String");
                            }

                            dt.Columns.Add(myDataColumn);
                        }
                        break;
                    }

                  

                    for (rCnt = 2; rCnt <= range.Rows.Count; rCnt++)
                    {
                        progressBar1.PerformStep();

                        DataRow rnew = dt.NewRow();
                        for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                        {

                            Type typex = null;
                            try
                            {
                                typex = xlWorkSheet.Cells[rCnt, cCnt].Value.GetType();
                            }
                            catch
                            {
                                typex = Type.GetType("System.String");
                            }


                            if (typex.FullName.ToString() == "System.String")
                            {
                                try
                                {
                                    rnew[cCnt - 1] = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                                }
                                catch 
                                {
                                   
                                }
                            }

                            if (typex.FullName.ToString() == "System.Double")
                            {
                                try
                                {
                                    rnew[cCnt - 1] = (double)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                                }
                                catch 
                                {
                                
                                }
                            }

                            if (typex.FullName.ToString() == "System.DateTime")
                            {
                                try
                                {
                                    rnew[cCnt - 1] = DateTime.FromOADate((range.Cells[rCnt, cCnt] as Excel.Range).Value2);
                                }
                                catch 
                                {
                                  
                                }
                            }
                        }


                        if (rnew.ItemArray[0].ToString() != "")
                        {
                            dt.Rows.Add(rnew);
                        }
                    }
                    break;
                }

                xlWorkBook.Close(true, null, null);
                xlApp.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);
            }
            catch
            {
            }
            return dt;
        }

        private void btn_Exporta_Excel_Click(object sender, EventArgs e)
        {
            DataTable dtOrigem = LerExcel(lblarquivo_comparativo.Text);
            DataTable dtResultado = new DataTable();

            dtResultado = dtgv1.DataSource as DataTable;
            CriarExcel(lblarquivo_origem.Text.Replace(Path.GetExtension(lblarquivo_origem.Text),"")+"_resultado.xls",dtOrigem, dtResultado) ;
        }

        public void CriarExcel(string NomeArquivo, DataTable DtConteudo, DataTable dtResultado)
        {
            string retorno = "";
            Excel.Application XlObj = new Excel.Application();
            XlObj.Visible = false;
            Excel._Workbook WbObj = (Excel.Workbook)(XlObj.Workbooks.Add(""));
            Excel._Worksheet WsObj = (Excel.Worksheet)WbObj.ActiveSheet;

            Excel.Range celulas;

            try
            {
                int row = 1; int col = 1;
                foreach (DataColumn column in DtConteudo.Columns)
                {
                    WsObj.Cells[row, col] = column.ColumnName;
                    col++;
                }
                col = 1;
                row++;

                for (int i = 0; i < DtConteudo.Rows.Count; i++)
                {
                    foreach (var cell in DtConteudo.Rows[i].ItemArray)
                    {
                        foreach(DataRow dr in dtResultado.Rows)
                        {
                            if (Convert.ToInt32(dr["Linha"].ToString()) == row)
                            {
                                if (Convert.ToInt32(dr["Coluna"].ToString()) == col)
                                {
                                    celulas = (Excel.Range)WsObj.Cells[row, col];
                                    celulas.Interior.Color = ColorTranslator.ToWin32(Color.Red);
                                }
                            }
                        }


                        WsObj.Cells[row, col] = cell;
                        col++;
                    }
                    col = 1;
                    row++;
                }
                WbObj.SaveAs(NomeArquivo);
            }
            catch (Exception ex)
            {
                retorno = "Houve um erro na criação do arquivo. Consulte o administrador do sistema. n" + ex.Message;
            }
            finally
            {
                WbObj.Close();
            }
        }

    }
}
