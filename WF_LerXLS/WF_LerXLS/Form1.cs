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
            string leitura = "Houve um erro na leitura do arquivo. Consulte o administrador do sistema.";

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
                    MessageBox.Show("Houve um erro na leitura do arquivo. Consulte o administrador do sistema. n" + ex.Message);
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
            int iItemchave = -1;

            iItemchave = lstbox_campos.SelectedIndex;

            if ((lblarquivo_origem.Text.Trim() == "") || (lblarquivo_comparativo.Text.Trim() == ""))
            {
                MessageBox.Show("Para efetuar a comparação é necessário que sejam selecionados os arquivos origem e destino !");
            }
            else {
                if ((lblarquivo_origem.Text.Trim() != "") && (lblarquivo_comparativo.Text.Trim() != ""))
                {
                    if (iItemchave == -1)
                    {
                        DialogResult result = MessageBox.Show("Você realmente deseja fazer a comparação sem coluna de identificador único ?", "Questão", MessageBoxButtons.OKCancel);
                        if (result == DialogResult.OK)
                        {
                            Comparar();
                        }
                        else
                        {
                        }
                    }
                    else
                    {
                        MessageBox.Show("A comparação irá começar e pode demorar alguns minutos, por favor aguarde.");
                        Comparar();
                    }
                }
            }
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
                                    dtResultado.Rows.Add(rnova);
                                }
                            }
                        }

                        if (dtOrigem.Rows[clin].ItemArray[0].ToString() == dtComparativo.Rows[ccol].ItemArray[0].ToString())
                        {
                            for (icolunao = 0; icolunao < dtOrigem.Rows[clin].ItemArray.Count(); icolunao++)
                            {
                                //valida se é um cpf ou cnpj e se for verifica se o valor é válido
                                if (dtOrigem.Columns[icolunao].ColumnName.ToString() == "CPF")
                                {
                                    if ((dtComparativo.Rows[ccol].ItemArray[icolunao].ToString().Trim().Length < 11) ||
                                        (dtComparativo.Rows[ccol].ItemArray[icolunao].ToString().Trim().Length > 11) || 
                                        (IsCpf(dtComparativo.Rows[ccol].ItemArray[icolunao].ToString()) == false))
                                    {
                                        DataRow rnova = dtResultado.NewRow();
                                        for (caux = 0; caux < dtOrigem.Rows[clin].ItemArray.Count(); caux++)
                                        {
                                            rnova[caux] = dtOrigem.Rows[clin].ItemArray[caux];
                                        }
                                        sTemErro = dtOrigem.Columns[icolunao].ColumnName.ToString() + " Valor Inválido";
                                        rnova["Erro"] = sTemErro;
                                        dtResultado.Rows.Add(rnova);
                                    }
                                }

                                if (dtOrigem.Columns[icolunao].ColumnName.ToString() == "CNPJ")
                                {
                                    if ((dtComparativo.Rows[ccol].ItemArray[icolunao].ToString().Trim().Length < 14) ||
                                        (dtComparativo.Rows[ccol].ItemArray[icolunao].ToString().Trim().Length > 14) ||
                                        (IsCnpj(dtComparativo.Rows[ccol].ItemArray[icolunao].ToString()) == false))
                                    {
                                        DataRow rnova = dtResultado.NewRow();
                                        for (caux = 0; caux < dtOrigem.Rows[clin].ItemArray.Count(); caux++)
                                        {
                                            rnova[caux] = dtOrigem.Rows[clin].ItemArray[caux];
                                        }
                                        sTemErro = dtOrigem.Columns[icolunao].ColumnName.ToString() + " Valor Inválido";
                                        rnova["Erro"] = sTemErro;
                                        dtResultado.Rows.Add(rnova);
                                    }
                                }

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

                dtResumo = RemoveDuplicateRows(dtResumo, 0);
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
            Boolean bleitura = false;
            string saux;

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
                    progressBar1.Step = (1000 / range.Rows.Count);
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

                                if (((string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2 == "CPF") || ((string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2 == "CNPJ"))
                                {
                                    myDataColumn.DataType = Type.GetType("System.String");
                                }
                                else
                                {
                                    myDataColumn.DataType = Type.GetType(typex.ToString());
                                }
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
                            string sname = "";

                            try
                            {
                                typex = xlWorkSheet.Cells[rCnt, cCnt].Value.GetType();
                                sname = (string)(range.Cells[1, cCnt] as Excel.Range).Value2;
                            }
                            catch
                            {
                                typex = Type.GetType("System.String");
                            }

                            if (typex.FullName.ToString() == "System.String")
                            {
                                try
                                {
                                    saux = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                                    rnew[cCnt - 1] = saux;

                                    if (cCnt == 1)
                                    {
                                        if (saux.Trim().Length < 4)
                                        {
                                            //a linha não será lida pois não contém o mínimo de 3 caracteres na primeira coluna
                                            bleitura = true;
                                        }
                                    }
                                }
                                catch 
                                {
                                    if (cCnt == 1)
                                    {
                                        bleitura = true;
                                    }
                                }
                            }

                            if (typex.FullName.ToString() == "System.Double")
                            {
                                try
                                {
                                    rnew[cCnt - 1] = (double)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                                }
                                catch { }
                            }

                            if (typex.FullName.ToString() == "System.DateTime")
                            {
                                try
                                {
                                    rnew[cCnt - 1] = DateTime.FromOADate((range.Cells[rCnt, cCnt] as Excel.Range).Value2);
                                }
                                catch { }
                            }

                            if (bleitura == true)
                            {
                                break;
                            }
                        }

                        if (bleitura == true)
                        {
                            break;
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

        public static bool IsCnpj(string cnpj)
        {
            string CNPJ = cnpj.Replace(".", "");
            CNPJ = CNPJ.Replace("/", "");
            CNPJ = CNPJ.Replace("-", "");

            int[] digitos, soma, resultado;
            int nrDig;
            string ftmt;
            bool[] CNPJOk;

            ftmt = "6543298765432";
            digitos = new int[14];
            soma = new int[2];
            soma[0] = 0;
            soma[1] = 0;
            resultado = new int[2];
            resultado[0] = 0;
            resultado[1] = 0;
            CNPJOk = new bool[2];
            CNPJOk[0] = false;
            CNPJOk[1] = false;

            try
            {
                for (nrDig = 0; nrDig < 14; nrDig++)
                {
                    digitos[nrDig] = int.Parse(
                     CNPJ.Substring(nrDig, 1));
                    if (nrDig <= 11)
                        soma[0] += (digitos[nrDig] *
                        int.Parse(ftmt.Substring(
                          nrDig + 1, 1)));
                    if (nrDig <= 12)
                        soma[1] += (digitos[nrDig] *
                        int.Parse(ftmt.Substring(
                          nrDig, 1)));
                }

                for (nrDig = 0; nrDig < 2; nrDig++)
                {
                    resultado[nrDig] = (soma[nrDig] % 11);
                    if ((resultado[nrDig] == 0) || (resultado[nrDig] == 1))
                        CNPJOk[nrDig] = (
                        digitos[12 + nrDig] == 0);

                    else
                        CNPJOk[nrDig] = (
                        digitos[12 + nrDig] == (
                        11 - resultado[nrDig]));

                }

                return (CNPJOk[0] && CNPJOk[1]);

            }
            catch
            {
                return false;
            }
         }

        public static bool IsCpf(string cpf)
        {
            string valor = cpf.Replace(".", "");
            valor = valor.Replace("-", "");

            if (valor.Length != 11)
                return false;

            bool igual = true;
            for (int i = 1; i < 11 && igual; i++)
                if (valor[i] != valor[0])
                    igual = false;

            if (igual || valor == "12345678909")
                return false;

            int[] numeros = new int[11];
            for (int i = 0; i < 11; i++)
                numeros[i] = int.Parse(
                valor[i].ToString());

            int soma = 0;
            for (int i = 0; i < 9; i++)
                soma += (10 - i) * numeros[i];

            int resultado = soma % 11;
            if (resultado == 1 || resultado == 0)
            {
                if (numeros[9] != 0)
                    return false;
            }
            else if (numeros[9] != 11 - resultado)
                return false;

            soma = 0;
            for (int i = 0; i < 10; i++)
                soma += (11 - i) * numeros[i];

            resultado = soma % 11;

            if (resultado == 1 || resultado == 0)
            {
                if (numeros[10] != 0)
                    return false;

            }
            else
                if (numeros[10] != 11 - resultado)
                    return false;
            return true;
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
            Excel.Application XlObj = new Excel.Application();
            XlObj.Visible = false;
            Excel._Workbook WbObj = (Excel.Workbook)(XlObj.Workbooks.Add(""));
            Excel._Worksheet WsObj = (Excel.Worksheet)WbObj.ActiveSheet;

            Excel.Range celulas;
            string serro;

            try
            {
                int row = 1; int col = 1;
                foreach (DataColumn column in DtConteudo.Columns)
                {
                    if (col == 1)
                    {
                        WsObj.Cells[row, col] = "NP";
                        col++;
                    }
                    WsObj.Cells[row, col] = column.ColumnName;
                    col++;
                }
                col = 1;
                row++;

                for (int i = 0; i < DtConteudo.Rows.Count; i++)
                {
                    foreach (var cell in DtConteudo.Rows[i].ItemArray)
                    {
                        if (col==1)
                        {
                            WsObj.Cells[row, col] = (row - 1);
                        }
                        WsObj.Cells[row, (col+1)] = cell;
                        col++;
                    }
                    col = 1;
                    row++;
                }

                row = 0;
                col = 0;

                for (int i = 0; i < DtConteudo.Rows.Count; i++)
                {
                    for (int h = 0; h < DtConteudo.Rows[i].ItemArray.Count(); h++)
                    {
                        for (int j = 0; j < dtResultado.Rows.Count; j++)
                        {
                            for (int k = 0; k < dtResultado.Rows[j].ItemArray.Count(); k++)
                            {
                                if (DtConteudo.Rows[i].ItemArray[0].ToString() == dtResultado.Rows[j].ItemArray[0].ToString())
                                {
                                    serro = dtResultado.Rows[j]["Erro"].ToString();
                                    if (serro.Replace(" Valor Inválido", "").Replace(" Diferente", "").Replace(" Data Inválida", "") == dtResultado.Columns[k].ColumnName.ToString())
                                    {
                                        celulas = (Excel.Range)WsObj.Cells[(i + 2), (k + 2)];
                                        celulas.Interior.Color = ColorTranslator.ToWin32(Color.Red);
                                    }
                                    if (serro == "Duplicidade")
                                    {
                                        if (k == 0)
                                        {
                                            celulas = (Excel.Range)WsObj.Cells[(i + 2), (k + 2)];
                                            celulas.Interior.Color = ColorTranslator.ToWin32(Color.Red);
                                        }
                                    }
                                }
                            }
                        }
                        col++;
                    }
                    col = 1;
                    row++;
                }


                Excel._Worksheet WsObj1 = (Excel.Worksheet)XlObj.Worksheets.Add();
                WsObj1 = (Excel.Worksheet)WbObj.ActiveSheet;

                row = 1;
                col = 1;

                try
                {
                    foreach (DataColumn column in dtResultado.Columns)
                    {
                        WsObj1.Cells[row, col] = column.ColumnName;
                        col++;
                    }
                    col = 1;
                    row++;

                    for (int i = 0; i < dtResultado.Rows.Count; i++)
                    {
                        foreach (var cell in dtResultado.Rows[i].ItemArray)
                        {
                            WsObj1.Cells[row, col] = cell;
                            col++;
                        }
                        col = 1;
                        row++;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Houve um erro na criação do arquivo. Consulte o administrador do sistema. n" + ex.Message);
                }

                WbObj.SaveAs(NomeArquivo);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Houve um erro na criação do arquivo. Consulte o administrador do sistema. n" + ex.Message);
            }
            finally
            {
                WbObj.Close();
                releaseObject(WbObj);
            }
        }

        public DataTable RemoveDuplicateRows(DataTable dTable, int indice)
        {
            Hashtable hTable = new Hashtable();
            ArrayList duplicateList = new ArrayList();

            foreach (DataRow drow in dTable.Rows)
            {
                if (hTable.Contains(drow[indice]))
                {
                    duplicateList.Add(drow);
                }
                else
                    hTable.Add(drow[indice], string.Empty);
            }

            foreach (DataRow dRow in duplicateList)
            {
                dTable.Rows.Remove(dRow);
            }
            return dTable;
        }
    }
}
