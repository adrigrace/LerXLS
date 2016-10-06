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
            try
            {
                int cDuplic;
                int iItemchave = -1;
     
                int clin = 0;
                int ccol = 0;

                int icolunao = 0;
                int icolunac = 0;

                int caux = 0;

                DateTime datavalue;

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
                                    rnova["Erro"] = "Duplicidade";
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
                                    rnova["Erro"] = dtOrigem.Columns[icolunao].ColumnName.ToString() + " Diferente";
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
                                        rnova["Erro"] = dtOrigem.Columns[icolunao].ColumnName.ToString() + " Data Inválida";
                                        dtResultado.Rows.Add(rnova);
                                    }
                                }
                            }
                        }
                    }
                }

                dtgv1.DataSource = dtResultado;
            }
            catch 
            {
            }
            progressBar1.Visible = false;
            btn_comparar.Visible = true;
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
                                    //blinha = false;
                                    //sfalha = ex1.Message.ToString();
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
                                    //blinha = false;
                                    //sfalha = ex1.Message.ToString();
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
                                    //blinha = false;
                                    //sfalha = ex1.Message.ToString();
                                }
                            }
                        }


                        if (rnew.ItemArray[0].ToString() != "")
                        {
                            dt.Rows.Add(rnew);
                        }

                        //if (blinha == false)
                        //{
                        //    DataRow rnova = dtErro.NewRow();
                        //    for (ccol = 0; ccol < rnew.ItemArray.Count(); ccol++)
                        //    {
                        //        rnova[ccol] = rnew[ccol];
                        //    }
                        //    rnova["Erro"] = sfalha.ToString();
                        //    dtErro.Rows.Add(rnova);
                        //}
                    }
                    //ds.Tables.Add(dt);

                    //DataTable dt1;
                    //dt1 = dt.Clone();
                    //dt1 = dt.Copy();

                    //dtSemDup = RemoveDuplicateRows(dt, 1);
                    //dtDuplic = RemoveRows(dt1, 1);

                    //CriarExcel(nomesemextensao + "_resultado" + ".xls", dtSemDup);
                    //CriarExcel(nomesemextensao + "_duplicado" + ".xls", dtDuplic);
                    //CriarExcel(nomesemextensao + "_falha" + ".xls", dtErro);
                    break;
                }

                xlWorkBook.Close(true, null, null);
                xlApp.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);
                //return dt;
            }
            catch
            {
            }
            return dt;
        }

    }



    //private void button1_Click(object sender, EventArgs e)
    //{


    //    Stream myStream = null;
    //    string leitura = "Falha na leitura";

    //    OpenFileDialog openFileDialog1 = new OpenFileDialog();

    //    openFileDialog1.InitialDirectory = "c:\\";
    //    openFileDialog1.Filter = "xls files (*.xls)|*.xls|All files (*.*)|*.*";
    //    openFileDialog1.FilterIndex = 2;
    //    openFileDialog1.RestoreDirectory = true;


    //    if (openFileDialog1.ShowDialog() == DialogResult.OK)
    //    {
    //        try
    //        {
    //            if ((myStream = openFileDialog1.OpenFile()) != null)
    //            {
    //                leitura = LerExcel(openFileDialog1.FileName, Path.GetExtension(openFileDialog1.FileName), 
    //                    openFileDialog1.FileName.Replace(Path.GetExtension(openFileDialog1.FileName),""));

    //                if (leitura != "")
    //                {
    //                    MessageBox.Show(leitura);
    //                }
    //            }
    //        }
    //        catch (Exception ex)
    //        {
    //            MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
    //        }
    //    }
    //    if (leitura == "")
    //    {
    //        MessageBox.Show("Leitura concluída com sucesso");
    //    }
    //}

    //public string LerExcel(string NomeArquivo, string extensao, string nomesemextensao)
    //{
    //    try 
    //    {
    //        Excel.Application xlApp;
    //        Excel.Workbook xlWorkBook;
    //        Excel.Worksheet xlWorkSheet;
    //        Excel.Range range;

    //        DataTable dtSemDup;
    //        DataTable dtDuplic;
    //        string sfalha = "";

    //        int rCnt = 0;
    //        int cCnt = 0;
    //        int ccol = 0;

    //        xlApp = new Excel.Application();
    //        xlWorkBook = xlApp.Workbooks.Open(NomeArquivo, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
    //        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

    //        range = xlWorkSheet.UsedRange;

    //        DataSet ds = new DataSet();
    //        foreach (Excel.Worksheet sheet in xlWorkBook.Sheets)
    //        {
    //            DataTable dt = new DataTable(sheet.Name);
    //            DataTable dtErro = new DataTable(sheet.Name);

    //            for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
    //            {
    //                for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
    //                {
    //                    DataColumn myDataColumn;

    //                    myDataColumn = new DataColumn();
    //                    try
    //                    {
    //                        Type typex = xlWorkSheet.Cells[2, cCnt].Value.GetType();
    //                        myDataColumn.ColumnName = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
    //                        myDataColumn.DataType = Type.GetType(typex.ToString());
    //                    }
    //                    catch
    //                    {
    //                        myDataColumn.ColumnName = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
    //                        myDataColumn.DataType = Type.GetType("System.String");
    //                    }

    //                    dt.Columns.Add(myDataColumn);
    //                }
    //                break;
    //            }

    //            dtErro = dt.Clone();

    //            DataColumn dc;
    //            dc = new DataColumn();
    //            dc.ColumnName = "Erro";
    //            dc.DataType = Type.GetType("System.String");
    //            dtErro.Columns.Add(dc);

    //            Boolean blinha;
    //            sfalha = "";

    //            blinha = true;

    //            for (rCnt = 2; rCnt <= range.Rows.Count; rCnt++)
    //            {
    //                DataRow rnew = dt.NewRow();
    //                for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
    //                {

    //                    Type typex = null;
    //                    try
    //                    {
    //                        typex = xlWorkSheet.Cells[rCnt, cCnt].Value.GetType();
    //                    }
    //                    catch
    //                    {
    //                        typex = Type.GetType("System.String");
    //                    }


    //                    if (typex.FullName.ToString() == "System.String")
    //                    {
    //                        try
    //                        {
    //                            rnew[cCnt - 1] = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
    //                        }
    //                        catch (Exception ex1)
    //                        {
    //                            blinha = false;
    //                            sfalha = ex1.Message.ToString();
    //                        }
    //                    }

    //                    if (typex.FullName.ToString() == "System.Double")
    //                    {
    //                        try
    //                        {
    //                            rnew[cCnt - 1] = (double)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
    //                        }
    //                        catch (Exception ex1)
    //                        {
    //                            blinha = false;
    //                            sfalha = ex1.Message.ToString();
    //                        }
    //                    }

    //                    if (typex.FullName.ToString() == "System.DateTime")
    //                    {
    //                        try
    //                        {
    //                            rnew[cCnt - 1] = DateTime.FromOADate((range.Cells[rCnt, cCnt] as Excel.Range).Value2);
    //                        }
    //                        catch (Exception ex1)
    //                        {
    //                            blinha = false;
    //                            sfalha = ex1.Message.ToString();
    //                        }
    //                    }
    //                }
    //                dt.Rows.Add(rnew);

    //                if (blinha == false)
    //                {
    //                    DataRow rnova = dtErro.NewRow();
    //                    for (ccol = 0; ccol < rnew.ItemArray.Count(); ccol++)
    //                    {
    //                        rnova[ccol] = rnew[ccol];
    //                    }
    //                    rnova["Erro"] = sfalha.ToString();
    //                    dtErro.Rows.Add(rnova);
    //                }
    //            }
    //            ds.Tables.Add(dt);

    //            DataTable dt1;
    //            dt1 = dt.Clone();
    //            dt1 = dt.Copy();

    //            dtSemDup = RemoveDuplicateRows(dt, 1);
    //            dtDuplic = RemoveRows(dt1, 1);

    //            CriarExcel(nomesemextensao+ "_resultado"+".xls", dtSemDup);
    //            CriarExcel(nomesemextensao + "_duplicado" + ".xls", dtDuplic);
    //            CriarExcel(nomesemextensao + "_falha" + ".xls", dtErro);
    //            break;
    //        }

    //        xlWorkBook.Close(true, null, null);
    //        xlApp.Quit();

    //        releaseObject(xlWorkSheet);
    //        releaseObject(xlWorkBook);
    //        releaseObject(xlApp);
    //        return "";
    //    }
    //    catch
    //    {
    //        return "houve uma falha";
    //    }
    //}

    // public string CriarExcel(string NomeArquivo, DataTable DtConteudo)
    //{
    //    string retorno = "";
    //    Excel.Application XlObj = new Excel.Application();
    //    XlObj.Visible = false;
    //    Excel._Workbook WbObj = (Excel.Workbook)(XlObj.Workbooks.Add(""));
    //    Excel._Worksheet WsObj = (Excel.Worksheet)WbObj.ActiveSheet;

    //    try
    //    {
    //        int row = 1; int col = 1;
    //        foreach (DataColumn column in DtConteudo.Columns)
    //        {
    //            WsObj.Cells[row, col] = column.ColumnName;
    //            col++;
    //        }
    //        col = 1;
    //        row++;
    //        for (int i = 0; i < DtConteudo.Rows.Count; i++)
    //        {
    //            foreach (var cell in DtConteudo.Rows[i].ItemArray)
    //            {
    //                WsObj.Cells[row, col] = cell;
    //                col++;
    //            }
    //            col = 1;
    //            row++;
    //        }
    //        WbObj.SaveAs(NomeArquivo);
    //    }
    //    catch (Exception ex)
    //    {
    //        retorno = "Houve um erro na criação do arquivo. Consulte o administrador do sistema. n" + ex.Message;
    //    }
    //    finally
    //    {
    //        WbObj.Close();
    //    }

    //    return retorno; 
    //}

    //public DataTable RemoveDuplicateRows(DataTable dTable, int indice)
    //{
    //    Hashtable hTable = new Hashtable();
    //    ArrayList duplicateList = new ArrayList();

    //    foreach (DataRow drow in dTable.Rows)
    //    {
    //        if (hTable.Contains(drow[indice]))
    //        {
    //            duplicateList.Add(drow);
    //        }
    //        else
    //            hTable.Add(drow[indice], string.Empty);
    //    }

    //    foreach (DataRow dRow in duplicateList)
    //    {
    //        dTable.Rows.Remove(dRow);
    //    }
    //    return dTable;
    //}

    //public DataTable RemoveRows(DataTable dTable, int indice)
    //{
    //    Hashtable hTable = new Hashtable();
    //    ArrayList duplicateList = new ArrayList();

    //    foreach (DataRow drow in dTable.Rows)
    //    {
    //        if (hTable.Contains(drow[indice]))
    //        {
    //            duplicateList.Add(drow);
    //        }
    //        else
    //            hTable.Add(drow[indice], string.Empty);
    //    }

    //    DataTable dtAux;
    //    dtAux = dTable.Clone();


    //    foreach (DataRow dRow in duplicateList)
    //    {
    //        dtAux.Rows.Add(dRow.ItemArray);
    //    }
    //    return dtAux;
    //}
}
