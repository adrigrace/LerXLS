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

namespace WF_LerXLS
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Stream myStream = null;
            string leitura = "Falha na leitura";
                     
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "xls files (*.xls)|*.xls|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

     
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if ((myStream = openFileDialog1.OpenFile()) != null)
                    {
                        leitura = LerExcel(openFileDialog1.FileName, Path.GetExtension(openFileDialog1.FileName), 
                            openFileDialog1.FileName.Replace(Path.GetExtension(openFileDialog1.FileName),""));

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
            if (leitura == "")
            {
                MessageBox.Show("Leitura concluída com sucesso");
            }
        }

        public string LerExcel(string NomeArquivo, string extensao, string nomesemextensao)
        {
            try 
            {
			    Excel.Application xlApp;
			    Excel.Workbook xlWorkBook;
			    Excel.Worksheet xlWorkSheet;
			    Excel.Range range;

			    DataTable dtSemDup;
			    DataTable dtDuplic;

			    int rCnt = 0;
			    int cCnt = 0;

			    xlApp = new Excel.Application();
			    xlWorkBook = xlApp.Workbooks.Open(NomeArquivo, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
			    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

			    range = xlWorkSheet.UsedRange;

			    DataSet ds = new DataSet();
			    foreach (Excel.Worksheet sheet in xlWorkBook.Sheets)
			    {
				    DataTable dt = new DataTable(sheet.Name);

				    for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
				    {
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
					    dt.Rows.Add(rnew);
				    }
				    ds.Tables.Add(dt);

				    DataTable dt1;
				    dt1 = dt.Clone();
				    dt1 = dt.Copy();

				    dtSemDup = RemoveDuplicateRows(dt, 1);
				    dtDuplic = RemoveRows(dt1, 1);

                    CriarExcel(nomesemextensao+"_resultado"+extensao, dtSemDup);
                    CriarExcel(nomesemextensao + "_duplicado" + extensao, dtDuplic);
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

        public string CriarExcel(string NomeArquivo, DataTable DtConteudo)
        {
            try
            {
                Excel.Application xlApp;
			    Excel.Workbook xlWorkBook;
			    Excel.Worksheet xlWorkSheet;

                xlApp = new Excel.Application();
 
                object misValue = System.Reflection.Missing.Value;
 
                xlWorkBook = xlApp.Workbooks.Add(System.Reflection.Missing.Value); 
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets["Plan1"]; 
 
                int iLinha = 0;
                foreach (DataRow Dr in DtConteudo.Rows)
                {
                    iLinha = iLinha + 1;
                    int iColuna = 0;
                    foreach (DataColumn Dc in DtConteudo.Columns)
                    {
                        iColuna = iColuna + 1;
                        xlWorkSheet.Cells[iLinha, iColuna] = Dr[Dc.ColumnName];
                    }
                }

                xlWorkBook.SaveAs(NomeArquivo, Excel.XlFileFormat.xlWorkbookNormal,  misValue, misValue, false, misValue,  Excel.XlSaveAsAccessMode.xlNoChange, misValue, misValue, misValue,  misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
 
                xlApp.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);
             }
             catch (Exception e)
             {
                return "Houve um erro na criação do arquivo. Consulte o administrador do sistema. n" + e.Message;
             }
 
            return "";
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

        public DataTable RemoveRows(DataTable dTable, int indice)
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

            DataTable dtAux;
            dtAux = dTable.Clone();

            //foreach (DataRow dRow in dTable.Rows)
            //{
            //    dTable.Rows.Remove(dRow);
            //}
            //dTable.AcceptChanges();

            foreach (DataRow dRow in duplicateList)
            {
                dtAux.Rows.Add(dRow.ItemArray);
            }
            //dTable.AcceptChanges();
            return dtAux;
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
    }
}
