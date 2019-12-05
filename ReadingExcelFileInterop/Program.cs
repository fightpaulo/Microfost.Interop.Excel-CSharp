using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadingExcelFileInterop
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var inicio = DateTime.Now;
            var produtosDataTable = ReadExcelFile();
            var final = DateTime.Now;
        }

        public static System.Data.DataTable ReadExcelFile()
        {
            System.Data.DataTable produtos = new System.Data.DataTable();
            produtos.Columns.Add("Campo 1", typeof(object));
            produtos.Columns.Add("Campo 2", typeof(object));
            produtos.Columns.Add("Campo 3", typeof(object));
            produtos.Columns.Add("Campo 4", typeof(object));
            produtos.Columns.Add("Campo 5", typeof(object));
            produtos.Columns.Add("Campo 6", typeof(object));

            Application excelApp = new Application();
            Workbook workbook = excelApp.Workbooks.Open(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\BALANCETE SETEMBRO.xlsx");
            Worksheet template = (Worksheet)workbook.Sheets[1];
            Range range = null;
            List<object> listaAuxiliar = new List<object>();
            int columnIndex = 1;

            try
            {
                range = template.UsedRange;

                Range specificRangeAux = excelApp.Range[template.Cells[16, "C"], template.Cells[range.Rows.Count, "J"]];

                int rowCount = specificRangeAux.Rows.Count;
                int columnCount = specificRangeAux.Columns.Count;

                specificRangeAux = specificRangeAux.Resize[rowCount, columnCount];
                Array specificRange = (Array)specificRangeAux.Value[XlRangeValueDataType.xlRangeValueDefault];

                foreach (var i in specificRange)
                {
                    // As colunas 3 e 5 não possuem valores, por isso que os valores não são pegos
                    if (columnIndex != 3 && columnIndex != 5)
                    {
                        listaAuxiliar.Add(i);
                    }

                    // Se o indíce da coluna for 8, significa que a linha toda foi lida
                    // Então eu pego as informações e coloco em cada linha do DataTable
                    if (columnIndex % 8 == 0)
                    {
                        produtos.Rows.Add
                            (
                                listaAuxiliar[0], 
                                listaAuxiliar[1], 
                                listaAuxiliar[2],
                                listaAuxiliar[3],
                                listaAuxiliar[4],
                                listaAuxiliar[5]
                            );

                        listaAuxiliar = new List<object>();
                        columnIndex = 0;
                    }

                    columnIndex++;
                }
            }

            catch (Exception ex)
            {
                workbook.Close();
            }
            finally
            {
                workbook.Close();
            }

            return produtos;
        }
    }
}
