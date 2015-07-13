using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using PlanilhaExcel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Reflection;
using Microsoft.Office.Interop.Excel;

namespace WebServiceFET
{
    public class Excel
    {
        #region Construtor

        public Excel(DataSet ds)
        {
            createWorkbook();
            addWorksheet(ds, true);
        }

        #endregion

        #region Variáveis privadas

        private string extension = null;
        private Scripts scripts = new Scripts();
        private PlanilhaExcel.Workbook generatedFile;
        private PlanilhaExcel.Application app;
        private string[] path = new string[2];

        #endregion

        #region Métodos facilitadores de funcionalidade

        /// <summary>
        /// Retorna o nome de arquivo completo. Ex: file.xlsx
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="extension"></param>
        private string completeFileName(string fileName, string extension)
        {
            return fileName + "." + extension;
        }

        /// <summary>
        /// Define a extensão do arquivo que será salvo.
        /// </summary>
        /// <param name="oldOffice"></param>
        private void defineExtension(bool oldOffice)
        {
            if (oldOffice)
                extension = "xls";
            else
                extension = "xlsx";
        }

        public void defaultConfiguration(int sheetIndex, string planName = "Plan") {
            getWorksheet(sheetIndex).Name = planName;
            editSheet(sheetIndex, "A1").EntireRow.Font.Bold = true;
            editSheet(sheetIndex, "A1").EntireRow.ColumnWidth = 25;
            editSheet(sheetIndex, "A1").EntireRow.EntireColumn.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        }

        /// <summary>
        /// Retira da memória os objetos criados para geração da planilha do Excel.
        /// </summary>
        /// <param name="obj"></param>
        private void ReleaseObject(object obj)
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            GC.Collect();
        }
        
        #endregion

        #region Métodos de criação e manipulação

        /// <summary>
        /// Cria um novo arquivo Excel
        /// </summary>
        private void createWorkbook()
        {
            app = new PlanilhaExcel.Application();
            generatedFile = app.Workbooks.Add();
        }

        /// <summary>
        /// Disponibiliza a edição da aba solicitada.
        /// </summary>
        /// <param name="index"></param>
        /// <param name="celula1"></param>
        /// <param name="celula2"></param>
        public PlanilhaExcel.Range editSheet(int index, object celula1, object celula2 = null) {
            return (celula2 == null) ? getWorksheet(index).Range[celula1] : getWorksheet(index).Range[celula1, celula2];
        }

        /// <summary>
        /// Cria uma aba dentro de um arquivo Excel.
        /// </summary>
        /// <param name="ds"></param>
        public void addWorksheet(DataSet ds, bool constructorDataSet = false)
        {
            for (int i = 0; i < ds.Tables.Count; i++)
            {
                System.Data.DataTable data = ds.Tables[i];
                PlanilhaExcel.Worksheet worksheet = (constructorDataSet ? (PlanilhaExcel.Worksheet)generatedFile.Worksheets.Item[1] : (PlanilhaExcel.Worksheet)generatedFile.Worksheets.Add());
                for (int j = 0; j < data.Columns.Count; j++)
                    worksheet.Cells[1, (j + 1)] = ((DataColumn)data.Columns[j]).ColumnName.ToString();

                for (int j = 0; j < data.Rows.Count; j++)
                    for (int k = 0; k < data.Rows[j].ItemArray.Length; k++)
                        worksheet.Cells[1 + (j + 1), 1 + k] = data.Rows[j][k].ToString();
            }
        }

        /// <summary>
        /// Abre um arquivo de Excel.
        /// </summary>
        /// <param name="path"></param>
        public PlanilhaExcel.Workbook openExcel(string path, string fileName)
        {
            this.path[0] = path;
            this.path[1] = fileName;
            return new PlanilhaExcel.Application().Workbooks.Open(this.path[0] + this.path[1]);
        }

        #endregion

        #region Getters

        /// <summary>
        /// Retorna a aba solicitada.
        /// </summary>
        /// <param name="i"></param>
        public PlanilhaExcel.Worksheet getWorksheet(int i) {
            if(i <= 0 || getWorkbook().Worksheets.Count < i)
                throw new Exception();

            return (PlanilhaExcel.Worksheet)getWorkbook().Worksheets.Item[i];
        }

        /// <summary>
        /// Retorna o arquivo Excel.
        /// </summary>
        private PlanilhaExcel.Workbook getWorkbook()
        {
            return generatedFile;
        }

        private void finalizeFiles(){
            generatedFile.Close(SaveChanges: true);
            ReleaseObject(app);
            ReleaseObject(generatedFile);
        }

        #endregion

        #region Salvar Arquivo

        /// <summary>
        /// Salva o arquivo Excel no diretório especificado.
        /// </summary>
        /// <param name="pathToSave"></param>
        /// <param name="fileName"></param>
        /// /// <param name="oldOffice"></param>
        public void saveExcelAs(string pathToSave, string fileName, bool oldOffice = false)
        {
            defineExtension(oldOffice);
            string xlsxName = completeFileName(fileName, extension);
            generatedFile.SaveAs(Filename: pathToSave + xlsxName, FileFormat: PlanilhaExcel.XlFileFormat.xlWorkbookNormal,
                                        AccessMode: PlanilhaExcel.XlSaveAsAccessMode.xlExclusive);
            finalizeFiles();
        }

        #endregion
    }
}