using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Globalization;

namespace WebServiceFET
{
    public class Scripts
    {
        private List<SqlParameter> lParam;
        private const int limitToSelect = 7;

        private DataSet getDataSet(string command, List<SqlParameter> lParam)
        {
            var conn = new Connector();
            var adapter = new SqlDataAdapter();
            adapter.SelectCommand = new SqlCommand(command, conn.getSqlConnection());

            // Incluindo parâmetros
            for (int i = 0; i < lParam.Count; i++)
                adapter.SelectCommand.Parameters.Add(lParam[i]);

            var dataSet = new DataSet();
            adapter.Fill(dataSet);
            return dataSet;
        }

        private void includeToListParameter(ref List<SqlParameter> listParameter, string parameterName, object parameterValue)
        {
            listParameter.Add(new SqlParameter(parameterName, parameterValue));
        }

        public DataSet consultaEstabelecimento(int empresaID)
        {
            string command = "SELECT " +
                             "   empresa as estabelecimento, " +
                             "   tipo_documento, " +
                             "   documento, " +
                             "   razao_social, " +
                             "   endereco, " +
                             "   cidade, " +
                             "   estado, " +
                             "   bairro, " +
                             "   telefone, " +
                             "   cod_estabelecimento " +
                             "FROM " +
                             "   cadastro_empresa " +
                             "WHERE " +
                             "   (empresa = @empresaID OR @empresaID = 0) SELECT TOP 1 * FROM usuarios ";

            lParam = new List<SqlParameter>();
            includeToListParameter(ref lParam, "@empresaID", empresaID);

            return getDataSet(command, lParam);
        }

        private bool daysValidator(int days)
        {
            return days <= limitToSelect;
        }

        public DataSet consultaLoja(int empresaID, int lojaID)
        {
            string command = "SELECT DISTINCT cad.[empresa] " +
                             "     ,cad.[loja] " +
                             "     ,cad.[nome_loja] " +
                             "     ,COUNT(pdv.pdv) AS [pdvs] " +
                             "     ,[endereco] " +
                             "     ,[cidade] " +
                             "     ,[estado] " +
                             " FROM cadastro_loja AS cad " +
                             " LEFT JOIN cadastro_pdv AS pdv ON pdv.loja = cad.loja AND pdv.empresa = cad.empresa " +
                             " WHERE " +
                             "   (cad.[empresa] = @empresaID OR @empresaID = 0) " +
                             "   AND (cad.[loja] = @lojaID OR @lojaID = 0) " +
                             " GROUP BY " +
                             "   cad.[empresa] " +
                             "     ,cad.[loja] " +
                             "     ,cad.[nome_loja] " +
                             "     ,[endereco] " +
                             "     ,[cidade] " +
                             "     ,[estado]";

            lParam = new List<SqlParameter>();
            includeToListParameter(ref lParam, "@empresaID", empresaID);
            includeToListParameter(ref lParam, "@lojaID", lojaID);

            return getDataSet(command, lParam);
        }

        public DataSet consultaConexao(int conexaoID)
        {
            string command = "SELECT [nomeconexao], [idconexao], [idmodulorede], [idservidor], [idinterfacerede], [tipoconexao], [comentarios], [portaip], [enderecoip] FROM [conexao] WHERE ([idconexao] = @conexaoID OR @conexaoID = 0)";

            lParam = new List<SqlParameter>();
            SqlParameter param = new SqlParameter("@conexaoID", conexaoID);
            lParam.Add(param);

            return getDataSet(command, lParam);
        }

        public DataSet consultaTerminal(string numeroLogico)
        {
            string command = "select " +
                             "   cl.empresa as estabelecimento, " +
                             "   cl.codterminal as numLogico, " +
                             "   cl.pdv as pdvs, " +
                             "   cl.nsutef as nsu, " +
                             "   cl.statusinicializacao," +
                             "   numeroinicializacoes, " +
                             "   workingkey " +
                             "from " +
                             "   cielo_loja as cl " +
                             "where " +
                             "   codterminal is not null " +
                             "   AND (codterminal = @codTerminal OR @codTerminal = '')";

            lParam = new List<SqlParameter>();
            includeToListParameter(ref lParam, "@codTerminal", numeroLogico);

            return getDataSet(command, lParam);
        }

        public DataSet consultaTransacoes(string numLogico, int daysAgo, bool isPendency)
        {
            if (!daysValidator(daysAgo))
                return null;

            char bit = '0';
            if (isPendency)
                bit = '1';

            string command = " [dbo].[pp_consultaTransacao] '" + numLogico + "', " + daysAgo + ", " + bit;

            lParam = new List<SqlParameter>();
            return getDataSet(command, lParam);
        }

        public DataSet consultaTerminaisEmUso(string numLogico)
        {
            string command = " [dbo].[pp_terminaisEmUso] @numLogico";

            lParam = new List<SqlParameter>();
            includeToListParameter(ref lParam, "@numLogico", numLogico);

            return getDataSet(command, lParam);
        }

        public DataSet consultaOffline(string numLogico)
        {
            string command = "SELECT [empresa] as [estabelecimento], [pdv], [data_hora], [cod_terminal] as [numLogico], [nsu] " +
                                "FROM [cielo_offline] co " +
                                "INNER JOIN acesso_usuario au ON au.idLoja = co.loja AND au.idEmpresa = co.empresa " +
                             "WHERE [status_envio] = 3 " +
                             "AND (cod_terminal = @codTerminal OR @codTerminal = '')";

            lParam = new List<SqlParameter>();
            includeToListParameter(ref lParam, "@codTerminal", numLogico);

            return getDataSet(command, lParam);
        }

        private string dataAtual() {
            DateTime date = DateTime.Now;
            return date.Year + "-" + date.Month + "-" + date.Day;
        }

        public string caminhoParaArquivos(){
            string caminho = "C:\\Users\\Igor\\Desktop\\",
                pastaMes = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(DateTime.Now.Month) + "\\";

            return caminho + pastaMes + dataAtual();
        }

        public void GerarArquivos() {
            Excel excel;
            Directory.CreateDirectory(caminhoParaArquivos());

            excel = new Excel(consultaEstabelecimento(0));
            excel.defaultConfiguration(1, "Estabelecimento");
            excel.saveExcelAs(caminhoParaArquivos() + "\\", "Estabelecimento-" + dataAtual(), true);

            excel = new Excel(consultaTerminal(""));
            excel.defaultConfiguration(1, "Terminal");
            excel.saveExcelAs(caminhoParaArquivos() + "\\", "Terminal-" + dataAtual(), true);

            excel = new Excel(consultaTransacoes("", 1, false));
            excel.defaultConfiguration(1, "Transações");
            excel.addWorksheet(consultaTransacoes("", 1, true));
            excel.defaultConfiguration(1, "Pendentes");
            excel.addWorksheet(consultaOffline(""));
            excel.defaultConfiguration(1, "Offline");
            excel.saveExcelAs(caminhoParaArquivos() + "\\", "Transações-" + dataAtual(), true);
        }
    }
}