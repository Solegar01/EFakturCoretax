using EFakturCoretax.Services;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EFakturCoretax.Helpers
{
    public static class QueryHelper
    {
        public static int GetLastDocNum(int seriesId)
        {
            int docNum = 0;
            try
            {
                SAPbobsCOM.Company oCompany = Services.CompanyService.GetCompany();
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string query = $"SELECT NextNumber FROM NNM1 WHERE Series = '{seriesId}'";
                rs.DoQuery(query);
                if (!rs.EoF)
                {
                    docNum = (int)rs.Fields.Item("NextNumber").Value;
                }
               return docNum;
            }
            catch (Exception e)
            {

                throw e;
            }
        }

        public static int GetSeriesIdCoretax()
        {
            try
            {
                Company oCompany = Services.CompanyService.GetCompany();
                Recordset oRecordset = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                // panggil SQL function
                string sql = "SELECT dbo.T2_GET_SERIES_ID_CORETAX() AS SeriesId";
                oRecordset.DoQuery(sql);

                if (!oRecordset.EoF)
                {
                    return Convert.ToInt32(oRecordset.Fields.Item("SeriesId").Value);
                }
                else
                {
                    throw new Exception("SeriesId not found.");
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error GetSeriesIdCoretax2025: {ex.Message}", ex);
            }
        }

        public static string GetSeriesName(int seriesId)
        {
            Company oCompany = Services.CompanyService.GetCompany();
            Recordset oRs = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string seriesName = "";
            try
            {
                string sql = $@"
                SELECT ISNULL(T0.SeriesName, '') AS SeriesName
                FROM NNM1 T0
                WHERE T0.Series = {seriesId}";

                oRs.DoQuery(sql);

                if (!oRs.EoF)
                {
                    seriesName = oRs.Fields.Item("SeriesName").Value.ToString();
                }

                return seriesName;
            }
            catch (Exception ex)
            {
                throw new Exception("Error get series: " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs);
            }
        }

        public static Dictionary<string, string> GetDataCbBranch()
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            try
            {
                Company oCompany = Services.CompanyService.GetCompany();
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery($"SELECT BPLId AS Code, BPLName AS Name FROM OBPL ORDER BY BPLId"); // replace with your query

                while (!rs.EoF)
                {
                    string code = rs.Fields.Item("Code").Value.ToString();
                    string name = rs.Fields.Item("Name").Value.ToString();

                    result.Add(code, name);
                    rs.MoveNext();
                }
                return result;
            }
            catch (Exception)
            {

                throw;
            }
        }

        public static Dictionary<string, string> GetDataCbOutlet()
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            try
            {
                Company oCompany = Services.CompanyService.GetCompany();
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery($"SELECT T0.[PrcCode] AS Code, T0.[PrcName] AS Name FROM OPRC T0 WHERE T0.[DimCode] = '4' AND T0.[Locked] = 'N' ORDER BY T0.[PrcCode] "); // replace with your query

                while (!rs.EoF)
                {
                    string code = rs.Fields.Item("Code").Value.ToString();
                    string name = rs.Fields.Item("Name").Value.ToString();

                    result.Add(code, name);
                    rs.MoveNext();
                }
                return result;
            }
            catch (Exception)
            {

                throw;
            }
        }

        public static Dictionary<string, string> GetDataSeries(SAPbouiCOM.Form oForm, string udoCode)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            try
            {
                Company oCompany = Services.CompanyService.GetCompany();
                // Build SQL depending on FormMode
                string sql = $@"
                            SELECT Series AS Code, SeriesName AS Name
                            FROM NNM1 
                            WHERE ObjectCode = '{udoCode}' ";

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    sql += " AND Locked = 'N' ";
                }

                sql += " ORDER BY Series DESC";

                SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRS.DoQuery(sql);

                while (!oRS.EoF)
                {
                    string code = oRS.Fields.Item("Code").Value.ToString();
                    string name = oRS.Fields.Item("Name").Value.ToString();

                    result.Add(code, name);
                    oRS.MoveNext();
                }
                return result;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
}
