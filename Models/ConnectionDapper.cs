using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dapper;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
namespace InstockAutoMailService.Models
{
    public class ConnectionDapper
    {

        SqlConnection connection; string connConfig = string.Empty;
        public ConnectionDapper()
        {
            connConfig = ConfigurationManager.ConnectionStrings["DBCS"].ConnectionString;

        }

        public DataTable getInstockFullData(string date)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Site", typeof(string));
            dt.Columns.Add("Project", typeof(string));
            dt.Columns.Add("GROUP_LEVEL", typeof(string));
            dt.Columns.Add("MODEL_CODE", typeof(string));
            dt.Columns.Add("QTY", typeof(int));
            List<InstockFull> instockFullCollection = new List<InstockFull>();
            try
            {
                connection = new SqlConnection(connConfig);
                var data = connection.Query<InstockFull>("[dbo].[SP_Instock_AutoMail_Daily]", new {Date = date }, commandType: CommandType.StoredProcedure);
                instockFullCollection = data.ToList();

                foreach (var item in instockFullCollection)
                {
                    DataRow row = dt.NewRow();
                    row["Site"] = item.Site.ToString();
                    row["Project"] = item.Project.ToString();
                    row["GROUP_LEVEL"] = item.Group.ToString();
                    row["MODEL_CODE"] = item.ModelCode.ToString();
                    row["QTY"] = int.Parse(item.QTY.ToString());
                    dt.Rows.Add(row);
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {
                connection.Close();
            }
            return dt;
        }
    }
}
