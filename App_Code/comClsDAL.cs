using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;

using System.Configuration;
using System.ComponentModel.DataAnnotations;

/// <summary>
/// Summary description for comClsDAL
/// </summary>
/// 

public class commandParam
{
    public commandParam(Boolean ext=true)
    {
    //
    // TODO: Add constructor logic here
    //
      this.execute = ext;
    }

    public String cmdText { set; get; }
    public IList<object> cmdParams { set; get; }
    public Boolean execute { set; get; }

}

public class ExecCommandsReturn
{
  public ExecCommandsReturn()
  {
    retInt32 = 0;
    retInt32B = 0;
    retInt32C = 0;
    retInt32D = 0;
    activityId = 0;
    activityMemberId = 0;
    shiftDays = 0;
    retString = "";
    retObjArr = new List<object>() { };
    obsolete = false;
    removed = false;
    invalidScope = false;
  }
  public String retString { set; get; }
  public Int32 retInt32 { set; get; }
  public Int32 retInt32B { set; get; }
  public Int32 retInt32C { set; get; }
  public Int32 retInt32D { set; get; }
  public Int32 activityId { set; get; }
  public Int32 activityMemberId { set; get; }
  
  public Int32 shiftDays { set; get; }
  public Boolean invalidScope { set; get; }
  public Boolean blockMembers { set; get; }
  public Boolean obsolete { set; get; }
  public DateTime lastUpdated { set; get; }
  public String lastUpdatedBy { set; get; }
  public Boolean removed { set; get; }
  public List<object> retObjArr { set; get; }
  public Exception retErr { set; get; }
}


public static class GblVar
{
    public const string DB_CONN = "cnsOPlan";
    public const string IMSA_CONN = "OraConnIMSA";
}

public class comClsDAL
{
    static public void comExecuteCommand(string sp, IList<object> prms = null, bool closeConnection=true, string pcnName = "")
    {
        string cnName = (pcnName.Length == 0 ? GblVar.DB_CONN : pcnName);

        string cs = ConfigurationManager.ConnectionStrings[cnName].ConnectionString;

        using (OleDbConnection con = new OleDbConnection(cs))
        {
            OleDbCommand cmd = new OleDbCommand(sp, con);

            con.Open();

            if (prms != null)
            {
                int idx = 0;
                foreach (object o in prms)
                {
                    cmd.Parameters.Add(new OleDbParameter("p" + idx, o));
                    idx++;
                }
            }

            if (sp.TrimStart(' ').ToLower().Substring(0, 7) != "insert " && sp.TrimStart(' ').ToLower().Substring(0, 7) != "delete ")
            {
                cmd.CommandType = CommandType.StoredProcedure;
            }
            else
            {
                cmd.CommandType = CommandType.Text;
            }

            cmd.ExecuteNonQuery();

            if (closeConnection)if(con.State==ConnectionState.Open)con.Close();


        }


    }

    static public ExecCommandsReturn comExeCommands(IList<commandParam> updateParams, String pcnName="")
    {

        OleDbConnection con=null;
        OleDbTransaction xact=null;

        string cnName = (pcnName.Length == 0 ? GblVar.DB_CONN : pcnName);
        string cs = ConfigurationManager.ConnectionStrings[cnName].ConnectionString;

        ExecCommandsReturn retVal = new ExecCommandsReturn();

        Object retObj = null;

        try
        {
            con = new OleDbConnection(cs);
            con.Open();
            xact = con.BeginTransaction();

            foreach (commandParam cp in updateParams)
            {
              if (cp.execute)
              {
                // execute only if required in the calling routine

                retObj = comExeCommand(con, xact, cp);
                if (retObj.GetType().Name != "Int32")
                {
                  Exception e = (Exception)retObj;
                  throw new Exception(e.Message, e);
                }
                else
                {
                  retVal.retInt32 += (Int32)retObj;
                }
              } // execute if required
            } // foreach

            if(xact!=null) xact.Commit();

            if (con != null)
            {
                if (con.State == ConnectionState.Open) con.Close();
            }
        }
        catch (Exception e)
        {
            if (xact!=null)
            {
                xact.Rollback();
            }

            if(con != null)
            {
                if (con.State == ConnectionState.Open) con.Close();
            }

            retVal.retErr = e;
        }
        return retVal;
    }

    static public Object comExeCommand(OleDbConnection con, OleDbTransaction xact, commandParam comObj)
    {

        OleDbCommand cmd;
        Int32 retVal = 0;
        String sql;

        try
        {
            cmd = new OleDbCommand(comObj.cmdText, con, xact);
            sql = comObj.cmdText.TrimStart(' ').ToLower().Substring(0, 7);

            if (comObj.cmdParams != null)
            {
                int idx = 0;
                foreach (object o in comObj.cmdParams)
                {
                    cmd.Parameters.Add(new OleDbParameter("p" + idx, o));
                    idx++;
                }
            }

      

            if (sql != "insert " && sql != "delete " && sql != "update ")
            {
                cmd.CommandType = CommandType.StoredProcedure;
            }
            else
            {
                cmd.CommandType = CommandType.Text;
            }

            retVal = cmd.ExecuteNonQuery();

            return retVal;
        }
        catch (Exception e)
        {
            return e;
        }

    }

    /*
    static public void comExeCommand(List<object[]> comObj)
    {
        string cnName = (pcnName.Length == 0 ? GblVar.DB_CONN : pcnName);

        string cs = ConfigurationManager.ConnectionStrings[cnName].ConnectionString;

        using (OleDbConnection con = new OleDbConnection(cs))
        {
            OleDbCommand cmd = new OleDbCommand(sp, con);

            con.Open();

            if (prms != null)
            {
                int idx = 0;
                foreach (object o in prms)
                {
                    cmd.Parameters.Add(new OleDbParameter("p" + idx, o));
                    idx++;
                }
            }

            if (sp.TrimStart(' ').ToLower().Substring(0, 7) != "select ")
            {
                cmd.CommandType = CommandType.StoredProcedure;
            }
            else
            {
                cmd.CommandType = CommandType.Text;
            }

            cmd.ExecuteNonQuery();

            if (closeConnection) if (con.State == ConnectionState.Open) con.Close();
        }


    }*/

    static public DataSet comReadDataSet(string sp, IList<object> prms = null, string pcnName = "",Boolean isTable = false)
    {

        string cnName = (pcnName.Length == 0 ? GblVar.DB_CONN : pcnName);
        string cs = ConfigurationManager.ConnectionStrings[cnName].ConnectionString;
        DataSet ds = new DataSet();

        using (OleDbConnection con = new OleDbConnection(cs))
        {
            OleDbCommand cmd = new OleDbCommand(sp, con);

            if (prms != null)
            {
                int idx = 0;
                foreach (object o in prms)
                {
                    cmd.Parameters.Add(new OleDbParameter("p" + idx, o));
                    idx++;
                }
            }

            if (sp.TrimStart(' ').ToLower().Substring(0, 7) != "select ")
            {
                cmd.CommandType = isTable ? CommandType.TableDirect :  CommandType.StoredProcedure;
            }
            else
            {
                cmd.CommandType = CommandType.Text;
            }

            OleDbDataAdapter da = new OleDbDataAdapter(cmd);

            int tblCtr = 1;

            da.Fill(ds, "tbl_" + tblCtr);

            con.Close();
        }


        return ds;

    }

    static public string comReadDataToJSON(string sp, Type objType, string[] prm = null)
    {
        // get connection string
        string cs = ConfigurationManager.ConnectionStrings[GblVar.DB_CONN].ConnectionString;
        DataSet ds = new DataSet();
        using (OleDbConnection con = new OleDbConnection(cs))
        {
            OleDbCommand cmd = new OleDbCommand(sp, con);
            if (sp.TrimStart(' ').ToLower().Substring(0, 7) != "select ")
            {
                cmd.CommandType = CommandType.StoredProcedure;
            }
            else
            {
                cmd.CommandType = CommandType.Text;
            }

            OleDbDataAdapter da = new OleDbDataAdapter(cmd);



            int tblCtr = 1;

            da.Fill(ds, "tbl_" + tblCtr);

        }

        return "Hello JSON data!!!";
    }
}


