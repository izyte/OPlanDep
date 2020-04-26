using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Services;

using System.IO;

using System.Configuration;
using System.Data;
using System.Text;
using System.Data.OleDb;

using System.Data.OleDb;

using System.Configuration;

using System.Web.Script.Serialization;

using System.IO.Compression;

/// <summary>
/// Summary description for oplanData
/// </summary>
[WebService(Namespace = "http://tempuri.org/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
// To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
[System.Web.Script.Services.ScriptService]

public class PRMSGetPersonnelActivities{
    PRMSGetPersonnelActivities()
    {

    }
    public string start { set; get; }
    public string end { set; get; }
}

public class ActivityAssignment
{
  public ActivityAssignment(String actName,String actStartDate, String actEndDate)
  {
    name = actName;
    startDate = actStartDate;
    endDate = actEndDate;
  }
  public String name { set; get; }
  public String startDate { set; get; }
  public String endDate { set; get; }
}

public class OverStay
{
  public OverStay()
  {

  }
  public Int64 stayDays { set; get; }
  public List<ActivityAssignment> assignments { set; get; }

}

public class Conflict
{
  Conflict()
  {

  }
  public Int64 activityMemberID { set; get; }
  public Int64 personnelID { set; get; }
  public Int64 activityID1 { set; get; }
  public String startDate1 { set; get; }
  public String endDate1 { set; get; }
  public Int64 activityID2 { set; get; }
  public String startDate2 { set; get; }
  public String endDate2 { set; get; }
  public Boolean isContinuation { set; get; }     // end date of the earlier span conincides with the start date of the later span
  public Boolean isContMob { set; get; }          // continuation but mob is on for start date of the later span
  public Boolean isContDemob { set; get; }        // continuation but demob is on for end date of the earlier span
}

public class OplanConfig
{
  public OplanConfig()
  {
    this.id = 0;
    this.type = "";
  }
  public Int64 id { set; get; }
  public String type { set; get; }
  public String setDate { set; get; }
  public Int64 value { set; get; }
  public Double dblValue { set; get; }
  public String startDate { set; get; }
  public String endDate { set; get; }
}



public class opnLoginRights
{
    public opnLoginRights()
    {
        //
        // TODO: Add constructor logic here
        //

        uid = -1;
        login = "visitor";
        role = "read-only";
        fullName = "Visitor";

        //userSession = HTTPContext.Current.Session.SessionID;

        teamAdd = false;
        teamEdit = false;
        teamDelete = false;
        teamMemAdmin = false;

        activityAdd = false;
        activityEdit = false;
        activityDelete = false;
        activityMemAdmin = false;

        personnelAdd = false;
        personnelEdit = false;
        personnelDelete = false;

        settingsAdmin = false;

    }



    public long uid { set; get; }
    public String login { set; get; }
    public String fullName { set; get; }

    public string userSession { set; get; }

    public String role { set; get; }
    public Boolean teamAdd { set; get; }
    public Boolean teamEdit { set; get; }
    public Boolean teamDelete { set; get; }
    public Boolean teamMemAdmin { set; get; }
    public Boolean activityAdd { set; get; }
    public Boolean activityEdit { set; get; }
    public Boolean activityDelete { set; get; }
    public Boolean activityMemAdmin { set; get; }
    public Boolean personnelAdd { set; get; }
    public Boolean personnelEdit { set; get; }
    public Boolean personnelDelete { set; get; }
    public Boolean settingsAdmin { set; get; }

}


public class opnPersonnel
{
    public opnPersonnel()
    {
        //
        // TODO: Add constructor logic here
        //
    }

    public long id { set; get; }
    public string fullName { set; get; }
    public string firstName { set; get; }
    public string lastName { set; get; }
    public string nameGroup { set; get; }
    public string email { set; get; }
    public long companyID { set; get; }
    public long positionID { set; get; }
    public long activityCount { set; get; }


}

public class actValidationParams
{
  public actValidationParams()
  {
    personParams = new List<actMemberValidationParams>() { };
  }

  public String start { set; get; } // activity start
  public String end { set; get; }   // activity end
  public Int32 activityID { set; get; } // activity id
  public List<actMemberValidationParams> personParams { set; get; }  //

}

public class actMemberValidationParams
{
  public actMemberValidationParams()
  {

  }

  public Int32 personID { set; get; } // person id
  public String start { set; get; } // activity start
  public String end { set; get; }   // activity end

}

public class commonData
{
  public commonData()
  {
  }
  public Int32 id { set; get; }
  public String name { set; get; }
  public DateTime actExtracted { set; get; }
  public String actStart { set; get; }
  public String actEnd { set; get; }


}

public class DateScope
{
  public DateScope()
  {
  }
  public string site { set; get; }
  public string startDate { set; get; }
  public string endDate { set; get; }

  public Int32 actId { set; get; }
  public Int32 uid { set; get; }

    public string oldEndDate { set; get; }
  public string oldStartDate { set; get; }

  public string newEndDate { set; get; }
  public string newStartDate { set; get; }

  public string mode { set; get; }
  public bool block { set; get; }

  public Int32 months { set; get; }
}


public class teamData
{
  public teamData()
  {
    this.id = -1;
  }
  public Int32 id { set; get; }
  public String name { set; get; }
  public String description { set; get; }
  public Int32 order { set; get; }
  public Boolean core { set; get; }
  public Boolean upman { set; get; }
  public List<teamMember> members { set; get; }
}

public class teamMember
{
  public teamMember()
  {
  }
  public Int32 id { set; get; }
  public Int32 teamID { set; get; }
  public Int32 perID { set; get; }
  public String shift { set; get; }
  public Int32 subTeamID { set; get; }

}

public class activityData
{
  public activityData()
  {
    this.shiftBy = 0; // if not supplied, this will be the default value which will cause date shifting update to be bypassed
  }
  public Int32 id { set; get; }
  public String name { set; get; }
  public String description { set; get; }
  public String startDate { set; get; }
  public String endDate { set; get; }
  public Int32 type { set; get; }
  public Boolean showInChart { set; get; }
  public Boolean noPOBCount { set; get; }
  public String ready { set; get; }
  public Boolean shiftDates { set; get; }
  public Boolean blockMembers { set; get; }
  public Int32 shiftBy { set; get; }
  public List<activityMember> members { set; get; }
  public String site { set; get; }
  public Boolean vessel { set; get; }
  public Boolean upmanning { set; get; }
  public DateTime extracted { set; get; }
}

public class activityMember
{
  public activityMember()
  {
  }
  public Int32 id { set; get; }
  public Int32 actID { set; get; }
  public Int32 teamID { set; get; }
  public Int32 perID { set; get; }
  public String startDate { set; get; }
  public String endDate { set; get; }
  public Boolean mob { set; get; }
  public Boolean demob { set; get; }
  public Boolean night { set; get; }
  public Boolean isDay { set; get; }
  public Int32 coyID { set; get; }
  public Int32 posID { set; get; }
  public Int32 groupID { set; get; }

  public DateTime actExtracted { set; get; }
  public String actStart { set; get; }
  public String actEnd { set; get; }

}



public class personData
{
    public personData()
    {
        
    }
    public Int32 status { set; get; }
    public Int32 id { set; get; }
    public String fullName { set; get; }
    public String email { set; get; }
    public String companyID { set; get; }
    public String positionID { set; get; }
    public List<personActivity> activities { set; get; }
    public List<personCertificate> certificates { set; get; }
  public String gender { set; get; }
  public bool xl { set; get; }
  public String birthDate { set; get; }
  public String birthDateShort { set; get; }
    
}

public class personCertificate
{
  public personCertificate()
  {
  }
  public Int32 id { set; get; }
  public Boolean req { set; get; }
  public String issue { set; get; }
  public String expiry { set; get; }

}

public class personActivity
{
    public personActivity()
    {
    }
    public Int32 aid { set; get; }
    public String start { set; get; }
    public String end { set; get; }
    public Int32 mob { set; get; }
    public Int32 demob { set; get; }
    public Int32 night { set; get; }

}


public class personnelData
{
    public personnelData()
    {
        personnel = new List<object>() { };
        actLookup = new List<object>() { };
    }

    public List<object> personnel { set; get; }
    public List<object> actLookup { set; get; }

}


public class opnChartData
{
    public opnChartData()
    {
        //
        // TODO: Add constructor logic here
        //
        activities = new List<object> { };
        companies = new List<object> { };
        teams = new List<object> { };
        positions = new List<object> { };
        personnels= new List<object> { };
        maxBeds = new List<object> { };
        upManning = new List<object> { };
    }

    public List<object> activities { set; get; }
    public List<object> companies { set; get; }
    public List<object> teams { set; get; }
    public List<object> positions { set; get; }
    public List<object> personnels { set; get; }
    public List<object> maxBeds { set; get; }
    public List<object> upManning { set; get; }

    public double activitiesMS { set; get; }
    public double companiesMS { set; get; }
    public double teamsMS { set; get; }
    public double positionsMS { set; get; }
    public double personnelsMS { set; get; }
    public double totalMS { set; get; }

}

public class opnTeamsData
{
    public opnTeamsData()
    {
        //
        // TODO: Add constructor logic here
        //
        teams = new List<object> { };
        subTeams = new List<object> { };
    }

    public List<object> teams { set; get; }
    public List<object> subTeams { set; get; }

}


public class opnTest
{
    public opnTest()
    {
        //
        // TODO: Add constructor logic here
        //
    }

    public string name { set; get; }

}


public class oplanData : System.Web.Services.WebService
{

    public oplanData()
    {

        //Uncomment the following line if using designed components 
        //InitializeComponent(); 
    }

  [WebMethod]
  public void getPersonActivityActiveList()
  {
    DataTable tbl = comClsDAL.comReadDataSet("qryActivitiesCountPerMember").Tables[0];
    DataTable tblAct = comClsDAL.comReadDataSet("qryActivitiesActiveMembers").Tables[0];
    DataTable tblActLkp = comClsDAL.comReadDataSet("qryActivitiesActiveLookup").Tables[0];
    
    List<object> actsLkp = new List<object> { };
    List<object> arr = new List<object> { };
    List<object> acts;
    DataTable tblTemp;
    DataView dv;
    Int32 maxDays = 21;
    Int32 maxStay = 0;

    foreach (DataRow r in tblActLkp.Rows)
    {
      actsLkp.Add(new object[]
        {
            r["id"].ToString(),
            r["name"].ToString()
        });
    }

      foreach (DataRow r in tbl.Rows)
    {
      maxStay = 0;
      acts = new List<object>{ };
      
      // filter activities for the current personnel...
      dv = new DataView(tblAct, "pid=" + r["acp_per_id"].ToString(), "", DataViewRowState.CurrentRows);
      tblTemp = dv.ToTable();

      DateTime? curStart = (DateTime?)null;
      DateTime? curEnd = (DateTime?)null;
      DateTime? nextStart = (DateTime?)null;
      DateTime? nextEnd = (DateTime?)null;

      Int32 days = 0;

      foreach (DataRow ar in tblTemp.Rows)
      {


        if (curStart == null)
        {
          days = Convert.ToInt32(ar["dur"]);

          if (days != 0)
          {
            curStart = Convert.ToDateTime(ar["startDate"]);
            curEnd = Convert.ToDateTime(ar["endDate"]);
          }

          maxStay = days;
        }
        else if (Convert.ToInt32(ar["dur"])==0)
        {
          // do not do anything but to get the higher value of maxDays
          if (maxStay < days) maxStay = days;
        }
        else
        {
          nextStart = Convert.ToDateTime(ar["startDate"]);
          nextEnd = Convert.ToDateTime(ar["endDate"]);
          TimeSpan ts = (TimeSpan)(nextStart - curEnd);


          if (curStart <= nextStart && curEnd >= nextEnd)
          {
            // within the previous range, do not count
            // do nothing
          }
          else if (nextStart < curEnd && nextEnd > curEnd)
          {
            // current range start or end is within the previous range
            ts = (TimeSpan)(nextEnd - curEnd);
            days += (ts.Days);
            curEnd = nextEnd;
          }
          else if (ts.Days == 1)
          {
            // continuation of the previous assignment

            days += Convert.ToInt32(ar["dur"]);

          }
          else
          {

            /*if (days > maxDays)
            {
              // record max days
              if (maxStay < days) maxStay = days;
            }*/
            if (maxStay < days) maxStay = days;

            curStart = Convert.ToDateTime(ar["startDate"]);
            curEnd = Convert.ToDateTime(ar["endDate"]);
            days = Convert.ToInt32(ar["dur"]);
            

          }

          /*if (days > maxDays)
          {
            // record max days
            if (maxStay < days) maxStay = days;

          }*/
          if (maxStay < days) maxStay = days;


        }// else

        acts.Add(new object[] {
          ar["aid"].ToString(),
          ar["start"].ToString(),
          ar["end"].ToString(),
          Convert.ToBoolean(ar["mob"]) ? 1 : 0,
          Convert.ToBoolean(ar["demob"]) ? 1 : 0,
          Convert.ToBoolean(ar["night"]) ? 1 : 0,
          Convert.ToBoolean(ar["isDay"]) ? 1 : 0
        });
      }

      arr.Add(new object[]
        {
            r["acp_per_id"].ToString(),
            Cnv32( r["acts"]),
            maxStay,
            acts
        });

      curStart = (DateTime?)null;
      curEnd = (DateTime?)null;

    }

    JavaScriptSerializer js = new JavaScriptSerializer();
    Context.Response.Write("{\"persons\":" + js.Serialize(arr) + ",\"activities\":"+ js.Serialize(actsLkp) + "}");

  }


  [WebMethod]
  public void getPersonActivitiesX(Int32 personID)
  {
    DataTable tbl = comClsDAL.comReadDataSet("qspActivitiesPerMember", new List<object>() { personID }).Tables[0];
    List<object> arr = new List<object> { };


    foreach (DataRow ar in tbl.Rows)
    {
      arr.Add(new object[] {
          ar["aid"].ToString(),
          ar["start"].ToString(),
          ar["end"].ToString(),
          Convert.ToBoolean(ar["mob"]) ? 1 : 0,
          Convert.ToBoolean(ar["demob"]) ? 1 : 0,
          Convert.ToBoolean(ar["night"]) ? 1 : 0
      });
    }


    JavaScriptSerializer js = new JavaScriptSerializer();
    Context.Response.Write("{\"rows\":" + js.Serialize(arr) + "}");
  }

  [WebMethod]
  public void getPersonContPOBX(Int32 personID)
  {
    // crate companies lookup table
    DataTable tbl = comClsDAL.comReadDataSet("qryActivitiesPOBCountPerMember", new List<object>() { personID }).Tables[0];
    List<object> arr = new List<object> { };

    DateTime? curStart = (DateTime?)null;
    DateTime? curEnd = (DateTime?)null;

    DateTime? nextStart = (DateTime?)null;
    DateTime? nextEnd = (DateTime?)null;

    Int32 days = 0;
    Int32 maxDays = 21;

    List<OverStay> OverStays=new List<OverStay>() { };
    OverStay curStay=new OverStay();

    foreach (DataRow r in tbl.Rows)
    {
      if (curStart == null)
      {
        curStart = Convert.ToDateTime(r["startDate"]);
        curEnd = Convert.ToDateTime(r["endDate"]);
        days = Convert.ToInt32(r["dur"]);

        
        curStay = new OverStay()
        {
          stayDays = days,
          assignments = new List<ActivityAssignment>() { }
        };

        curStay.assignments.Add(new ActivityAssignment(r["name"].ToString(), r["startDateFmt"].ToString(), r["endDateFmt"].ToString()));

      }
      else
      {
        nextStart = Convert.ToDateTime(r["startDate"]);
        nextEnd= Convert.ToDateTime(r["endDate"]);
        TimeSpan ts = (TimeSpan)(nextStart - curEnd); ;

        if (curStart<=nextStart && curEnd >= nextEnd)
        {
          // within the previous range, do not count
          // do nothing on date calculation but include activity to assignments
          curStay.assignments.Add(new ActivityAssignment(r["name"].ToString(), r["startDateFmt"].ToString(), r["endDateFmt"].ToString()));
        }
        else if (nextStart < curEnd && nextEnd > curEnd )
        {
          // next start is within the previous range and end next end date is later than the current date
          // add days from the current end date to the next end date

          ts = (TimeSpan)(nextEnd - curEnd);

          days += (ts.Days);

          curEnd = nextEnd;

          curStay.stayDays = days;
          curStay.assignments.Add(new ActivityAssignment(r["name"].ToString(), r["startDateFmt"].ToString(), r["endDateFmt"].ToString()));


        }
        else if (ts.Days==1)
        {
          // continuation of the previous assignment

          days += Convert.ToInt32(r["dur"]);

          curStay.stayDays = days;
          curStay.assignments.Add(new ActivityAssignment(r["name"].ToString(), r["startDateFmt"].ToString(), r["endDateFmt"].ToString()));

        }
        else
        {

          if (days > maxDays)
          {
            OverStays.Add(curStay);
          }

          curStart = Convert.ToDateTime(r["startDate"]);
          curEnd = Convert.ToDateTime(r["endDate"]);
          days = Convert.ToInt32(r["dur"]);

          curStay = new OverStay()
          {
            stayDays = days,
            assignments = new List<ActivityAssignment>() { }
          };

          curStay.assignments.Add(new ActivityAssignment(r["name"].ToString(), r["startDateFmt"].ToString(), r["endDateFmt"].ToString()));

        }

      }
    }


    if (days > maxDays)
    {
      OverStays.Add(curStay);
    }

    JavaScriptSerializer js = new JavaScriptSerializer();

    Context.Response.Write(js.Serialize(OverStays));

  }

  [WebMethod]
  public void getConflict(Int32 activityID)
  {
    // crate companies lookup table
    DataTable tbl = comClsDAL.comReadDataSet("qspConflictPerActivity",new List<object>() { activityID }).Tables[0];
    List<object> arr = new List<object> { }; ;

    foreach (DataRow r in tbl.Rows)
    {
      arr.Add(new object[] {
                r["acid"].ToString(),
                r["pid"].ToString(),
                r["aid1"].ToString(),
                r["tid1"].ToString(),
                r["sd1"].ToString(),
                r["ed1"].ToString(),
                r["aid2"].ToString(),
                r["tid1"].ToString(),
                r["sd2"].ToString(),
                r["ed2"].ToString(),
                CnvBl(r["iscont"]) ? 1 : 0,
                CnvBl(r["iscontm"]) ? 1 : 0,
                CnvBl(r["iscontd"]) ? 1 : 0
            });

    }

    tbl = comClsDAL.comReadDataSet("qspCompetencyExp", new List<object>() { activityID }).Tables[0];

    List<object> comp = new List<object> { }; ;

    foreach (DataRow r in tbl.Rows)
    {
      comp.Add(new object[] {
          r["acp_id"].ToString(),
          r["cmp_per_id"].ToString(),
          r["cmp_crt_id"].ToString(),
          r["issued"].ToString(),
          r["expiry"].ToString(),

          CnvBl(r["notTaken"]) ? 1 : 0,
          CnvBl(r["noExpiry"]) ? 1 : 0,
          CnvBl(r["expired"]) ? 1 : 0,
          CnvBl(r["lastDayToday"]) ? 1 : 0,
          CnvBl(r["exp7D"]) ? 1 : 0,
          CnvBl(r["exp30D"]) ? 1 : 0
      });

    }

    JavaScriptSerializer js = new JavaScriptSerializer();

    Context.Response.Write("{\"rows\":" + js.Serialize(arr) + ", \"comp\":" + js.Serialize(comp) + "}");
    //Context.Response.Write("{\"rows\":" + js.Serialize(arr)  + "}");

  }



  [WebMethod]
    public void getPersonnels()
    {

        string retVal = "";

        DataTable tbl = comClsDAL.comReadDataSet("qryPersonnels").Tables[0];

        List<opnPersonnel> listPersonnels = new List<opnPersonnel>();

        List<string[]> arrayPersonnels = new List<string[]>();

        //string[] arr= new string[];

        foreach (DataRow r in tbl.Rows)
        {
            //opnPersonnel person = new opnPersonnel()
            //{
            //    id = Convert.ToInt32(r["id"]),
            //    fullName = r["fullName"].ToString(),
            //    firstName = r["firstName"].ToString(),
            //    lastName = r["lastName"].ToString(),
            //    nameGroup = r["nameGroup"].ToString(),
            //    email = r["email"].ToString(),
            //    companyID = Convert.ToInt32(r["companyID"]),
            //    positionID= Convert.ToInt32(r["positionID"]),
            //    activityCount= Convert.ToInt32(r["activityCount"])
            //};
            //listPersonnels.Add(person);

            arrayPersonnels.Add(new string[] { r["id"].ToString(), r["fullName"].ToString(),
                    r["nameGroup"].ToString(),r["email"].ToString(),
                    r["companyID"].ToString(),r["positionID"].ToString(),r["activityCount"].ToString()
            });


        }

        JavaScriptSerializer js = new JavaScriptSerializer();

        //retVal = js.Serialize(listPersonnels);
        retVal = js.Serialize(arrayPersonnels);

        Context.Response.Write(retVal);

    }

    [WebMethod]
    public void getPersonnel()
    {

        string retVal = "";

        DataTable tbl = comClsDAL.comReadDataSet("qryPersonnel").Tables[0];
        DataTable tblActivities = comClsDAL.comReadDataSet("qryActivitiesActiveMembers").Tables[0];
        //DataTable tblCert = comClsDAL.comReadDataSet("qryPersonnelPreMobReq").Tables[0];

        DataTable tblTemp;
        Int32 acts;
        Int32 certs;

        personnelData perData = new personnelData();
        DataView dv;

        List<object> listActivities = null;
        //List<object> listCertificates = null;

        foreach (DataRow r in tbl.Rows)
        {

            listActivities = new List<object>() { };
            acts = Convert.ToInt32("0" + r["activityCount"].ToString());

            if (acts > 0)
            {
                dv = new DataView(tblActivities, "pid=" + r["id"].ToString(), "", DataViewRowState.CurrentRows);
                tblTemp = dv.ToTable();
                foreach (DataRow ar in tblTemp.Rows)
                {
                    listActivities.Add(new object[]
                    {

                        ar["aid"].ToString(),
                        ar["start"].ToString(),
                        ar["end"].ToString(),
                        Convert.ToBoolean(ar["mob"]) ? 1 : 0,
                        Convert.ToBoolean(ar["demob"]) ? 1 : 0,
                        Convert.ToBoolean(ar["night"]) ? 1 : 0,
                        Convert.ToBoolean(ar["isDay"]) ? 1 : 0
                    });
                }
            }


            /*listCertificates = new List<object>() { };
            certs = Convert.ToInt32("0" + r["certCount"].ToString());


            if (certs > 0)
            {
                dv = new DataView(tblCert, "per_id=" + r["id"].ToString(), "", DataViewRowState.CurrentRows);
                tblTemp = dv.ToTable();
                foreach (DataRow ar in tblTemp.Rows)
                {
                    listCertificates.Add(new object[]
                    {
              ar["crt_id"].ToString(),
              ar["reqd"].ToString(),
              ar["issued"],
              ar["expiry"],
                    });

                }

            }

            */

            perData.personnel.Add(new object[] {
              r["id"].ToString(),
              r["fullName"].ToString(),
              r["email"].ToString(),
              r["companyID"].ToString(),
              r["positionID"].ToString(),
              acts,
              listActivities,
              r["gender"].ToString(),
              null,
              r["bday"].ToString(),
              r["xl"]
            });

        }

        tbl.Dispose();
        tblActivities.Dispose();

        DataTable tblActivitiesLookup = comClsDAL.comReadDataSet("qryActivitiesActiveLookup").Tables[0];

        foreach (DataRow r in tblActivitiesLookup.Rows)
        {
            perData.actLookup.Add(new object[]
            {
                r["id"].ToString(),
                r["name"].ToString()
            });
        }

        tblActivitiesLookup.Dispose();

        JavaScriptSerializer js = new JavaScriptSerializer();
        retVal = js.Serialize(perData);
        Context.Response.Write(retVal);

    }

    [WebMethod]
    public void savePersonnelx()
    {
        Context.Response.Write("{\"stat\":\"success\",\"action\":\"AddPersonnel\"}");
    }

    public object getDynamicParams()
    {
        JavaScriptSerializer js = new JavaScriptSerializer();

        var jsonString = String.Empty;

        Context.Request.InputStream.Position = 0;
        using (var inputStream = new StreamReader(Context.Request.InputStream))
        {
            jsonString = inputStream.ReadToEnd();
        }

        dynamic data = js.Deserialize(jsonString, typeof(object));

        return data;

    }

    [WebMethod]
    public void getPersonCertificates()
    {

        JavaScriptSerializer js = new JavaScriptSerializer();
        
        dynamic data = getDynamicParams();

        try
        {

            Int32 personID = Cnv32(data["id"]);

            DataTable tbl = comClsDAL.comReadDataSet("qspGetCertsByPersonID",
                new List<object>()
                {
                    personID
                }).Tables[0];

            List<object> certs = new List<object>() { };

            foreach (DataRow r in tbl.Rows)
            {
                certs.Add(new object[] {
                  r["crt_id"].ToString(),
                  r["reqd"].ToString(),
                  r["issued"],
                  r["expiry"],
              });
            }

            // js
            Context.Response.Write("{\"result\":\"success\",\"certs\":"+ js.Serialize(certs) +"}");

        }
        catch (Exception e)
        {
            Context.Response.Write("{\"result\":\"error\",\"errorMessage\":\"" + js.Serialize(e.Message) + "\"}");
        }
    }


    //qspCompetencyExp
    [WebMethod]
    public void getActivityValidation(Int32 activityID)
    {
        String retVal = "";
        ExecCommandsReturn result = new ExecCommandsReturn();
        String operation = "";

        try
        {

            // deserialize personnel data

            DataTable tbl = comClsDAL.comReadDataSet("qspConflictPerActivity",
               new List<object>()
               {
           activityID
               }).Tables[0];

            List<object> conflicts = new List<object>() { };
            //foreach(DataRow r in tbl.Rows)
            //{
            //  conflicts.Add
            //}

        }
        catch (Exception e)
        {

            result.retErr = e;
            operation = "Exception:savePersonnel";

        }

        retVal = parseResult(operation, result);

        JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();

        //Context.Response.Write("{\"rows\":"+ javaScriptSerializer.Serialize() + "}");

    }

    [WebMethod]
    public void deleteActivityMember()
    {
        String retVal = "";
        ExecCommandsReturn result;
        String operation = "DeleteActivityMember";
        DateTime updated = getDateTimeValue(withTime: true);
        Int32 actId = -1;

        try
        {

            var jsonString = String.Empty;

            Context.Request.InputStream.Position = 0;
            using (var inputStream = new StreamReader(Context.Request.InputStream))
            {
                jsonString = inputStream.ReadToEnd();
            }

            // deserialize personnel data
            JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();
            var Data = javaScriptSerializer.Deserialize<activityMember>(jsonString);

            // check if activity was modified by another user/another session
            DataRow refRow = null;
            DataTable tbl = comClsDAL.comReadDataSet("qsp_ActivityMembersDateScope",
                  new List<object>() { Cnv32(Data.actID), stringToDate(Data.actStart), stringToDate(Data.actEnd), Data.actExtracted }).Tables[0];
            if (tbl.Rows.Count != 0) refRow = tbl.Rows[0];
            tbl.Dispose();
            actId = Cnv32(Data.actID);

            if (refRow != null)
            {
                /* This section is bypassed on 2019/12/19 to allow
                 * Deletion of unwanted activity member
                 * 
                if (!CnvBl(refRow["isValid"]))
                {
                    // activity member assignment will be shifted and the new minmum and maximum asignment dates will not fit within the new activity start and end dates
                    result = new ExecCommandsReturn();
                    result.invalidScope = !CnvBl(refRow["inScope"]);
                    result.obsolete = CnvBl(refRow["isObsolete"]);
                    result.shiftDays = Cnv32(refRow["shftDays"]);
                    retVal = parseResult(operation, result);
                    Context.Response.Write(retVal);
                    return;
                }
                */
            }

            // add - [aid] Long, [acid] Long, [tid] Long, [pid] Long, [start] DateTime, [end] DateTime, [mob] Bit, [demob] Bit, [night] Bit
            result = comClsDAL.comExeCommands(new List<commandParam>(){


          new commandParam()
          {
              cmdText= "qspActivityMemberDel",
              cmdParams = new List<object>(){Cnv32(Data.id)}
          },

          new commandParam()
        {
            cmdText= "update tblActivities set act_updated=@p0 where act_id=@p1",
            cmdParams = new List<object>(){
                    updated,
                    actId
            }
        }

        });
        }
        catch (Exception e)
        {

            result = new ExecCommandsReturn();
            result.retErr = e;
            operation = "Exception:" + operation;

        }


        result.retString = "";

        retVal = parseResult(operation, result);
        Context.Response.Write(retVal);
    }


  [WebMethod]
  public void saveActivityMember()
  {
    String retVal = "";
    ExecCommandsReturn result=null;
    String operation = "";
    DateTime updated = getDateTimeValue(withTime: true);
    Int32 actId=-1;

    try
    {

      var jsonString = String.Empty;

      Context.Request.InputStream.Position = 0;
      using (var inputStream = new StreamReader(Context.Request.InputStream))
      {
        jsonString = inputStream.ReadToEnd();
      }

      // deserialize personnel data
      JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();
      var DataArr = javaScriptSerializer.Deserialize<List<activityMember>>(jsonString);

      for(Int16 dIdx=0; dIdx < DataArr.Count; dIdx++) { 
        activityMember Data = DataArr[dIdx];

        DataRow actRow = comClsDAL.comReadDataSet("qsp_ActivityMembersDateScope",
              new List<object>() { Cnv32(Data.actID), stringToDate(Data.actStart),
                                   stringToDate(Data.actEnd), Data.actExtracted }).Tables[0].Rows[0];
        actId = Cnv32(Data.actID);
        actRow.Table.Dispose();
        if (!CnvBl(actRow["isValid"]))
        {
          // activity member assignment will be shifted and the new minmum and maximum asignment dates will not fit within the new activity start and end dates
          result = new ExecCommandsReturn();
          result.invalidScope = !CnvBl(actRow["inScope"]);
          result.obsolete = CnvBl(actRow["isObsolete"]);
          result.shiftDays = Cnv32(actRow["shftDays"]);
          retVal ="[" + parseResult(operation, result) + "]";
          Context.Response.Write(retVal);
          return;
        }


        Boolean isNew = Cnv32(Data.id) == -1;

        operation = isNew ? "AddActivityMember" : "UpdateActivityMember";

        Int32 newId = -1;

        if (isNew)
        {

          DataTable tbl = comClsDAL.comReadDataSet("qryActivityMemberNewID").Tables[0];
          newId = Convert.ToInt32(tbl.Rows[0]["newId"]);

          // add - [aid] Long, [acid] Long, [tid] Long, [pid] Long, [start] DateTime, [end] DateTime, [mob] Bit, [demob] Bit, [night] Bit
          result = comClsDAL.comExeCommands(new List<commandParam>()
          {
            new commandParam()
            {
                cmdText= "qspActivityMemberNew",
                cmdParams = new List<object>(){
                        Cnv32(Data.actID),
                        newId,
                        Cnv32(Data.teamID),
                        Cnv32(Data.perID),
                        stringToDate(Data.startDate),
                        stringToDate(Data.endDate),
                        CnvBl(Data.mob),
                        CnvBl(Data.demob),
                        CnvBl(Data.night),
                        Cnv32(Data.coyID),
                        Cnv32(Data.posID),
                        Cnv32(Data.groupID),
                        CnvBl(Data.isDay),
                }
            }

          });

          result.retInt32 = newId;
          result.activityMemberId = newId;
        }
        else
        {

        // update - [acid] Long, [start] DateTime, [end] DateTime, [mob] Bit, [demob] Bit, [night] Bit;
          result = comClsDAL.comExeCommands(new List<commandParam>()
          {
            new commandParam()
            {
                cmdText= "qspActivityMemberUpdate",
                cmdParams = new List<object>(){
                        Cnv32(Data.id),
                        stringToDate(Data.startDate),
                        stringToDate(Data.endDate),
                        CnvBl(Data.mob),
                        CnvBl(Data.demob),
                        CnvBl(Data.night),
                        Cnv32(Data.coyID),
                        Cnv32(Data.posID),
                        Cnv32(Data.groupID),
                        CnvBl(Data.isDay),
                        Cnv32(Data.perID),
                        Cnv32(Data.teamID),
                }
            }

          });

          result.activityMemberId = Cnv32(Data.id);

        }

        // update activity record with last update date
        comClsDAL.comExeCommands(new List<commandParam>()
        {
          new commandParam()
          {
              cmdText= "update tblActivities set act_updated=@p0 where act_id=@p1",
              cmdParams = new List<object>(){
                      updated,
                      actId
              }
          }

        });

        result.activityId = actId;
        result.retString = "";
        result.lastUpdated = updated;

        retVal += (retVal.Length==0 ? "" : ", ") +  parseResult(operation, result);

      } // end of data supplied data array
    }
    catch (Exception e)
    {

      result = new ExecCommandsReturn();
      result.retErr = e;
      operation = "Exception:saveActivity";

      result.retString = "";
      result.lastUpdated = updated;

      retVal = parseResult(operation, result);

    }


    Context.Response.Write("[" + retVal + "]");
  }

  [WebMethod]
  public void saveTeam()
  {
    String retVal = "";
    ExecCommandsReturn result;
    String operation = "";


    try
    {

      var jsonString = String.Empty;

      Context.Request.InputStream.Position = 0;
      using (var inputStream = new StreamReader(Context.Request.InputStream))
      {
        jsonString = inputStream.ReadToEnd();
      }

      // deserialize personnel data
      JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();
      var Data = javaScriptSerializer.Deserialize<teamData>(jsonString);

      //Boolean isNew = Convert.ToInt32(Data.id.ToString())==-1;
      Boolean isNew = Cnv32(Data.id) == -1;

      operation = isNew ? "AddTeam" : "UpdateTeam";

      Int32 newId = -1;
      Int32 newOrder = -1;

      if (isNew)
      {
        DataTable tbl = comClsDAL.comReadDataSet("qryTeamNewId").Tables[0];
        newId = Convert.ToInt32(tbl.Rows[0]["newId"]);
        newOrder = Convert.ToInt32(tbl.Rows[0]["newOrder"]);
        
      }

      // tid Long, name LongText, [desc] LongText, ord Long, core Bit, um Bit;
      result = comClsDAL.comExeCommands(new List<commandParam>()
            {
                new commandParam()
                {
                    cmdText=isNew ? "qspTeamNew" : "qspTeamUpdate",
                    cmdParams=new List<object>(){
                            isNew ? newId : Cnv32(Data.id),
                            Data.name!=null ? Data.name : "",
                            Data.description !=null ? Data.description : "",
                            isNew ? newOrder : Cnv32(Data.order),
                            CnvBl(Data.core),
                            CnvBl(Data.upman)
                    }
                }

            });

      if (isNew)
      {
        result.retInt32 = newId;
        result.retInt32B = newOrder;
      }

    }
    catch (Exception e)
    {

      result = new ExecCommandsReturn();
      result.retErr = e;
      operation = "Exception:saveTeam";
    }

    result.retString = "";

    retVal = parseResult(operation, result);
    Context.Response.Write(retVal);
  }



  [WebMethod]
  public void saveNewConfig()
  {
    String retVal = "";
    ExecCommandsReturn result;
    String operation = "AddTeamMember";
    String sp = "";

    try
    {

      var jsonString = String.Empty;

      Context.Request.InputStream.Position = 0;
      using (var inputStream = new StreamReader(Context.Request.InputStream))
      {
        jsonString = inputStream.ReadToEnd();
      }

      // deserialize personnel data
      JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();
      var Data = javaScriptSerializer.Deserialize<OplanConfig>(jsonString);
      Int32 newId=-1;

      switch (Data.type)
      {
        case "SPOB":
          newId = getNewId("tblMaximumBeds", "mxb_id");
          sp = "qspCfgNewStdPOBLimit";
          break;
        case "UPOB":
          newId = getNewId("tblUpManning", "upm_id");
          sp = "qspCfgNewUpmPOBLimit";
          break;
        case "UPER":
          newId = getNewId("tblUpManningPeriods", "upmp_id");
          sp = "qspCfgNewUpmPeriod";
          break;
        case "MLIM":
          newId = getNewId("tblMobLimit", "mob_id");
          sp = "qspCfgNewMobLimit";
          break;
        case "DLIM":
          newId = getNewId("tblDemobLimit", "dmb_id");
          sp = "qspCfgNewDemobLimit";
          break;
        default:
          break;
      }
      
      if(Data.type == "UPER")
      {
        result = comClsDAL.comExeCommands(new List<commandParam>()
              {
                  new commandParam()
                  {
                      cmdText=sp,
                      cmdParams=new List<object>(){
                              newId,
                            stringToDate(Data.startDate.ToString()),
                            stringToDate(Data.endDate.ToString()),
                      }
                  }
              });
      }
      else
      { 
        result = comClsDAL.comExeCommands(new List<commandParam>()
              {
                  new commandParam()
                  {
                      cmdText=sp,
                      cmdParams=new List<object>(){
                              newId,
                              stringToDate(Data.setDate.ToString()),
                              Cnv32(Data.value)
                      }
                  }
              });
      }

      result.retInt32 = newId;

    }
    catch (Exception e)
    {

      result = new ExecCommandsReturn();
      result.retInt32 = -1;
      result.retErr = e;
      operation = "Exception:"+sp;

    }

    result.retString = "";
    retVal = parseResult(operation, result);

    Context.Response.Write(retVal);
  }


  [WebMethod]
  public void deleteConfig()
  {
    String retVal = "";
    ExecCommandsReturn result=null;
    String operation = "DeleteConfiguration";
    String sp = "";

    try
    {

      var jsonString = String.Empty;

      Context.Request.InputStream.Position = 0;
      using (var inputStream = new StreamReader(Context.Request.InputStream))
      {
        jsonString = inputStream.ReadToEnd();
      }

      // deserialize personnel data
      JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();
      var Data = javaScriptSerializer.Deserialize<OplanConfig>(jsonString);

      Int32 keyVal = Cnv32(Data.id);

      switch (Data.type)
      {
        case "SPOB":
          result = deleteRecord("tblMaximumBeds", "mxb_id", keyVal);
          break;
        case "UPOB":
          result = deleteRecord("tblUpManning", "upm_id", keyVal);
          break;
        case "UPER":
          result = deleteRecord("tblUpManningPeriods", "upmp_id", keyVal);
          break;
        case "MLIM":
          result = deleteRecord("tblMobLimit", "mob_id", keyVal);
          break;
        case "DLIM":
          result = deleteRecord("tblDemobLimit", "dmb_id", keyVal);
          break;
        default:
          break;
      }

      if (result != null)result.retInt32 = keyVal;

    }
    catch (Exception e)
    {

      result = new ExecCommandsReturn();
      result.retInt32 = -1;
      result.retErr = e;
      operation = "Exception:deleteConfig " ;

    }

    result.retString = "";
    retVal = parseResult(operation, result);

    Context.Response.Write(retVal);
  }



  [WebMethod]
  public void saveTeamMember()
  {
    String retVal = "";
    ExecCommandsReturn result;
    String operation = "AddTeamMember";


    try
    {

      var jsonString = String.Empty;

      Context.Request.InputStream.Position = 0;
      using (var inputStream = new StreamReader(Context.Request.InputStream))
      {
        jsonString = inputStream.ReadToEnd();
      }

      // deserialize personnel data
      JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();
      var Data = javaScriptSerializer.Deserialize<teamMember>(jsonString);


      DataTable tbl = comClsDAL.comReadDataSet("qryTeamMemberNewId").Tables[0];
      Int32 newId = Convert.ToInt32(tbl.Rows[0]["newId"]);

      // [tid] Long, [tmmid] Long, [pid] Long, [shift] LongText, [sbtid] Long;
      result = comClsDAL.comExeCommands(new List<commandParam>()
            {
                new commandParam()
                {
                    cmdText="qspTeamMemberNew",
                    cmdParams=new List<object>(){
                            Cnv32(Data.teamID),
                            newId,
                            Cnv32(Data.perID),
                            Data.shift.ToString(),
                            Cnv32(Data.subTeamID)
                    }
                }

            });

      result.retInt32 = newId;

    }
    catch (Exception e)
    {

      result = new ExecCommandsReturn();
      result.retErr = e;
      operation = "Exception:saveTeam";

    }

    result.retString = "";

    retVal = parseResult(operation, result);
    Context.Response.Write(retVal);
  }

  [WebMethod]
  public void saveActivityCalendarVisibility()
  {
    String retVal = "";
    ExecCommandsReturn result;
    String operation = "";


    try
    {

      var jsonString = String.Empty;

      Context.Request.InputStream.Position = 0;
      using (var inputStream = new StreamReader(Context.Request.InputStream))
      {
        jsonString = inputStream.ReadToEnd();
      }

      // deserialize personnel data
      JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();
      var Data = javaScriptSerializer.Deserialize<activityData>(jsonString);

      operation = "SetActivityCalendarVisibility";

      result = comClsDAL.comExeCommands(new List<commandParam>()
            {
                new commandParam()
                {
                    cmdText="qspActivityInCalendar",
                    cmdParams=new List<object>(){
                            Cnv32(Data.id),
                            CnvBl(Data.showInChart),
                    }
                },

            });


    }
    catch (Exception e)
    {

      result = new ExecCommandsReturn();
      result.retErr = e;
      operation = "Exception:saveActivityCalendarVisibility";

    }

    result.retString = "";

    retVal = parseResult(operation, result);
    Context.Response.Write(retVal);
  }

  [WebMethod]
  public void testZip()
  {


    string sourceFullPath = Context.Request.MapPath("App_Data/oplandb.mdb");
    //string backupSourceFile = sourceFullPath.Replace("oplan.mdb", DateTime.Now.ToString("source/oplandb_yyyyMMdd_HHmmss") + ".mdb");
    

    string startPath = Context.Request.MapPath("App_Data/source");
    string backupSourceFile  = Context.Request.MapPath("App_Data/source/oplandb_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".mdb");
    //string backupSourceFile = startPath + "\\" + DateTime.Now.ToString("\oplandb_yyyyMMdd_HHmmss") + ".mdb";


    DirectoryInfo di = new DirectoryInfo(startPath);
    if (!di.Exists) di.Create();

    FileInfo fi = new FileInfo(backupSourceFile);
    if (fi.Exists) fi.Delete();

    fi = new FileInfo(sourceFullPath);


    System.IO.File.Copy(sourceFullPath, backupSourceFile, true);

    //string zipPath = Context.Request.MapPath("App_Data/backup");

    //System.IO.File.Copy(sourceFullPath, destFile, true);


    Context.Response.Write(sourceFullPath);
    Context.Response.Write("<br/>" + backupSourceFile);
    Context.Response.Write("<br/>" + di.Exists);
    Context.Response.Write("<br/>" + fi.Exists);

    fi = new FileInfo(backupSourceFile);
    Context.Response.Write("<br/>" + fi.Exists);



    //string zipPath = @"D:\Profiles\alv\Documents\SOGA\ng4\oplan\webservice\App_Data\result.zip";
    //string extractPath = @"D:\Profiles\alv\Documents\SOGA\ng4\oplan\webservice\App_Data\extract";

    //ZipFile.CreateFromDirectory(startPath, zipPath);

    //ZipFile.ExtractToDirectory(zipPath, extractPath);
  }

  [WebMethod]
  public void saveActivity()
  {
    String retVal = "";
    ExecCommandsReturn result;
    String operation = "";
    DateTime updated = getDateTimeValue(withTime: true);

    try
    {

      var jsonString = String.Empty;

      Context.Request.InputStream.Position = 0;
      using (var inputStream = new StreamReader(Context.Request.InputStream))
      {
        jsonString = inputStream.ReadToEnd();
      }

      // deserialize personnel data
      JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();
      var Data = javaScriptSerializer.Deserialize<activityData>(jsonString);

      Data.shiftBy = 0; // reset shift by days to 0 because this parameter will be determined during server-side record validation

      //Boolean isNew = Convert.ToInt32(Data.id.ToString())==-1;
      Boolean isNew = Cnv32(Data.id) == -1;

      operation = isNew ? "AddActivity" : "UpdateActivity";

      Int32 newId = -1;
      DataRow refRow = null;
      Int32 offDate = 0;

      if (isNew)
      {
        DataTable tbl = comClsDAL.comReadDataSet("qryActivityNewId").Tables[0];
        newId = Convert.ToInt32(tbl.Rows[0]["newId"]);
      }
      else
      {
        DataTable tbl = comClsDAL.comReadDataSet("qsp_ActivityMembersDateScope",
              new List<object>() { Cnv32(Data.id), stringToDate(Data.startDate), stringToDate(Data.endDate), Data.extracted }).Tables[0];
        if (tbl.Rows.Count != 0) refRow = tbl.Rows[0];
        tbl.Dispose();

        if (refRow != null)
        {
          Data.shiftBy = Cnv32(refRow["shftDays"]); // assign shiftby days during server-side validation
          Data.blockMembers = CnvBl(refRow["isBlockMembers"]); // assign shiftby days during server-side validation
          if (!CnvBl(refRow["isValid"]))
          {
            // activity member assignment will be shifted and the new minmum and maximum asignment dates will not fit within the new activity start and end dates
            result = new ExecCommandsReturn();
            result.invalidScope = !CnvBl(refRow["inScope"]);
            result.obsolete = CnvBl(refRow["isObsolete"]);
            result.blockMembers = CnvBl(refRow["isBlockMembers"]);
            result.shiftDays = Cnv32(refRow["shftDays"]);
            retVal = parseResult(operation, result);
            Context.Response.Write(retVal);
            return;
          }
        }

      }

      List<object> prms = new List<object>()
      {
          isNew ? newId : Cnv32(Data.id),
          Data.name!=null ? Data.name : "",
          Data.description !=null ? Data.description : "",
          stringToDate(Data.startDate),
          stringToDate(Data.endDate),
          Cnv32(Data.type),
          CnvBl(Data.showInChart),
          CnvBl(Data.noPOBCount),
          Data.ready !=null ? Data.ready: "",

          Data.site !=null ? Data.site: "",
          CnvBl(Data.vessel),
          CnvBl(Data.upmanning),
      };

      if (!isNew)
      {
        // add date extracted parameter
        prms.Add(updated);
        prms.Add(Data.extracted);
      }
      else
      {

      }

      result = comClsDAL.comExeCommands(new List<commandParam>()
            {
                new commandParam()
                {
                    cmdText=isNew ? "qspActivityNew" : "qspActivityUpdate",
                    cmdParams=prms
                }

                ,
                new commandParam(Data.shiftBy != 0 && !Data.blockMembers)
                {
                    cmdText = "qspShiftMemberDates",
                    cmdParams=new List<object>(){
                            Cnv32(Data.id),
                            Cnv32(Data.shiftBy)
                    }
                }
                ,
                new commandParam(Data.blockMembers)
                {
                    // update activity member start/end dates to be the same as the activity start/end dates
                    cmdText = "qspMakeMemberDatesSameAsActivity",
                    cmdParams=new List<object>(){
                        Cnv32(Data.id),
                        stringToDate(Data.startDate),
                        stringToDate(Data.endDate),
                    }
                }

            });

      if (Data.shiftBy != 0) result.shiftDays = Data.shiftBy;
      if (Data.blockMembers) result.blockMembers = true;

      if (isNew)
      {
        result.retInt32 = newId;
      }
      else
      {
        if (result.retInt32 == 0)
        {
          // no record has been updated because the activity was modified by another user or it has been removed by another user
          DataTable tbl = comClsDAL.comReadDataSet("select * from tblActivities where act_id=@p1",
              new List<object>() { Cnv32(Data.id) }).Tables[0];

          if (tbl.Rows.Count == 0)
          {
            result.removed = true;
          }
          else
          {
            DateTime lastUpdated = Convert.ToDateTime(tbl.Rows[0]["act_updated"]);
            if (Data.extracted < lastUpdated)  // if extract date is earlier than the last date the record has been updated
            {
              result.obsolete = true;
              result.lastUpdated = lastUpdated;
              result.lastUpdatedBy = tbl.Rows[0]["act_updated_by"].ToString();

            }
          }

        } // if records affected is 0
        else
        {
          result.lastUpdated = updated;
          result.lastUpdatedBy = "";  // to be replaced by the current user... 

          // run activity member assignment date update only if activity update was successful

        }

      } // if updating record


    }
    catch (Exception e)
    {

      result = new ExecCommandsReturn();
      result.retErr = e;
      operation = "Exception:saveActivity";

    }

    result.retString = "";

    retVal = parseResult(operation, result);
    Context.Response.Write(retVal);
  }



  [WebMethod]
  public void saveActivityX20180730()
  {
    String retVal = "";
    ExecCommandsReturn result;
    String operation = "";
    DateTime updated = getDateTimeValue(withTime: true);

    try
    {

      var jsonString = String.Empty;

      Context.Request.InputStream.Position = 0;
      using (var inputStream = new StreamReader(Context.Request.InputStream))
      {
        jsonString = inputStream.ReadToEnd();
      }

      // deserialize personnel data
      JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();
      var Data = javaScriptSerializer.Deserialize<activityData>(jsonString);

      Data.shiftBy = 0; // reset shift by days to 0 because this parameter will be determined during server-side record validation

      //Boolean isNew = Convert.ToInt32(Data.id.ToString())==-1;
      Boolean isNew = Cnv32(Data.id) == -1;

      operation = isNew ? "AddActivity" : "UpdateActivity";

      Int32 newId = -1;
      DataRow refRow = null;
      Int32 offDate = 0;

      if (isNew)
      {
        DataTable tbl = comClsDAL.comReadDataSet("qryActivityNewId").Tables[0];
        newId = Convert.ToInt32(tbl.Rows[0]["newId"]);
      }
      else
      {
        DataTable tbl = comClsDAL.comReadDataSet("qsp_ActivityMembersDateScope",
              new List<object>() { Cnv32(Data.id), stringToDate(Data.startDate), stringToDate(Data.endDate) , Data.extracted }).Tables[0];
        if (tbl.Rows.Count != 0) refRow = tbl.Rows[0];
        tbl.Dispose();

        if (refRow != null)
        {
          Data.shiftBy = Cnv32(refRow["shftDays"]); // assign shiftby days during server-side validation
          if (!CnvBl(refRow["isValid"]))
          {
            // activity member assignment will be shifted and the new minmum and maximum asignment dates will not fit within the new activity start and end dates
            result = new ExecCommandsReturn();
            result.invalidScope = !CnvBl(refRow["inScope"]);
            result.obsolete = CnvBl(refRow["isObsolete"]);
            result.shiftDays = Cnv32(refRow["shftDays"]);
            retVal = parseResult(operation, result);
            Context.Response.Write(retVal);
            return;
          }
        }

      }

      List<object> prms = new List<object>()
      {
          isNew ? newId : Cnv32(Data.id),
          Data.name!=null ? Data.name : "",
          Data.description !=null ? Data.description : "",
          stringToDate(Data.startDate),
          stringToDate(Data.endDate),
          Cnv32(Data.type),
          CnvBl(Data.showInChart),
          CnvBl(Data.noPOBCount),
          Data.ready !=null ? Data.ready: "",

          Data.site !=null ? Data.site: "",
          CnvBl(Data.vessel),
          CnvBl(Data.upmanning),
      };

      if (!isNew)
      {
        // add date extracted parameter
        prms.Add(updated);
        prms.Add(Data.extracted);
      }
      else
      {

      }

      result = comClsDAL.comExeCommands(new List<commandParam>()
            {
                new commandParam()
                {
                    cmdText=isNew ? "qspActivityNew" : "qspActivityUpdate",
                    cmdParams=prms
                }

                ,
                new commandParam(Data.shiftBy != 0)
                {
                    cmdText = "qspShiftMemberDates",
                    cmdParams=new List<object>(){
                            Cnv32(Data.id),
                            Cnv32(Data.shiftBy)
                    }
                }

            });

      if (Data.shiftBy != 0) result.shiftDays = Data.shiftBy;

      if (isNew)
      {
        result.retInt32 = newId;
      }
      else
      {
        if (result.retInt32 == 0)
        {
          // no record has been updated because the activity was modified by another user or it has been removed by another user
          DataTable tbl = comClsDAL.comReadDataSet("select * from tblActivities where act_id=@p1",
              new List<object>() { Cnv32(Data.id) }).Tables[0];

          if (tbl.Rows.Count == 0)
          {
            result.removed = true;
          }
          else
          {
            DateTime lastUpdated = Convert.ToDateTime(tbl.Rows[0]["act_updated"]);
            if (Data.extracted < lastUpdated)  // if extract date is earlier than the last date the record has been updated
            {
              result.obsolete = true;
              result.lastUpdated = lastUpdated;
              result.lastUpdatedBy = tbl.Rows[0]["act_updated_by"].ToString();

            }
          }

        } // if records affected is 0
        else
        {
          result.lastUpdated = updated;
          result.lastUpdatedBy = "";  // to be replaced by the current user... 

          // run activity member assignment date update only if activity update was successful

        }

      } // if updating record


    }
    catch (Exception e)
    {

      result = new ExecCommandsReturn();
      result.retErr = e;
      operation = "Exception:saveActivity";

    }

    result.retString = "";

    retVal = parseResult(operation, result);
    Context.Response.Write(retVal);
  }


    [WebMethod]
    public void savePersonnel()
    {
        String retVal = "";
        ExecCommandsReturn result;
        String operation = "";

        try
        {

            var jsonString = String.Empty;

            Context.Request.InputStream.Position = 0;
            using (var inputStream = new StreamReader(Context.Request.InputStream))
            {
                jsonString = inputStream.ReadToEnd();
            }

            // deserialize personnel data
            JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();
            var Data = javaScriptSerializer.Deserialize<personData>(jsonString);

            //Boolean isNew = Convert.ToInt32(Data.id.ToString())==-1;
            Boolean isNew = Cnv32(Data.id) == -1;

            // Check uniqueness here....
            DataTable tblExisting = comClsDAL.comReadDataSet("qspGetExistingPersonnel",
                   new List<object>() { Data.fullName }).Tables[0];

            if (isNew)
            {
                if(tblExisting.Rows.Count!=0) throw new Exception("Duplicate record found!");
            }
            else
            {
                if (tblExisting.Rows.Count != 0)
                {
                    if (Convert.ToInt32(tblExisting.Rows[0]["per_id"]) != Data.id) throw new Exception("Duplicate record found!");
                }

            }

            operation = isNew ? "AddPersonnel" : "UpdatePersonnel";

            Int32 newId = -1;

            if (isNew)
            {
                DataTable tbl = comClsDAL.comReadDataSet("qryPersonnelNewId").Tables[0];
                newId = Convert.ToInt32(tbl.Rows[0]["newId"]);
            }
            else
            {

                Int32 perId = Cnv32(Data.id);

                // validate and save certificate details
                DataTable certs = comClsDAL.comReadDataSet("select * from tblPersonnelPreMobReq where cmp_per_id=@p0",
                    new List<object>() { perId }).Tables[0];

                // iterate through the certificate records submitted from client
                DataView cdv = new DataView(certs);
                Boolean certExist = false;
                foreach (personCertificate c in Data.certificates)
                {
                    cdv.RowFilter = "cmp_crt_id=" + c.id;
                    certExist = (cdv.Count != 0);

                    object expiry = null;
                    object issue = null;

                    if (c.expiry.Length != 0) expiry = stringToDate(c.expiry);
                    if (c.issue.Length != 0) issue = stringToDate(c.issue);

                    if (expiry == null && certExist && issue == null & !c.req)
                    {
                        // delete existing record
                        comClsDAL.comExeCommands(new List<commandParam>()
                        {
                          new commandParam()
                          {
                            cmdText="delete * from tblPersonnelPreMobReq where cmp_per_id=@p0 And cmp_crt_id=@p1;",
                            cmdParams=new List<object>(){perId, c.id}
                          }
                        });
                    }
                    else if ((expiry != null || issue != null || c.req) && certExist)
                    {
                        // upadate existing record
                        comClsDAL.comExeCommands(new List<commandParam>()
            {
              new commandParam(expiry != null && issue != null)
              {
                cmdText="update tblPersonnelPreMobReq set cmp_required=@p0, cmp_issued=@p1, cmp_expiry=@p2 " +
                  "where cmp_per_id=@p3 and cmp_crt_id=@p4;",
                cmdParams=new List<object>(){c.req, issue, expiry,perId, c.id }
              },
              new commandParam(expiry == null && issue != null)
              {
                cmdText="update tblPersonnelPreMobReq set cmp_required=@p0, cmp_issued=@p1, cmp_expiry=Null " +
                  "where cmp_per_id=@p2 and cmp_crt_id=@p3;",
                cmdParams=new List<object>(){c.req, issue,perId, c.id }
              },
              new commandParam(expiry != null && issue == null)
              {
                cmdText="update tblPersonnelPreMobReq set cmp_required=@p0, cmp_expiry=@p1, cmp_issued=Null " +
                  "where cmp_per_id=@p2 and cmp_crt_id=@p3;",
                cmdParams=new List<object>(){c.req, expiry, perId, c.id }
              },
              new commandParam(expiry == null && issue == null && c.req)
              {
                cmdText="update tblPersonnelPreMobReq set cmp_required=@p0, cmp_expiry=Null, cmp_issued=Null " +
                  "where cmp_per_id=@p1 and cmp_crt_id=@p2;",
                cmdParams=new List<object>(){c.req, perId, c.id }
              }

            });


                    }
                    else if ((expiry != null || issue != null || c.req) && !certExist)
                    {
                        // create new record record

                        comClsDAL.comExeCommands(new List<commandParam>()
            {
              new commandParam(issue!=null && expiry!=null)
              {
                cmdText="insert into tblPersonnelPreMobReq(cmp_per_id, cmp_crt_id, cmp_required, cmp_issued, cmp_expiry) " +
                  "select @p0 as cmp_per_id, @p1 as cmp_crt_id, @p2 as cmp_required, @p3 as cmp_issued, @p4 as cmp_expiry;",
                cmdParams=new List<object>(){perId, c.id, c.req, issue, expiry}
              },
              new commandParam(issue!=null && expiry==null)
              {
                cmdText="insert into tblPersonnelPreMobReq(cmp_per_id, cmp_crt_id, cmp_required, cmp_issued) " +
                  "select @p0 as cmp_per_id, @p1 as cmp_crt_id, @p2 as cmp_required, @p3 as cmp_issued;",
                cmdParams=new List<object>(){perId, c.id, c.req, issue}
              },
              new commandParam(issue==null && expiry!=null)
              {
                cmdText="insert into tblPersonnelPreMobReq(cmp_per_id, cmp_crt_id, cmp_required, cmp_expiry) " +
                  "select @p0 as cmp_per_id, @p1 as cmp_crt_id, @p2 as cmp_required, @p3 as cmp_expiry;",
                cmdParams=new List<object>(){perId, c.id, c.req, expiry}
              },
              new commandParam(issue==null && expiry==null && c.req)
              {
                cmdText="insert into tblPersonnelPreMobReq(cmp_per_id, cmp_crt_id, cmp_required) " +
                  "select @p0 as cmp_per_id, @p1 as cmp_crt_id, @p2 as cmp_required;",
                cmdParams=new List<object>(){perId, c.id, c.req}
              }
            });
                    }
                }

            }

            var nullValue = (object)DBNull.Value;

            result = comClsDAL.comExeCommands(new List<commandParam>()
            {
                new commandParam()
                {
                    cmdText=isNew ? "qspPersonnelNew" : "qspPersonnelUpdate",
                    cmdParams=new List<object>(){
                            isNew ? newId : Cnv32(Data.id),
                            Data.fullName!=null ? Data.fullName : "",
                            Data.email !=null ? Data.email : "",
                            Data.companyID!=null ? Cnv32(Data.companyID) : 0,
                            Data.positionID!=null ? Cnv32(Data.positionID): 0,
                            Data.gender !=null ? Data.gender : "",
                            Data.birthDateShort!="" ? stringToDate(Data.birthDateShort) : nullValue,
                            Data.xl,

                            //stringToDate(Data.birthDateShort),

                    }
                }, new commandParam(isNew)
                {
                    cmdText ="qspAppendDefaultCerts",
                    cmdParams=new List<object>() { newId }
                }

            });

            if (isNew) result.retInt32 = newId;


        }
        catch (Exception e)
        {

            result = new ExecCommandsReturn();
            result.retErr = e;
            operation = "Exception:savePersonnel";

        }



        result.retString = "Test if new";

        retVal = parseResult(operation, result);
        Context.Response.Write(retVal);
        
    }

  public Boolean IsPersonnelExisting(string personnelName)
  {
    Boolean ret = false;

    //Replace(Replace(Replace(Replace(Trim(LTrim(LCase(Replace([per_fullName],","," ")))),"   "," "),"  "," "),".","")," ","_")
    return ret;
  }

  [WebMethod]
  public void deleteActivity()
  {
    String retVal = "";
    ExecCommandsReturn result;
    String operation = "DeleteActivity";

    try
    {

      var jsonString = String.Empty;

      Context.Request.InputStream.Position = 0;
      using (var inputStream = new StreamReader(Context.Request.InputStream))
      {
        jsonString = inputStream.ReadToEnd();
      }

      // deserialize personnel data
      JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();
      var Data = javaScriptSerializer.Deserialize<commonData>(jsonString);


      // check if activity was modified by another user/another session
      DataRow refRow=null;
      DataTable tbl = comClsDAL.comReadDataSet("qsp_ActivityMembersDateScope",
            new List<object>() { Cnv32(Data.id), stringToDate(Data.actStart), stringToDate(Data.actEnd), Data.actExtracted }).Tables[0];
      if (tbl.Rows.Count != 0) refRow = tbl.Rows[0];
      tbl.Dispose();

      if (refRow != null)
      {
        if (!CnvBl(refRow["isValid"]))
        {
          // activity member assignment will be shifted and the new minmum and maximum asignment dates will not fit within the new activity start and end dates
          result = new ExecCommandsReturn();
          result.invalidScope = !CnvBl(refRow["inScope"]);
          result.obsolete = CnvBl(refRow["isObsolete"]);
          result.shiftDays = Cnv32(refRow["shftDays"]);
          retVal = parseResult(operation, result);
          Context.Response.Write(retVal);
          return;
        }
      }


      // add - [tmmid] Long, [person name] string
      result = comClsDAL.comExeCommands(new List<commandParam>(){


          new commandParam()
          {
              cmdText= "qspActivityDel",
              cmdParams = new List<object>(){Cnv32(Data.id)}
          }

        });

      result.retString = Data.name;

    }
    catch (Exception e)
    {

      result = new ExecCommandsReturn();
      result.retErr = e;
      operation = "Exception:" + operation;

    }

    result.retString = "";

    retVal = parseResult(operation, result);
    Context.Response.Write(retVal);
  }


  [WebMethod]
  public void deleteTeam()
  {
    String retVal = "";
    ExecCommandsReturn result;
    String operation = "DeleteTeam";

    try
    {

      var jsonString = String.Empty;

      Context.Request.InputStream.Position = 0;
      using (var inputStream = new StreamReader(Context.Request.InputStream))
      {
        jsonString = inputStream.ReadToEnd();
      }

      // deserialize personnel data
      JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();
      var Data = javaScriptSerializer.Deserialize<commonData>(jsonString);

      // add - [tmmid] Long, [person name] string
      result = comClsDAL.comExeCommands(new List<commandParam>(){


          new commandParam()
          {
              cmdText= "qspTeamDel",
              cmdParams = new List<object>(){Cnv32(Data.id)}
          }

        });

      result.retString = Data.name;

    }
    catch (Exception e)
    {

      result = new ExecCommandsReturn();
      result.retErr = e;
      operation = "Exception:" + operation;

    }

    result.retString = "";

    retVal = parseResult(operation, result);
    Context.Response.Write(retVal);
  }


  [WebMethod]
  public void deleteTeamMember()
  {
    String retVal = "";
    ExecCommandsReturn result;
    String operation = "DeleteTeamMember";

    try
    {

      var jsonString = String.Empty;

      Context.Request.InputStream.Position = 0;
      using (var inputStream = new StreamReader(Context.Request.InputStream))
      {
        jsonString = inputStream.ReadToEnd();
      }

      // deserialize personnel data
      JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();
      var Data = javaScriptSerializer.Deserialize<commonData>(jsonString);

      // add - [tmmid] Long, [person name] string
      result = comClsDAL.comExeCommands(new List<commandParam>(){


          new commandParam()
          {
              cmdText= "qspTeamMemberDel",
              cmdParams = new List<object>(){Cnv32(Data.id)}
          }

        });

      result.retString = Data.name;

    }
    catch (Exception e)
    {

      result = new ExecCommandsReturn();
      result.retErr = e;
      operation = "Exception:" + operation;

    }

    result.retString = "";

    retVal = parseResult(operation, result);
    Context.Response.Write(retVal);
  }


  [WebMethod]
  public void deletePersonnel()
  {
    String retVal = "";
    ExecCommandsReturn result = new ExecCommandsReturn();
    String operation = "";

    try
    {

      var jsonString = String.Empty;

      Context.Request.InputStream.Position = 0;
      using (var inputStream = new StreamReader(Context.Request.InputStream))
      {
        jsonString = inputStream.ReadToEnd();
      }

      // deserialize personnel data
      JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();
      var Data = javaScriptSerializer.Deserialize<commonData>(jsonString);
      Int32 perId = Cnv32(Data.id);

      result = comClsDAL.comExeCommands(new List<commandParam>(){

          new commandParam()
          {
              cmdText= "qspPersonnelDelMain",
              cmdParams = new List<object>(){ perId }
          },

          new commandParam()
          {
            cmdText="delete * from tblPersonnelPreMobReq where cmp_per_id=@p0;",
            cmdParams=new List<object>(){ perId }
          }

    });

      result.retInt32 = Cnv32(Data.id);
      result.retString = Data.name.ToString();

    }
    catch (Exception e)
    {

      result.retErr = e;
      operation = "Exception:savePersonnel";

    }

    retVal = parseResult(operation, result);
    Context.Response.Write(retVal);

  }


  [WebMethod]
  public void deletePersonnelX(Int32 perId)
  {

      var jsonString = String.Empty;

      Context.Request.InputStream.Position = 0;
      using (var inputStream = new StreamReader(Context.Request.InputStream))
      {
          jsonString = inputStream.ReadToEnd();
      }


      ExecCommandsReturn result = comClsDAL.comExeCommands(new List<commandParam>()
          {
              new commandParam()
              {
                  cmdText="qspPersonnelDelMain",
                  cmdParams=new List<object>(){perId}
              }

          });

      String retVal = parseResult("DelPersonnel", result);

      Context.Response.Write(retVal);
  }



    [WebMethod]
    public void getChartActivities(string start, string end,string fltd)
    {
        // start - yyyymmdd (1-based month)
        // end - yyyymmdd (1-based month)

        DateTime timeStart = DateTime.Now;

        string retVal = "";

        List<object[]> arrayActivities = new List<object[]>();
        List<object[]> arrayActivityItem = null;
        opnChartData chartData = new opnChartData();

        object[] tmpArr = null;
        Int32 prvTeam = -1;
        String prvGroup = "";
        Int32 prvPerson = -1;


    // crate activities lookup table
    DataTable tbl = comClsDAL.comReadDataSet(
                            "select * from qspActivitiesByActivityDateRange",
                            new List<object>() {
                                stringToDate(start),
                                stringToDate(end)
                            }
            ).Tables[0];
        
        foreach (DataRow r in tbl.Rows)
        {
            chartData.activities.Add(new object[] {
                r["id"].ToString(),
                r["name"].ToString(),
                r["start"].ToString(),
                r["end"].ToString()
            });
        }

        chartData.activitiesMS = ((TimeSpan)(DateTime.Now - timeStart)).TotalMilliseconds;
        chartData.totalMS += chartData.activitiesMS;
        timeStart = DateTime.Now;

        // crate companies lookup table
        tbl = comClsDAL.comReadDataSet(
                            "select * from qspCompaniesByActivityDateRange",
                            new List<object>() {
                                stringToDate(start),
                                stringToDate(end)
                            }
            ).Tables[0];

        foreach (DataRow r in tbl.Rows)
        {
            chartData.companies.Add(new object[] {
                r["id"].ToString(),
                r["name"].ToString()
            });
        }

        chartData.companiesMS= ((TimeSpan)(DateTime.Now - timeStart)).TotalMilliseconds;
        chartData.totalMS += chartData.companiesMS;
        timeStart = DateTime.Now;

        // crate positions lookup table
        tbl = comClsDAL.comReadDataSet(
                            "select * from qspPositionsByActivityDateRange",
                            new List<object>() {
                                stringToDate(start),
                                stringToDate(end)
                            }
            ).Tables[0];

        foreach (DataRow r in tbl.Rows)
        {
            chartData.positions.Add(new object[] {
                r["id"].ToString(),
                r["name"].ToString()
            });
        }

        chartData.positionsMS= ((TimeSpan)(DateTime.Now - timeStart)).TotalMilliseconds;
        chartData.totalMS += chartData.positionsMS;
        timeStart = DateTime.Now;

        // create teams lookup table
        tbl = comClsDAL.comReadDataSet(
                            "select * from qspTeamsByActivityDateRange",
                            new List<object>() {
                                stringToDate(start),
                                stringToDate(end)
                            }
            ).Tables[0];

        foreach (DataRow r in tbl.Rows)
        {
            chartData.teams.Add(new object[] {
                r["id"].ToString(),
                r["name"].ToString(),
                Convert.ToBoolean(r["isCore"]) ? 1 :0,
                Convert.ToInt32(r["order"]),
                Convert.ToBoolean(r["upman"]) ? 1 :0
            });
    }

        chartData.teamsMS= ((TimeSpan)(DateTime.Now - timeStart)).TotalMilliseconds;
        chartData.totalMS += chartData.teamsMS;
        timeStart = DateTime.Now;


        // create maximum beds and upmanning limits
        tbl = comClsDAL.comReadDataSet("qryMaxBedsAndUpManning").Tables[0];

        foreach (DataRow r in tbl.Rows)
        {
            if (r["src"].ToString() == "1")
            {
                chartData.maxBeds.Add(new object[] {
                    r["dt"].ToString(),
                    Convert.ToInt32(r["limit"])
                });
            }
            else
            {
                chartData.upManning.Add(new object[] {
                    r["dt"].ToString(),
                    Convert.ToInt32(r["limit"])
                });
            }
        }


    // create personnel activities table
    tbl = comClsDAL.comReadDataSet(
                "select * from qspPersonnelActivitiesByDateRange",
                new List<object>() {
                    stringToDate(start),
                    stringToDate(end),
                    (fltd.Length==0 ? Convert.DBNull :  stringToDate(fltd))
                }
        ).Tables[0];


        foreach (DataRow r in tbl.Rows)
        {
          string full = r["per_fullName"].ToString();
          if ((prvTeam != Convert.ToInt32(r["acp_team_id"])) || 
              (prvGroup != r["grp"].ToString()) || 
              (prvPerson != Convert.ToInt32(r["acp_per_id"])) ||
              (full.Substring(0, 3) == "TBA"))
              {

                // append the most recent processed person record
                if (tmpArr != null) chartData.personnels.Add(tmpArr);

                arrayActivityItem = new List<object[]>();   // initialize activities details list

                
                Int64 perId = Convert.ToInt32(r["acp_per_id"]);

        if (full.Substring(0, 3) == "TBA")
        {
          //full += ":" + r["acp_id"].ToString();
          //perId = perId * 1000000 + Convert.ToInt64(r["acp_per_id"]);
        }

                tmpArr = new object[] {
                    perId,
                    full,
                    Convert.ToInt32(r["acp_team_id"]),
                    Convert.ToInt32(r["team_order"]),
                    Convert.ToInt32(r["acp_coy_id"]),
                    Convert.ToInt32(r["acp_pos_id"]),
                    r["grp"].ToString(),
                    arrayActivityItem,
                    Convert.ToInt32(r["acp_act_id"]),
                    r["gender"].ToString(),
                };

        // record the new team/person key
        prvTeam = Convert.ToInt32(r["acp_team_id"]);
                prvGroup = r["grp"].ToString();
                prvPerson = Convert.ToInt32(r["acp_per_id"]);

            }

            arrayActivityItem.Add(new object[]{
                     Convert.ToInt32(r["acp_act_id"]),
                     r["start"].ToString(),
                     r["end"].ToString(),
                     Convert.ToBoolean(r["acp_isMob"]) ? 1 : 0,
                     Convert.ToBoolean(r["acp_isDeMob"]) ? 1 : 0,
                     Convert.ToInt32(r["acp_coy_id"]),
                     Convert.ToInt32(r["acp_pos_id"]),
                     Convert.ToInt32(r["acp_team_id"]),
                     Convert.ToBoolean(r["acp_isNight"]) ? 1 : 0,
                     Convert.ToBoolean(r["act_no_POB_count"]) ? 1 : 0,
                     Convert.ToInt32(r["acp_tmg_id"]),
                     Convert.ToBoolean(r["acp_isDayTrip"]) ? 1 : 0,
                     Convert.ToInt32(r["acp_id"]),
            });

            
        }

        chartData.personnels.Add(tmpArr); // append the last person record

        chartData.personnelsMS= ((TimeSpan)(DateTime.Now - timeStart)).TotalMilliseconds;
        chartData.totalMS += chartData.personnelsMS;
        timeStart = DateTime.Now;

        

        JavaScriptSerializer js = new JavaScriptSerializer();

        retVal = js.Serialize(chartData);
        
        Context.Response.Write(retVal);
    }

    [WebMethod]
    public void getUserInfoX()
    {
      Context.Response.Write("{\"user\":\"me\"}");
    }

    private Boolean inPeriod(DataRow r, Int32 per)
    {
        // set if activity is existing in a period
        Int64 id = Convert.ToInt64(r["id"]);
        if (id == 1020 && (per==3 || per == 2))
        {
            string tmp = "x";
        }

        DateTime actStart = CnD(r["bndSD"]);
        DateTime actEnd = CnD(r["bndED"]);
        DateTime fltStart = CnD(r["startDate"]);
        DateTime fltEnd = CnD(r["endDate"]);

        DateTime periodDate = CnD(r["dtKey" + per]);

        

        //return actStart <= periodDate && actEnd >= periodDate;

        if (per == 1)
        {
            return (actStart >= fltStart && actStart <= periodDate) || (actEnd >= fltStart && actEnd <= periodDate);
        }
        else
        {
            DateTime prevPeriodDate = CnD(r["dtKey" + (per - 1)]);

            //return actStart <= periodDate && actEnd >= periodDate;

            //return (CnD(r["dtKey" + (per - 1)]) <= CnD(r["bndSD"]) && CnD(r["bndSD"]) < CnD(r["dtKey" + per])) ||
            //       (CnD(r["dtKey" + (per - 1)]) < CnD(r["bndED"]) && CnD(r["bndED"]) <= CnD(r["dtKey" + per])) ||
            //       (CnD(r["bndSD"]) < CnD(r["dtKey" + (per - 1)]) && CnD(r["bndED"]) > CnD(r["dtKey" + per]));
        /*              T                   T
            *     6/30/19 <= 9/30/19 && 9/30/19 <= 9/30/19 ||
            *              T                   F
            *     6/30/19 < 11/10/19 && 11/10/19 <= 9/30/19 ||
            *              F                   T
            *     9/20/19 < 6/30/19 && 11/10/19 > 9/20/19
            * 
            */

        return (prevPeriodDate <= actStart && actStart <= periodDate) ||
                (prevPeriodDate < actEnd && actEnd <= periodDate) ||
                (actStart < prevPeriodDate && actEnd > periodDate);


        //return (prevPeriodDate <= actStart && actStart < periodDate) ||
        //       (prevPeriodDate < actEnd && actEnd <= periodDate) ||
        //       (actStart < prevPeriodDate && actEnd > periodDate);

        }
    }
  
  [WebMethod]
  public void getCalendarMTIAPSummary()
  {

    try
    {


      var jsonString = String.Empty;

      Context.Request.InputStream.Position = 0;
      using (var inputStream = new StreamReader(Context.Request.InputStream))
      {
        jsonString = inputStream.ReadToEnd();
      }

      // deserialize personnel data
      JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();
      var Data = javaScriptSerializer.Deserialize<DateScope>(jsonString);

      DateTime sd;
      DateTime ed;
      String site;
      Int32 mos;

      if (Data == null)
      {
        sd = stringToDate(Context.Request.QueryString["startDate"]);
        ed = stringToDate(Context.Request.QueryString["endDate"]);
        site = Context.Request.QueryString["site"];
        mos = Cnv32(Context.Request.QueryString["mos"]);
      }
      else
      {
        sd = stringToDate(Data.startDate);
        ed = stringToDate(Data.endDate);
        site = Data.site;
        mos = Cnv32(Data.months);
      }

      DataTable tblActs = comClsDAL.comReadDataSet("qspCalActivitiesMTIAP",
          new List<object>() {
             sd, ed,site,mos
          }).Tables[0];

      List<object> arrActs = new List<object> { };

      foreach (DataRow r in tblActs.Rows)
      {

        Boolean sem = (CnD(r["dtKey4"]) <= CnD(r["bndSD"]) && CnD(r["bndSD"]) < CnD(r["dtKey5"])) ||
                      (CnD(r["dtKey4"]) <= CnD(r["bndED"]) && CnD(r["bndED"]) < CnD(r["dtKey5"])) ||
                      (CnD(r["bndSD"]) < CnD(r["dtKey4"]) && CnD(r["bndED"]) > CnD(r["dtKey5"]));
        arrActs.Add(new object[]
          {
            r["id"].ToString(),
            r["type"].ToString(),
            r["title"].ToString(),
            r["fmtSD"].ToString(),
            r["fmtED"].ToString(),
            r["ready"].ToString(),
            inPeriod(r,1),
            inPeriod(r,2),
            inPeriod(r,3),
            inPeriod(r,4),
            inPeriod(r,5),
            inPeriod(r,6),
            inPeriod(r,7),
            inPeriod(r,8),
            inPeriod(r,9),
            inPeriod(r,10),
            inPeriod(r,11),
            inPeriod(r,12),
            Cnv32( r["wkSD"]),
            Cnv32( r["wkED"]),
            CnvBl( r["isVessel"]),
            r["fmtBndSD"].ToString(),
            r["fmtBndED"].ToString(),

            inPeriod(r,13),
            inPeriod(r,14),
            inPeriod(r,15),
            inPeriod(r,16),
            inPeriod(r,17),
            inPeriod(r,18),
            inPeriod(r,19),
            inPeriod(r,20),
            inPeriod(r,21),
            inPeriod(r,22),
            inPeriod(r,23),
            inPeriod(r,24),

          }
        );
      }

      DataTable tblTypes = comClsDAL.comReadDataSet("select * from tblActivityType order by iif(atp_priority=999,-1,atp_priority), atp_name").Tables[0];
      List<object> arrTypes = new List<object> { };
      foreach (DataRow r in tblTypes.Rows)
      {
        arrTypes.Add(new object[] {
              Cnv32(r["atp_id"]),
              r["atp_name"].ToString(),
              Cnv32(r["atp_priority"])
          });
      }


      JavaScriptSerializer js = new JavaScriptSerializer();

      Context.Response.Write("{\"ACTS\":" + js.Serialize(arrActs) + ", \"TYPES\":" + js.Serialize(arrTypes) + "}");

    }
    catch (Exception e)
    {
      Context.Response.Write(e.Message);
    }


  }


  [WebMethod]
  public void setActivityDate()
  {

    var jsonString = String.Empty;

    Context.Request.InputStream.Position = 0;
    using (var inputStream = new StreamReader(Context.Request.InputStream))
    {
      jsonString = inputStream.ReadToEnd();
    }

    // deserialize personnel data
    JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();
    var Data = javaScriptSerializer.Deserialize<DateScope>(jsonString);

    DateTime sd;
    DateTime ed;
    string mode;

    Int32 aid=-1;
    DateTime osd = new DateTime();
    DateTime oed = new DateTime();
    DateTime nsd = new DateTime();
    DateTime ned=new DateTime();
    string site="";
    bool block= false;




    if (Data == null)
    {
      sd = stringToDate(Context.Request.QueryString["startDate"]);
      ed = stringToDate(Context.Request.QueryString["endDate"]);
      site = Context.Request.QueryString["site"];

      mode = Context.Request.QueryString["mode"].ToLower();

      //if(Context.Request.QueryString["newStartDate"]!=null) nsd = stringToDate(Context.Request.QueryString["newStartDate"]);
      //if(Context.Request.QueryString["newEndDate"]!=null) ned = stringToDate(Context.Request.QueryString["newEndDate"]);

      //if(Context.Request.QueryString["actId"]!=null) aid = Cnv32(Context.Request.QueryString["actId"]);
      //if(Context.Request.QueryString["site"]!=null) site = Context.Request.QueryString["site"];

      if (mode == "update")
      {
        osd = stringToDate(Context.Request.QueryString["oldStartDate"]);
        oed = stringToDate(Context.Request.QueryString["oldEndDate"]);

        nsd = stringToDate(Context.Request.QueryString["newStartDate"]);
        ned = stringToDate(Context.Request.QueryString["newEndDate"]);

        aid = Cnv32(Context.Request.QueryString["actId"]);

        block = CnvBl(Context.Request.QueryString["block"].ToLower());

      }

    }
    else
    {
      sd = stringToDate(Data.startDate);
      ed = stringToDate(Data.endDate);
      site = Data.site;

      mode = Data.mode.ToLower();

      if (mode == "update")
      {
        osd = stringToDate(Data.oldStartDate);
        oed = stringToDate(Data.oldEndDate);

        nsd = stringToDate(Data.newStartDate);
        ned = stringToDate(Data.newEndDate);

        aid = Data.actId;

        block = Data.block;
      }

    }
    bool withError = false;
    string error = "";
    bool isInitialize = (mode=="initialize" || mode == "cancel");
    bool isUpdate = (mode == "update");
    bool isBlock = block;
    bool isCommit = (mode == "save");

    //isUpdate = false;
    //isCommit = false;

    List<object> prms = new List<object>()
      { Data.uid, sd, ed, site };

    try
    {


      ExecCommandsReturn result = comClsDAL.comExeCommands(new List<commandParam>()
            {
                new commandParam(isInitialize)
                {
                    //cmdText="delete * from tblTmpActivities;"
                    cmdText="qspTmpDeleteActivities",
                    cmdParams = new List<object>() { Data.uid }
                },

                new commandParam(isInitialize)
                {
                    cmdText="qspTmpCalSummaryActs",
                    cmdParams=prms
                },

                new commandParam(isInitialize)
                {
                    cmdText="qspTmpCalSummaryActsMems",
                    cmdParams=prms
                },
                new commandParam(isUpdate)
                {
                  cmdText = "qspTmpActivityNewDates",
                  cmdParams = new List<object>() { Data.uid, aid, nsd, ned}
                },
                new commandParam(isUpdate && !isBlock)
                {
                    // update member dates based on offset from old start date and new start date
                    // applicable to non-block activities
                    cmdText="qspTmpActivityMemberNewDates",
                    cmdParams= new List<object>() { Data.uid, aid, (nsd - osd).TotalDays }
                },

                new commandParam(isUpdate && isBlock)
                {
                    // update member dates based on the new activity start and end dates
                    // applicable to block activities
                    cmdText="qspTmpActivityMemberNewDatesBlock",
                    cmdParams= new List<object>() { Data.uid, aid, nsd, ned }
                },

                new commandParam(isCommit)
                {
                    cmdText="qspTmpActivitySaveDates",
                    cmdParams= new List<object>() { Data.uid }
                },

                new commandParam(isCommit)
                {
                    cmdText="qspTmpActivityMemberSaveDates",
                    cmdParams= new List<object>() { Data.uid }
                },


            });

    }
    catch (Exception e)
    {
      withError = true;
      error = e.Message;
    }


    if (!withError)
    {

      DataTable tbl = comClsDAL.comReadDataSet("qspCalSummaryTemp", prms).Tables[0];

      List<String> arr = new List<String> { };

      foreach (DataRow r in tbl.Rows)
      {
        arr.Add(r["POB"].ToString());
      }

      Context.Response.Write("{\"POB\":" + new JavaScriptSerializer().Serialize(arr) + ", \"success\":true}");

    }
    else
    {
      Context.Response.Write("{\"error\":\"" + error + "\", \"success\":false}");
    }



  }

  [WebMethod]
  public void getCalendarSummary()
  {
    try
    {


      var jsonString = String.Empty;

      Context.Request.InputStream.Position = 0;
      using (var inputStream = new StreamReader(Context.Request.InputStream))
      {
        jsonString = inputStream.ReadToEnd();
      }

      // deserialize personnel data
      JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();
      var Data = javaScriptSerializer.Deserialize<DateScope>(jsonString);

      DateTime sd;
      DateTime ed;
      String site;

      if (Data == null)
      {
        sd = stringToDate(Context.Request.QueryString["startDate"]);
        ed = stringToDate(Context.Request.QueryString["endDate"]);
        site = Context.Request.QueryString["site"];

      }
      else
      {
        sd = stringToDate(Data.startDate);
        ed = stringToDate(Data.endDate);
        site = Data.site;
      }

      DataTable tbl = comClsDAL.comReadDataSet("qspCalSummary",
          new List<object>() {
             sd, ed, site
          }).Tables[0];

      List<String> arr = new List<String> { };

      foreach (DataRow r in tbl.Rows)
      {
        arr.Add(r["POB"].ToString());
      }

      DataTable tblActs = comClsDAL.comReadDataSet("qspCalActivities",
          new List<object>() {
             sd, ed,site
          }).Tables[0];

      List<object> arrActs = new List<object> { };

      foreach (DataRow r in tblActs.Rows)
      {
        arrActs.Add(new object[] 
          {
            r["id"].ToString(),
            r["type"].ToString(),
            r["title"].ToString(),
            r["sd"].ToString(),
            r["ed"].ToString(),
            r["ready"].ToString(),
            CnvBl(r["vessel"]),
            CnvBl(r["is_block"]),
            Cnv32(r["mems"]),
            r["mindt"].ToString(),
            r["maxdt"].ToString(),
          }
        );
      }

      //DataTable tblTypes = comClsDAL.comReadDataSet("select * from tblActivityType order by atp_priority, atp_name").Tables[0];
      DataTable tblTypes = comClsDAL.comReadDataSet("SELECT * FROM tblActivityType ORDER BY IIf([atp_id] = 7,0,1), atp_priority, atp_name").Tables[0];
      //SELECT * FROM tblActivityType ORDER BY IIf([atp_id] = 7,0,1), atp_priority, atp_name;

      List<object> arrTypes = new List<object> { };
      foreach (DataRow r in tblTypes.Rows)
      {
        arrTypes.Add(new object[] {
              Cnv32(r["atp_id"]),
              r["atp_name"].ToString(),
              Cnv32(r["atp_priority"])
          });
      }


      JavaScriptSerializer js = new JavaScriptSerializer();

      Context.Response.Write("{\"POB\":" + js.Serialize(arr) +
                              ", \"ACTS\":" + js.Serialize(arrActs) +
                              ", \"TYPES\":" + js.Serialize(arrTypes) + "}");

    }
    catch (Exception e)
    {
      Context.Response.Write(e.Message);
    }
  }


    [WebMethod]
    public void getPOBProfile()
    {

        String strSD = "";
        String strED = "";

        JavaScriptSerializer js = new JavaScriptSerializer();

        var jsonString = String.Empty;

        Context.Request.InputStream.Position = 0;
        using (var inputStream = new StreamReader(Context.Request.InputStream))
        {
            jsonString = inputStream.ReadToEnd();
        }

        // deserialize personnel data
        var Data = js.Deserialize<DateScope>(jsonString);

        if (Data == null)
        {
            strSD = Context.Request.QueryString["startDate"];
            strED = Context.Request.QueryString["endDate"];
        }
        else
        {
            strSD = Data.startDate;
            strED = Data.endDate;
        }


        String ret = "\"start\":" + strSD + ", \"end\":" + strED + ", ";

        // select statement is used because stored proc call is prompting DataType mismatch error
        DataTable tbl = comClsDAL.comReadDataSet("select * from qspPOBProfile",
            new List<object>()
              {
            stringToDate(strSD),
            stringToDate(strED),
              }
          ).Tables[0];



        List<String> teams = new List<string>() { };
        object POBData = new object();

        int colIdx = 0;
        foreach (DataColumn dc in tbl.Columns)
        {
            String fn = dc.ColumnName;
            if (fn.Substring(0, 2) == "T:")
            {
                // team field
                teams.Add(fn);

                //get data for the team
                List<Int32> dat = new List<Int32>() { };
                foreach (DataRow r in tbl.Rows)
                {
                    dat.Add(Cnv32(r[fn]));
                }

                ret += String.Format("\"d{0}\":", colIdx) + js.Serialize(dat) + ", ";

                colIdx++;
            } // if columnn represents a team...
            else if (fn == "upm_limit")
            {
                List<Int32> upm = new List<Int32>() { };
                foreach (DataRow r in tbl.Rows)
                {
                    upm.Add(Cnv32(r[fn]));
                }
                ret += String.Format("\"upm\":", colIdx) + js.Serialize(upm) + ", ";

            }
            else if (fn == "std_limit")
            {
                List<Int32> std = new List<Int32>() { };
                foreach (DataRow r in tbl.Rows)
                {
                    std.Add(Cnv32(r[fn]));
                }
                ret += String.Format("\"std\":", colIdx) + js.Serialize(std) + ", ";

            }
        }

        tbl.Dispose();


        ret += "\"teams\":" + js.Serialize(teams);
        Context.Response.Write("{" + ret + "}");

    }


  [WebMethod]
  public void test(String sss)
  {
    /*Replace(Replace(Replace(Replace(Trim(LTrim(LCase(Replace([per_fullName], ",", " ")))), "   ", " "), "  ", " "), ".", "")," ","_")*/
    Context.Response.Write(sss + "<br/>");

    Context.Response.Write(encodeStringPattern(sss));
  }

  public String encodeStringPattern(String s)
  {
    if (s == null) s = "";
    if (s.Length!=0)
    {
      return s.Replace("-", " ").Replace("(", " ").Replace(")", " ").Replace(",", " ").Replace(".", " ").ToLower().TrimEnd(' ').TrimStart(' ').Replace("    ", " ").Replace("   ", " ").Replace("  ", " ").Replace(" ", "_");
    }
    else
    {
      return "";
    }
  }


  [WebMethod]
  public void testMSAccesBuiltinFunction()
  {
    try
    {
      DataTable tbl = comClsDAL.comReadDataSet("QueryTemp_A").Tables[0];
      Context.Response.Write("records: " + tbl.Rows.Count);
    }
    catch (Exception e)
    {
      Context.Response.Write(e.Message);
    }
  }


    [WebMethod]
    public void getUserInfo()
    {
        string http = Context.Request.ServerVariables["HTTP_HOST"].ToLower();
        string[] s = Context.Request.ServerVariables["AUTH_USER"].ToString().Split('\\');
        string mark = "start";
        try
        {


            Boolean isDev = (http == "oplan.sogatech.net");

            mark = "before isDev assignment, http:" + http;
            if (!isDev)
            {
                if (http.IndexOf("soga-alv")!=-1)
                {
                    isDev = true;
                }
                else if (http.IndexOf("vm-w2k16") != -1)
                {
                    isDev = true;
                }
                else if (http.IndexOf("localhos") != -1)
                {
                    isDev = true;
                }
                else if (http.IndexOf("oplan.ivideolib.com") != -1)
                {
                    isDev = true;
                }
            }
            mark = "after isDev assignment";


            // crate companies lookup table
            string log = s[s.Length - 1];
            mark = "log";


            
            if (log.Length == 0 && isDev) log = "alv";

            // log = "alv";    // bypass for testing

            mark = "test to log";

            DataTable tbl = comClsDAL.comReadDataSet("qspGetRights",
                new List<object>()
                {
            log
                }
              ).Tables[0];

            opnLoginRights rgh = new opnLoginRights();

            if (tbl.Rows.Count != 0)
            {
                DataRow r = tbl.Rows[0];
                rgh.uid = Cnv32(r["user_id"]);
                rgh.login = r["user_login"].ToString();
                rgh.fullName = r["user_name"].ToString();

                rgh.role = r["role"].ToString();

                rgh.teamAdd = CnvBl(r["team_add"]);
                rgh.teamEdit = CnvBl(r["team_edit"]);
                rgh.teamDelete = CnvBl(r["team_delete"]);
                rgh.teamMemAdmin = CnvBl(r["team_memadmin"]);

                rgh.activityAdd = CnvBl(r["activity_add"]);
                rgh.activityEdit = CnvBl(r["activity_edit"]);
                rgh.activityDelete = CnvBl(r["activity_delete"]);
                rgh.activityMemAdmin = CnvBl(r["activity_memadmin"]);

                rgh.personnelAdd = CnvBl(r["personnel_add"]);
                rgh.personnelEdit = CnvBl(r["personnel_edit"]);
                rgh.personnelDelete = CnvBl(r["personnel_delete"]);

                rgh.settingsAdmin = CnvBl(r["settings_admin"]);

            }
            else
            {
                if (log.Length != 0)
                {
                    rgh.login = log;
                }
            }

            tbl.Dispose();

            JavaScriptSerializer js = new JavaScriptSerializer();

            Context.Response.Write(js.Serialize(rgh));

        }
        catch (Exception e)
        {
            Context.Response.Write("{\"xerror\":\"" + e.Message + ", s:" + s + ", mark:" + mark + "\"}");
        }

    }


    [WebMethod]
    public void getConfig()
    {
        // crate companies lookup table
        String retVal = "";
        DataTable tbl = comClsDAL.comReadDataSet("qryCfgUpManPeriods").Tables[0];
        List<object> arr = new List<object> { };
        JavaScriptSerializer js = new JavaScriptSerializer();

        foreach (DataRow r in tbl.Rows)
        {
            arr.Add(new object[] {
                  r["id"].ToString(),
                  r["start"].ToString(),
                  r["end"].ToString(),
              });

        }


        retVal = "\"upmanPeriods\":" + js.Serialize(arr);

        tbl = comClsDAL.comReadDataSet("qryCfgStdPOBLimits").Tables[0];
        arr = new List<object> { };

        foreach (DataRow r in tbl.Rows)
        {
            arr.Add(new object[] {
                  r["id"].ToString(),
                  r["date"].ToString(),
                  r["limit"].ToString(),
              });

        }

        retVal += ", \"stdPOBLimit\":" + js.Serialize(arr);

        tbl = comClsDAL.comReadDataSet("qryCfgUpmPOBLimits").Tables[0];
        arr = new List<object> { };

        foreach (DataRow r in tbl.Rows)
        {
            arr.Add(new object[] {
                  r["id"].ToString(),
                  r["date"].ToString(),
                  r["limit"].ToString(),
              });

        }

        retVal += ", \"upmPOBLimit\":" + js.Serialize(arr);

        tbl = comClsDAL.comReadDataSet("qryCfgMobLimits").Tables[0];
        arr = new List<object> { };

        foreach (DataRow r in tbl.Rows)
        {
            arr.Add(new object[] {
                  r["id"].ToString(),
                  r["date"].ToString(),
                  r["limit"].ToString(),
              });

        }

        retVal += ", \"mobLimit\":" + js.Serialize(arr);

        tbl = comClsDAL.comReadDataSet("qryCfgDemobLimits").Tables[0];
        arr = new List<object> { };

        foreach (DataRow r in tbl.Rows)
        {
            arr.Add(new object[] {
                  r["id"].ToString(),
                  r["date"].ToString(),
                  r["limit"].ToString(),
              });

        }

        retVal += ", \"demobLimit\":" + js.Serialize(arr);


        tbl = comClsDAL.comReadDataSet("qryNoFlight").Tables[0];
        arr = new List<object> { };

        foreach (DataRow r in tbl.Rows)
        {
            arr.Add(new object[] {
                  r["id"].ToString(),
                  r["date"].ToString(),
                  r["desc"].ToString(),
              });

        }

        retVal += ", \"noflight\":" + js.Serialize(arr);



        tbl = comClsDAL.comReadDataSet("tblTeamGroups", isTable: true).Tables[0];
        arr = new List<object> { };

        foreach (DataRow r in tbl.Rows)
        {
            arr.Add(new object[] {
                  r["tmg_id"].ToString(),
                  r["tmg_name"].ToString(),
              });

        }

        retVal += ", \"teamGroups\":" + js.Serialize(arr);

        tbl = comClsDAL.comReadDataSet("select * from tblCertificates order by crt_order", isTable: true).Tables[0];
        arr = new List<object> { };

        foreach (DataRow r in tbl.Rows)
        {
            arr.Add(new object[] {
                r["crt_id"].ToString(),
                r["crt_name"].ToString(),
                Cnv32(r["crt_frequency"]),
                CnvBl(r["crt_mandatory"]),
                CnvBl(r["crt_default"]),
            });

        }

        retVal += ", \"premobCerts\":" + js.Serialize(arr);



        Context.Response.Write("{" + retVal + "}");

    }


 

    [WebMethod]
    public void getCompanies()
    {
        // crate companies lookup table
        DataTable tbl = comClsDAL.comReadDataSet("qryCompanies").Tables[0];
        List<object> arr= new List<object> { }; ;

        foreach (DataRow r in tbl.Rows)
        {
            arr.Add(new object[] {
                r["id"].ToString(),
                r["name"].ToString()
            });

        }

        JavaScriptSerializer js = new JavaScriptSerializer();

        Context.Response.Write("{\"rows\":"+ js.Serialize(arr) + "}");

    }

    [WebMethod]
    public void getPositions()
    {
        // crate companies lookup table
        DataTable tbl = comClsDAL.comReadDataSet("qryPositions").Tables[0];
        List<object> arr = new List<object> { }; ;

        foreach (DataRow r in tbl.Rows)
        {
            arr.Add(new object[] {
                r["id"].ToString(),
                r["name"].ToString()
            });

        }

        JavaScriptSerializer js = new JavaScriptSerializer();

        Context.Response.Write("{\"rows\":" + js.Serialize(arr) + " }");

        tbl.Dispose();

    }

    [WebMethod]
    public void getTeams()
    {
        // crate companies lookup table
        DataTable tbl = comClsDAL.comReadDataSet("qryTeams").Tables[0];
        DataTable tblSubTeams = comClsDAL.comReadDataSet("qrySubTeams").Tables[0];
        DataTable tblMembers = comClsDAL.comReadDataSet("qryTeamsMembers").Tables[0];
        DataTable tblTemp;

        opnTeamsData teams = new opnTeamsData();
        List<object> members = null;
        DataView dv;

        foreach (DataRow r in tbl.Rows)
        {

            members = new List<object>() { };
            dv = new DataView(tblMembers, "tid=" + r["id"].ToString(),"",DataViewRowState.CurrentRows);
            tblTemp = dv.ToTable("t1");

            //Context.Response.Write("rowcount: " + tblTemp.Rows.Count);

            foreach (DataRow mr in tblTemp.Rows)
            {
                members.Add(new object[]
                {
                    //mr["id"].ToString(),
                    //mr["tid"].ToString(),
                    Cnv32(mr["id"]),
                    Cnv32(mr["pid"]),
                    mr["shift"].ToString(),
                    Cnv32(mr["sid"])
                });
            }

            /*


                id:number;
                name:string;
                description:string;
                core:boolean;
                order:number;
                members:Array<teamMember>;
            }

            export class teamMember{
                perID:number;
                shift:string;
                subTeamID:number;
            }


            */

            teams.teams.Add(new object[] {
                Cnv32(r["id"]),
                r["name"].ToString(),
                r["desc"].ToString(),
                CnvBl(r["core"]) ? 1 : 0,
                Cnv32(r["order"]),
                CnvBl(r["upman"]) ? 1 : 0,
                members
            });

        }

        foreach (DataRow r in tblSubTeams.Rows)
        {
            teams.subTeams.Add(new object[] {
                Cnv32(r["id"]),
                r["name"].ToString()
            });

        }

        JavaScriptSerializer js = new JavaScriptSerializer();

        Context.Response.Write(js.Serialize(teams));

    }

    [WebMethod]
    public void getActivities()
    {
        // crate companies lookup table
        DataTable tbl = comClsDAL.comReadDataSet("qryActivities").Tables[0];
        DataTable tblMembers = comClsDAL.comReadDataSet("qryActivitiesMembers").Tables[0];
        DataTable tblTemp;

        List<object> activities =new List<object>(){ };
        List<object> types =new List<object>(){ };
        List<object> members = null;
        DataView dv;

        foreach (DataRow r in tbl.Rows)
        {

            members = new List<object>() { };
            dv = new DataView(tblMembers, "aid=" + r["id"].ToString(), "", DataViewRowState.CurrentRows);
            tblTemp = dv.ToTable("t1");

            foreach (DataRow mr in tblTemp.Rows)
            {
                members.Add(new object[]
                {
                    mr["id"].ToString(),
                    mr["tid"].ToString(),
                    mr["pid"].ToString(),
                    mr["start"].ToString(),
                    mr["end"].ToString(),
                    Convert.ToBoolean(mr["mob"]) ? "1":"0",
                    Convert.ToBoolean(mr["demob"]) ? "1":"0",
                    Convert.ToBoolean(mr["night"]) ? "1":"0",
                    Convert.ToBoolean(mr["isDay"]) ? "1":"0",
                    mr["coy"].ToString(),
                    mr["pos"].ToString(),
                    mr["grp"].ToString(),
                });
      }

            activities.Add(new object[] {
                r["id"].ToString(),
                r["name"].ToString(),
                r["desc"].ToString(),
                r["start"].ToString(),
                r["end"].ToString(),
                Cnv32(r["type"]),
                Convert.ToBoolean(r["show"]) ? "1":"0",
                Convert.ToBoolean(r["nopob"]) ? "1":"0",
                r["ready"].ToString(),
                members,
                r["site"].ToString(),
                Convert.ToBoolean(r["vessel"]) ? "1":"0",
                Convert.ToBoolean(r["upmanning"]) ? "1":"0",
            });

    }

        DataTable tblTypes = comClsDAL.comReadDataSet("select * from tblActivityType order by atp_priority, atp_name").Tables[0];
        foreach (DataRow r in tblTypes.Rows)
        {
          types.Add(new object[] {
              Cnv32(r["atp_id"]),
              r["atp_name"].ToString(),
              Cnv32(r["atp_priority"])
          });
        }

    tblTypes.Dispose();

    JavaScriptSerializer js = new JavaScriptSerializer();

    DateTime extracted = getDateTimeValue(withTime: true);

    Context.Response.Write("{\"extracted\":" + js.Serialize(extracted) +",\"rows\":" + js.Serialize(activities) + ", \"types\":"+ js.Serialize(types) + "}");

  }

    [WebMethod()]
    public void setNoFlights()
    {
        string yr = Context.Request.QueryString["yr"] + "";
        string mo = Context.Request.QueryString["mo"] + "";
        string days = Context.Request.QueryString["days"] + "";
        if (yr.Length == 0 || mo.Length==0)
        {
            Context.Response.Write("Invalid Parameters:\n\n"+
                    "Syntax: <site>/oplanData.asmx/setNoFlights?yr=<target year>&mo=<month>&days=[d1,d2,d3,...,dn]\n\n"+
                    "yr=<target year>\nmo=<month> (e.g. jan, feb, ... dec or ss for Saturdays and Sundays\n" +
                    "days=[days] (optional comma delimited days. required when calendar month is specified not 'ss')");

            return;
        }
        if (mo == "ss")
        {

        }
        
        Context.Response.Write(Context.Request.QueryString["tmp"] == null);
        Context.Response.Write("Success " + Context.Request.QueryString["yr"] + " " + Context.Request.QueryString["mo"] +" " + Context.Request.QueryString["days"] + " " + Context.Request.QueryString["tmp"]) ;
    }

  [WebMethod()]
  public void getFlightPlanDataNoDownload()
  {

    String strSD = "";
    String strED = "";

    JavaScriptSerializer js = new JavaScriptSerializer();

    var jsonString = String.Empty;

    Context.Request.InputStream.Position = 0;
    using (var inputStream = new StreamReader(Context.Request.InputStream))
    {
      jsonString = inputStream.ReadToEnd();
    }

    // deserialize personnel data
    var Data = js.Deserialize<DateScope>(jsonString);

    if (Data == null)
    {
      strSD = Context.Request.QueryString["startDate"];
      strED = Context.Request.QueryString["endDate"];
    }
    else
    {
      strSD = Data.startDate;
      strED = Data.endDate;
    }

    // generate download file ...

    DataTable tbl = comClsDAL.comReadDataSet("select * from qspCalFlightData",
    new List<object>()
      {
            stringToDate(strSD),
            stringToDate(strED),
      }
    ).Tables[0];

    StringBuilder sb = new StringBuilder();

    foreach (DataColumn col in tbl.Columns)
    {
      sb.Append(col.ColumnName + ',');
    }

    sb.Remove(sb.Length - 1, 1);
    sb.Append(Environment.NewLine);

    foreach (DataRow row in tbl.Rows)
    {
      for (int i = 0; i < tbl.Columns.Count; i++)
      {
        sb.Append("\"" + row[i].ToString() + "\",");
      }

      sb.Append(Environment.NewLine);
    }

    tbl.Dispose();

    if (Context.Request.QueryString["out"] == null)
    {
      String fullPath = Context.Server.MapPath("app_data/FlightPlanData_" + strSD + "_" + strED + ".csv");
      FileInfo fi = new FileInfo(fullPath);
      if (fi.Exists) fi.Delete();

      File.WriteAllText(fullPath, sb.ToString());

      if (Context.Request.QueryString["dl"] != null)
      {
        SendFileToClient(fullPath);
        //Context.Response.Write("{\"start\":" + strSD + ",\"end\":" + strED + "}");
      }
      else
      {
        Context.Response.Write("{\"start\":" + strSD + ",\"end\":" + strED + "}");
      }

      
    }
    else
    {
      Context.Response.Write(sb.ToString());
    }


  }

  [WebMethod()]
  public void downloadFlightPlans()
  {
    SendFileToClient("app_data/FlightPlanData_"+Context.Request.QueryString["startDate"]+"_"+ Context.Request.QueryString["endDate"] + ".csv");
  }

  [WebMethod()]
  public void getFlightPlanData()
  {

    String strSD = "";
    String strED = "";

    JavaScriptSerializer js = new JavaScriptSerializer();

    var jsonString = String.Empty;

    Context.Request.InputStream.Position = 0;
    using (var inputStream = new StreamReader(Context.Request.InputStream))
    {
      jsonString = inputStream.ReadToEnd();
    }

    // deserialize personnel data
    var Data = js.Deserialize<DateScope>(jsonString);

    if (Data == null)
    {
      strSD = Context.Request.QueryString["startDate"];
      strED = Context.Request.QueryString["endDate"];
    }
    else
    {
      strSD = Data.startDate;
      strED = Data.endDate;
    }

    // generate download file ...

    DataTable tbl = comClsDAL.comReadDataSet("select * from qspCalFlightData",
    new List<object>()
      {
            stringToDate(strSD),
            stringToDate(strED),
      }
    ).Tables[0];

    StringBuilder sb = new StringBuilder();

    foreach(DataColumn col in tbl.Columns)
    {
      sb.Append(col.ColumnName + ',');
    }

    sb.Remove(sb.Length - 1, 1);
    sb.Append(Environment.NewLine);

    foreach (DataRow row in tbl.Rows)
    {
      for (int i = 0; i < tbl.Columns.Count; i++)
      {
        sb.Append("\"" + row[i].ToString() + "\"," );
      }

      sb.Append(Environment.NewLine);
    }

    tbl.Dispose();

    String fullPath = Context.Server.MapPath("app_data/FlightPlanData_"+ strSD +"_" + strED +".csv");
    FileInfo fi = new FileInfo(fullPath);
    if (fi.Exists) fi.Delete();

    File.WriteAllText(fullPath, sb.ToString());
  
    SendFileToClient(fullPath);


  }

  public void SendFileToClient(String fullPath,String contentType= "")
  {
    if (contentType == "") contentType = "application/x-msdownload";


    FileInfo fi = new FileInfo(fullPath);

    if(fi.Extension.ToLower()=="pdf") contentType = "application/pdf";

    Context.Response.AddHeader("content-disposition", "attachment; filename=" + fi.Name);
    Context.Response.ContentType = contentType;
    Context.Response.WriteFile(fullPath);
    Context.Response.End();

  }



  [WebMethod()]
  public byte[] DownloadFile(string FName)
  {
    //string fullpath = Context.Server.MapPath(FName);

    System.IO.FileStream fs1 = null;
    fs1 = System.IO.File.Open(FName, FileMode.Open, FileAccess.Read);
    long intFileSize = fs1.Length;
    byte[] b1 = new byte[fs1.Length];
    fs1.Read(b1, 0, (int)fs1.Length);
    fs1.Close();
    //Context.Response.Write(fullpath);

    Context.Response.AddHeader("content-disposition", "attachment; filename=" + FName);
    //Context.Response.ContentType="application/x-msdownload";
    Context.Response.ContentType = "application/pdf";
    //Context.Response.AddHeader("Content-Length",intFileSize.ToString());

    return b1;
  }

  private DateTime stringToDate(string strDate)
    {

        var arr = strDate.Split(' ');
        return new DateTime(
            Convert.ToInt32(strDate.Substring(0, 4)),
            Convert.ToInt32(strDate.Substring(4, 2)),
            Convert.ToInt32(strDate.Substring(6, 2))
        );
        
    }

  private string parseResult(String opn, ExecCommandsReturn res)
  {

    String retVal = "{\"action\":\"" + opn + "\"";

    if (res.retErr == null)
    {
      retVal += ", \"stat\":\"success\"";
      retVal += ", \"message\":\"\"";
    }
    else
    {
      retVal += ", \"stat\":\"fail\"";
      retVal += ", \"message\":\"" + res.retErr.Message + "\"";
    }

    JavaScriptSerializer js = new JavaScriptSerializer();

    retVal += ", \"resString\":\"" + res.retString + "\"";
    retVal += ", \"resInt32\":" + res.retInt32;
    retVal += ", \"resInt32B\":" + res.retInt32B;
    retVal += ", \"resInt32C\":" + res.retInt32C;
    retVal += ", \"resInt32D\":" + res.retInt32D;
    retVal += ", \"activityId\":" + res.activityId;
    retVal += ", \"activityMemberId\":" + res.activityMemberId;
    retVal += ", \"shiftDays\":" + res.shiftDays;
    retVal += ", \"invalidScope\":" + res.invalidScope.ToString().ToLower();
    retVal += ", \"blockMembers\":" + res.blockMembers.ToString().ToLower();
    retVal += ", \"removed\":" + res.removed.ToString().ToLower();
    retVal += ", \"obsolete\":" + res.obsolete.ToString().ToLower();
    retVal += ", \"lastUpdated\":" + js.Serialize(res.lastUpdated);
    retVal += ", \"lastUpdatedBy\":\"" + (res.lastUpdatedBy==null ? "" : res.lastUpdatedBy.ToString()) + "\"";
    retVal += ", \"resObjArr\":" + js.Serialize(res.retObjArr);

    return retVal += "}";
  }

  private ExecCommandsReturn deleteRecord(string tableName, string key, Int32 keyValue)
  {
    ExecCommandsReturn retVal = new ExecCommandsReturn();

   retVal = comClsDAL.comExeCommands(
              new List<commandParam>(){

                new commandParam()
                {
                    cmdText= "delete from " + tableName + " where " + key + " = " + keyValue.ToString() + ";",
                    cmdParams = null
                }
              }
            );




    return retVal;
  }



  private Int32 getNewId(string tableName, string key)
  {
    Int32 ret = 0;
    string sql = "SELECT Max(["+ key + "]+1) AS id FROM "+ tableName  + ";";

    DataTable tbl = comClsDAL.comReadDataSet(sql).Tables[0];
    ret = Cnv32(tbl.Rows[0][0]);
    return ret;
  }

  private DateTime CnD(object dtm)
  {
    DateTime retVal = new DateTime();
    if (dtm != null)
    {
      retVal = Convert.ToDateTime(dtm);
    }
    return retVal;
  }


  private Boolean CnvBl(Object bln)
    {
        Boolean retVal = false;
        if (bln != null)
        {
            retVal = Convert.ToBoolean(bln);
        }
        return retVal;
    }

    private Int32 Cnv32(Object num)
    {
        Int32 retVal = 0;
        if (num != null)
        {
            String numStr = num.ToString();
            if (numStr.Length == 0) numStr = "0";
            retVal = Convert.ToInt32(numStr);
        }

        return retVal;
    }


    private DateTime getDateTimeValue(DateTime? pDate = null, bool withTime = false)
    {

      if (pDate == null) pDate = DateTime.Now;

      DateTime newDate = (DateTime)pDate;

      if (withTime)
      {
        return new DateTime(newDate.Year, newDate.Month, newDate.Day, newDate.Hour, newDate.Minute, newDate.Second);
      }
      else
      {
        return new DateTime(newDate.Year, newDate.Month, newDate.Day);
      }
    }



  private String CnvStr(Object str)
    {
        return str.ToString();
    }

//    private object readInputJSON(object t)
//    {

//    var jsonString = String.Empty;

//    Context.Request.InputStream.Position = 0;
//    using (var inputStream = new StreamReader(Context.Request.InputStream))
//    {
//      jsonString = inputStream.ReadToEnd();
//    }

//    // deserialize personnel data
//    JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();
//    //Object Data = javaScriptSerializer.Deserialize<typeof(t)>(jsonString);
////    JavaScriptSerializer.Deserialize

//    //return Data;

//    }


}
