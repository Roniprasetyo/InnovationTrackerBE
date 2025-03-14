using innovation_tracker_backend.Helper;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.Data;

namespace innovation_tracker_backend.Controllers
{
    [ApiController]
    [Route("api/[controller]/[action]")]
    public class RencanaCircleController(IConfiguration configuration) : Controller
    {
        readonly PolmanAstraLibrary.PolmanAstraLibrary lib = new(configuration.GetConnectionString("DefaultConnection"));
        readonly LDAPAuthentication adAuth = new(configuration);
        DataTable dt = new();

        [Authorize]
        [HttpPost]
        public IActionResult CreateRencanaQCP([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                DataTable dt = lib.CallProcedure("ino_createRencanaQCP", EncodeData.HtmlEncodeObject(value));
                int rciId = Convert.ToInt32(dt.Rows[0]["hasil"]);

                foreach (var member in value["member"])
                {
                    lib.CallProcedure("ino_createMemberDetail", EncodeData.HtmlEncodeObject(new JObject
                    {
                        { "rciId", rciId },
                        { "memNpk", member["memNpk"] },
                        { "memPost", member["memPost"] }
                    }));
                }

                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult UpdateRencanaQCP([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                int rciId = Convert.ToInt32(value["rciId"]);

                dt = lib.CallProcedure("ino_updateRencanaQCP", EncodeData.HtmlEncodeObject(value));

                DataTable res = lib.CallProcedure("ino_getMemberDetailByRencanaCircle", EncodeData.HtmlEncodeObject(new JObject
                {
                    { "rciId", rciId }
                }));

                int pMemberCount = value["member"].ToArray().Length;
                int dbMemberCount = res.Rows.Count;
                bool isRemove = dbMemberCount > pMemberCount;

                if (isRemove)
                {
                    for (int i = 0; i < dbMemberCount; i++)
                    {
                        string npk = res.Rows[i]["Npk"].ToString();
                        string rciID = res.Rows[i]["RciId"].ToString();

                        JObject? matchingMember = value["member"]
                            .FirstOrDefault(m => m["memNpk"].ToString() == npk && rciId.ToString() == rciID) as JObject;


                        if (matchingMember != null)
                        {
                            Console.WriteLine($"Index {i}: update");
                            lib.CallProcedure("ino_updateMemberDetail", EncodeData.HtmlEncodeObject(new JObject
                            {
                                { "rciId", rciId },
                                { "memNpk", matchingMember["memNpk"] },
                                { "memPost", matchingMember["memPost"] }
                            }));
                        }
                        else
                        {
                            Console.WriteLine($"Index {i}: delete");
                            DataTable del = lib.CallProcedure("ino_deleteMemberDetail", EncodeData.HtmlEncodeObject(new JObject
                            {
                                { "rciId", rciId },
                                { "memNpk", npk }
                            }));
                            Console.WriteLine(JsonConvert.SerializeObject(del));
                        }
                    }
                } else
                {
                    for (int i = 0; i < pMemberCount; i++)
                    {
                        string npk = value["member"][i]["memNpk"].ToString();

                        DataRow matchingMember = res.AsEnumerable().FirstOrDefault(m => m["Npk"].ToString() == npk);

                        if (matchingMember != null)
                        {
                            Console.WriteLine($"Index {i}: update");
                            lib.CallProcedure("ino_updateMemberDetail", EncodeData.HtmlEncodeObject(new JObject
                            {
                                { "rciId", rciId },
                                { "memNpk", npk },
                                { "memPost", value["member"][i]["memPost"] }
                            }));
                        }
                        else
                        {
                            Console.WriteLine($"Index {i}: add");
                            lib.CallProcedure("ino_createMemberDetail", EncodeData.HtmlEncodeObject(new JObject
                            {
                                { "rciId", rciId },
                                { "memNpk", npk },
                                { "memPost", value["member"][i]["memPost"] }
                            }));
                        }
                    }
                   
                }

                



                
                //Console.WriteLine(JsonConvert.SerializeObject(res));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch (Exception ex) {
                Console.WriteLine(ex.Message);
                return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetRencanaQCP([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_getRencanaQCP", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetRencanaQCPById([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_getRencanaQCPById", EncodeData.HtmlEncodeObject(value));
                int rciId = Convert.ToInt32(dt.Rows[0]["Key"]);
                DataTable res = lib.CallProcedure("ino_getMemberDetailByRencanaCircle", EncodeData.HtmlEncodeObject(new JObject
                {
                    { "rciId", rciId }
                }));
                JObject result = new JObject();

                if (dt.Rows.Count > 0)
                {
                    foreach (DataColumn col in dt.Columns)
                    {
                        result[col.ColumnName] = JToken.FromObject(dt.Rows[0][col]);
                    }
                }

                JArray members = new JArray();
                foreach (DataRow row in res.Rows)
                {
                    JObject member = new JObject();
                    foreach (DataColumn col in res.Columns)
                    {
                        member[col.ColumnName] = JToken.FromObject(row[col]);
                    }
                    members.Add(member);
                }

                result["member"] = members;
                return Ok(JsonConvert.SerializeObject(result));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetListKaryawan([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_getListKaryawan", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult SentRencanaCircle([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_sentRencanaCircle", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }
    }
}
