using innovation_tracker_backend.Helper;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.Data;
using System.Net;
using System.Text.RegularExpressions;

namespace innovation_tracker_backend.Controllers
{
    [ApiController]
    [Route("api/[controller]/[action]")]
    public class RencanaSSController(IConfiguration configuration) : Controller
    {
        readonly PolmanAstraLibrary.PolmanAstraLibrary lib = new(configuration.GetConnectionString("DefaultConnection"));
        readonly LDAPAuthentication adAuth = new(configuration);
        DataTable dt = new();

        [Authorize]
        [HttpPost]
        public IActionResult CreateRencanaSS([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                DataTable dt = lib.CallProcedure("ino_createRencanaSS", EncodeData.HtmlEncodeObject(value));

                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult UpdateNilaiSS([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_updateNilaiSS", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult CreateKonvensiSS([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                DataTable dt = lib.CallProcedure("ino_createKonvensiSS", EncodeData.HtmlEncodeObject(value));

                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult UpdateRencanaSS([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());

                dt = lib.CallProcedure("ino_updateRencanaSS", EncodeData.HtmlEncodeObject(value));

                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch
            {
                return BadRequest();
            }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetRencanaSS([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_getRencanaSS", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetListReviewer([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_getListReviewer", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetRencanaSSById([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_getRencanaSSById", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetRencanaSSByIdV2([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_getRencanaSSByIdV2", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult SentRencanaSS([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_sentRencanaSS", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetPenilaian([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_getPenilaian", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetPenilaianById([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_getPenilaianById", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult getAllDataSS([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_getAllRencanaSS", EncodeData.HtmlEncodeObject(value));
                if (dt.Rows.Count > 0)
                {
                    dt.Columns.Remove("Key");
                    dt.Columns.Remove("Count");
                    dt.Columns.Remove("Facil");
                    dt.Columns.Add("Section");
                    dt.Columns.Add("Department");
                    dt.Columns.Add("Submitted On");


                    foreach (DataRow row in dt.Rows)
                    {
                        JObject kry = new JObject
                    {
                        { "username", "-" },
                        { "nama", "-" },
                        { "npk", "-" },
                        { "upt", "-" },
                        { "departmen", "-" }
                    };
                        foreach (var member in value["kryData"])
                        {
                            if (member["username"].ToString().Equals(row["Creaby"].ToString()))
                            {
                                kry = JObject.Parse(member.ToString());
                                break;
                            }
                        }

                        row["NPK"] = kry["npk"].ToString();
                        row["Name"] = kry["nama"];
                        row["Section"] = kry["upt"];
                        row["Department"] = kry["departmen"];
                        row["Submitted On"] = row["Creadate"];
                        row["Project Title"] = Regex.Replace(
                            WebUtility.HtmlDecode(row["Project Title"].ToString()),
                            "<.*?>",
                            string.Empty
                        );
                        row["Score"] = row["Score"].ToString().Equals("") ? 0 : row["Score"];
                    }
                    dt.Columns.Remove("Creadate");
                    dt.Columns.Remove("Creaby");
                    dt.Columns.Remove("Atasan Facil");

                    dt.Columns["No"].SetOrdinal(0);
                    dt.Columns["NPK"].SetOrdinal(1);
                    dt.Columns["Name"].SetOrdinal(2);
                    dt.Columns["Section"].SetOrdinal(3);
                    dt.Columns["Department"].SetOrdinal(4);
                    dt.Columns["Project Title"].SetOrdinal(5);
                    dt.Columns["Category"].SetOrdinal(6);
                    dt.Columns["Start Date"].SetOrdinal(7);
                    dt.Columns["End Date"].SetOrdinal(8);
                    dt.Columns["Period"].SetOrdinal(9);
                    dt.Columns["Batch"].SetOrdinal(10);
                    dt.Columns["Submitted On"].SetOrdinal(11);
                    dt.Columns["Score"].SetOrdinal(12);
                    dt.Columns["Scoring Position"].SetOrdinal(13);
                    dt.Columns["Status"].SetOrdinal(14);


                    return Ok(UtilitiesController.ExportToExcel(dt, "SS_" +
                        "_" +
                        DateTime.Now.Year +
                        DateTime.Now.Month +
                        DateTime.Now.Day +
                        "_" +
                        DateTime.Now.Hour +
                        DateTime.Now.Minute +
                        DateTime.Now.Second + ".xlsx"));
                }
                else
                {
                    return BadRequest();
                }


            }
            catch (Exception ex)
            {
                /*Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);*/
                return
                    BadRequest();
            }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetAllPenilaianById([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_getAllPenilaianById", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetPenilaianByIdForKaProd([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_getPenilaianByIdForKaProd", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetPenilaianByIdForDirorWadir([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_getPenilaianByIdForDirorWadir", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetDetailPenilaian([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_getPenilaianById2", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult SetApproveRencanaSS([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_setApproveRencanaSS", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetListDetailKriteriaPenilaian([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_getListDetailKriteriaPenilaian", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetListSettingRanking([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_getListSettingForRanking", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult SetBatchRencanaSS([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_setBatchRencanaSS", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult UpdateStatusPenilaian([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_updateStatusPenilaian", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult CreatePenilaian([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_createPenilaian2", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetPenilaianByIDScoring([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_getPenilaianByIdScoring", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetPenilaian2([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_getPenilaia2", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetCountSSNeedAction([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_getCountSSNeedAction", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult SetNonActiveRencanaSS([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_UpdateSistemSaranStatus", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetRencanaSSforInnoCoor([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_getRencanaSSforInnoCoor", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetMyRencanaSS([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_getMyRencanaSS", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }
        
        [Authorize]
        [HttpPost]
        public IActionResult GetListStrukturDepartment([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_getListStrukturDepartment", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult SetBatchRencanaSS([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_setBatchRencanaSS", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }
    }
}