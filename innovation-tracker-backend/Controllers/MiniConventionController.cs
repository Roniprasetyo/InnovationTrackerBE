using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using innovation_tracker_backend.Helper;
using System.Data;
using System.Runtime.Versioning;

namespace innovation_tracker_backend.Controllers
{
    [ApiController]
    [Route("api/[controller]/[action]")]
    public class MiniConventionController(IConfiguration configuration) : Controller
    {
        //readonly PolmanAstraLibrary.PolmanAstraLibrary lib = new(PolmanAstraLibrary.PolmanAstraLibrary.Decrypt(configuration.GetConnectionString("DefaultConnection"), "PoliteknikAstra_ConfigurationKey"));
        readonly PolmanAstraLibrary.PolmanAstraLibrary lib = new(configuration.GetConnectionString("DefaultConnection"));
        readonly LDAPAuthentication adAuth = new(configuration);
        DataTable dt = new();

        [Authorize]
        [HttpPost]
        public IActionResult CreateSetting([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString()); 
                dt = lib.CallProcedure("ino_createSetting", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetMiniConvention([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_getMiniConvention", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetMiniConventionById([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_getMiniConventionById", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetSettingById([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_getSettingById", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetListKriteriaPenilaian([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_getListKriteriaPenilaian", EncodeData.HtmlEncodeObject(value));
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
        public IActionResult UpdateSetting([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_updateSetting", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult setStatusSetting([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("ino_setStatusSetting", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }
    }
}
