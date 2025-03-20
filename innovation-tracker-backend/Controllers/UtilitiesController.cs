﻿using Microsoft.AspNetCore.Authorization;
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
    public class UtilitiesController(IConfiguration configuration) : Controller
    {
        //readonly PolmanAstraLibrary.PolmanAstraLibrary lib = new(PolmanAstraLibrary.PolmanAstraLibrary.Decrypt(configuration.GetConnectionString("DefaultConnection"), "PoliteknikAstra_ConfigurationKey"));
        readonly PolmanAstraLibrary.PolmanAstraLibrary lib = new(configuration.GetConnectionString("DefaultConnection"));
        readonly LDAPAuthentication adAuth = new(configuration);
        DataTable dt = new();

        [HttpPost]
        [SupportedOSPlatform("windows")]
        public IActionResult Login([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                //bool isAuthenticated = adAuth.IsAuthenticated(EncodeData.HtmlEncodeObject(value)[0], EncodeData.HtmlEncodeObject(value)[1]);
                //if (isAuthenticated)
                //{
                    dt = lib.CallProcedure("sso_getAuthenticationInnoTrack", EncodeData.HtmlEncodeObject(value));
                    if (dt.Rows.Count == 0) return Ok(JsonConvert.SerializeObject(new { Status = "LOGIN FAILED" }));
                    return Ok(JsonConvert.SerializeObject(dt));
                //}
                //return Ok(JsonConvert.SerializeObject(new { Status = "LOGIN FAILED" }));
            }
            catch { 
                return BadRequest(); }
        }

        [HttpPost]
        public IActionResult CreateJWTToken([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());

                JWTToken jwtToken = new();
                string token = jwtToken.IssueToken(
                    configuration,
                    EncodeData.HtmlEncodeObject(value)[0],
                    EncodeData.HtmlEncodeObject(value)[1],
                    EncodeData.HtmlEncodeObject(value)[2]
                );

                return Ok(JsonConvert.SerializeObject(new { Token = token }));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult CreateLogLogin([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                lib.CallProcedure("all_createLoginRecord", EncodeData.HtmlEncodeObject(value));
                dt = lib.CallProcedure("all_getLastLogin", [EncodeData.HtmlEncodeObject(value)[0], EncodeData.HtmlEncodeObject(value)[4]]);
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetDataNotifikasi([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("all_getDataNotifikasiReact", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult SetReadNotifikasi([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("all_setReadNotifikasiAllReact", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetDataCountingNotifikasi([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("all_getCountNotifikasiReact", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetListMenu([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                
                dt = lib.CallProcedure("ino_getMenuByRole", EncodeData.HtmlEncodeObject(value));
                Console.WriteLine(dt);
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }
    }
}
