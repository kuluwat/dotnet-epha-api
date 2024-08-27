using Class;
using Microsoft.AspNetCore.Mvc;
using Model;
using Newtonsoft.Json;

using Microsoft.AspNetCore.Antiforgery;

namespace Controllers
{ 
    [Route("api/[controller]")]
    [ApiController]
    [IgnoreAntiforgeryToken] // ข้ามการตรวจสอบ CSRF
    public class LoginController : ControllerBase
    {
        private readonly IAntiforgery _antiforgery;
        public LoginController(IAntiforgery antiforgery)
        {
            _antiforgery = antiforgery;
        }

        [HttpPost("GetAntiForgeryToken", Name = "GetAntiForgeryToken")]
        public IActionResult GetAntiForgeryToken(tokenModel param)
        {
            string user_name = param.user_name ?? "";
            if (!string.IsNullOrEmpty(user_name))
            {
                var tokens = _antiforgery.GetAndStoreTokens(HttpContext);
                return Ok(new { token = tokens.RequestToken });
            }
            else { return Ok(new { token = "" }); }
        }


        [HttpPost("CheckValidateAntiForgeryToken", Name = "CheckValidateAntiForgeryToken")]
        [ValidateAntiForgeryToken]
        public IActionResult CheckValidateAntiForgeryToken(LoadDocModel param)
        {
            // สมมุติว่า role type ถูกเก็บใน param.user_name
            string roleType = param.user_name ?? "unknown";

            // สร้าง JSON string โดยใช้ JsonConvert
            string jsonString = JsonConvert.SerializeObject(new { roleType });

            // ส่งกลับ JSON string โดยใช้ ContentResult และระบุ Content-Type เป็น application/json
            return Content(jsonString, "application/json");
        }


        [HttpPost("check_authorization_page_fix", Name = "check_authorization_page_fix")]
        //[ValidateAntiForgeryToken]
        public string check_authorization_page_fix(PageRoleListModel param)
        {
            ClassLogin cls = new ClassLogin();
            return cls.check_authorization_page_fix(param);

        }
        [HttpPost("check_authorization_page", Name = "check_authorization_page")]
        //[ValidateAntiForgeryToken]
        public string check_authorization_page(PageRoleListModel param)
        {
            ClassLogin cls = new ClassLogin();
            return cls.authorization_page(param);

        }
        [HttpPost("check_authorization", Name = "check_authorization")]
        //[ValidateAntiForgeryToken]
        public string check_authorization(LoginUserModel param)
        {
            ClassLogin cls = new ClassLogin();
            return cls.login(param);

        }
        [HttpPost("register_account", Name = "register_account")]
        public string register_account(RegisterAccountModel param)
        {
            ClassLogin cls = new ClassLogin();
            return cls.register_account(param);

        }
        [HttpPost("update_register_account", Name = "update_register_account")]
        //[ValidateAntiForgeryToken]
        public string update_register_account(RegisterAccountModel param)
        {
            ClassLogin cls = new ClassLogin();
            return cls.update_register_account(param);

        }
    }
}
