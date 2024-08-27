using Class;
using dotnet_epha_api.Class;
using Microsoft.AspNetCore.Mvc;
using Model;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    [IgnoreAntiforgeryToken] // ข้ามการตรวจสอบ CSRF
    public class MasterDataController : ControllerBase
    { 
        //manageuser
        [HttpPost("get_manageuser", Name = "get_manageuser")]
        //[ValidateAntiForgeryToken]
        public string get_manageuser(LoadMasterPageModel param)
        {
            ClassMasterData cls = new ClassMasterData();
            return cls.get_manageuser(param);
        }
        [HttpPost("set_manageuser", Name = "set_manageuser")]
        //[ValidateAntiForgeryToken]
        public string set_manageuser(SetManageuser param)
        {
            ClassMasterData cls = new ClassMasterData();
            return cls.set_manageuser(param);
        }

        //authorizationsetting
        [HttpPost("get_authorizationsetting", Name = "get_authorizationsetting")]
        //[ValidateAntiForgeryToken]
        public string get_authorizationsetting(LoadMasterPageModel param)
        {
            ClassMasterData cls = new ClassMasterData();
            return cls.get_authorizationsetting(param);
        }
        [HttpPost("set_authorizationsetting", Name = "set_authorizationsetting")]
        //[ValidateAntiForgeryToken]
        public string set_authorizationsetting(SetAuthorizationSetting param)
        {
            ClassMasterData cls = new ClassMasterData();
            return cls.set_authorizationsetting(param);
        }

        #region master systemwide

        [HttpPost("get_master_company", Name = "get_master_company")]
        //[ValidateAntiForgeryToken]
        public string get_master_company(LoadMasterPageModel param)
        {
            ClassMasterData cls = new ClassMasterData();
            return cls.get_master_company(param);
        }

        //area
        [HttpPost("get_master_area", Name = "get_master_area")]
        //[ValidateAntiForgeryToken]
        public string get_master_area(LoadMasterPageModel param)
        {
            ClassMasterData cls = new ClassMasterData();
            return cls.get_master_area(param);
        }
        [HttpPost("set_master_area", Name = "set_master_area")]
        //[ValidateAntiForgeryToken]
        public string set_master_area(SetDataMasterModel param)
        {
            param.page_name = "area";
            return _set_master_data(param);
        }

        //toc
        [HttpPost("get_master_toc", Name = "get_master_toc")]
        //[ValidateAntiForgeryToken]
        public string get_master_toc(LoadMasterPageModel param)
        {
            ClassMasterData cls = new ClassMasterData();
            return cls.get_master_toc(param);
        }
        [HttpPost("set_master_toc", Name = "set_master_toc")]
        //[ValidateAntiForgeryToken]
        public string set_master_toc(SetDataMasterModel param)
        {
            param.page_name = "toc";
            return _set_master_data(param);
        }

        //unit
        [HttpPost("get_master_unit", Name = "get_master_unit")]
        //[ValidateAntiForgeryToken]
        public string get_master_unit(LoadMasterPageModel param)
        {
            ClassMasterData cls = new ClassMasterData();
            return cls.get_master_unit(param);
        }
        [HttpPost("set_master_unit", Name = "set_master_unit")]
        //[ValidateAntiForgeryToken]
        public string set_master_unit(SetDataMasterModel param)
        {
            param.page_name = "unit";
            return _set_master_data(param);
        }
        private string _set_master_data(SetDataMasterModel param)
        {
            string msg = "";
            try
            {
                ClassMasterData cls = new ClassMasterData();
                return cls.set_master_systemwide(param);
            }
            catch (Exception e) { msg = e.Message.ToString(); }

            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }
        #endregion master systemwide


        #region jsea 
        //task type
        [HttpPost("get_master_tasktype", Name = "get_master_tasktype")]
        //[ValidateAntiForgeryToken]
        public string get_master_tasktype(LoadMasterPageModel param)
        {
            ClassMasterData cls = new ClassMasterData();
            return cls.get_master_tasktype(param);
        }
        [HttpPost("set_master_tasktype", Name = "set_master_tasktype")]
        //[ValidateAntiForgeryToken]
        public string set_master_tasktype(SetDataMasterModel param)
        {
            ClassMasterData cls = new ClassMasterData();
            return cls.set_master_tasktype(param);
        }


        //Tag ID/Equipment
        [HttpPost("get_master_tagid", Name = "get_master_tagid")]
        //[ValidateAntiForgeryToken]
        public string get_master_tagid(LoadMasterPageModel param)
        {
            ClassMasterData cls = new ClassMasterData();
            return cls.get_master_tagid(param);
        }
        [HttpPost("set_master_tagid", Name = "set_master_tagid")]
        //[ValidateAntiForgeryToken]
        public string set_master_tagid(SetDataMasterModel param)
        {
            ClassMasterData cls = new ClassMasterData();
            return cls.set_master_tagid(param);
        }

        //mandatory_note
        [HttpPost("get_master_mandatorynote", Name = "get_master_mandatorynote")]
        //[ValidateAntiForgeryToken]
        public string get_master_mandatorynote(LoadMasterPageModel param)
        {
            ClassMasterData cls = new ClassMasterData();
            return cls.get_master_mandatorynote(param);
        }
        [HttpPost("set_master_mandatorynote", Name = "set_master_mandatorynote")]
        //[ValidateAntiForgeryToken]
        public string set_master_mandatorynote(SetDataMasterModel param)
        {
            ClassMasterData cls = new ClassMasterData();
            return cls.set_master_mandatorynote(param);
        }
        #endregion jsea 

        #region hazop module 
        //Functional Location
        [HttpPost("get_master_functionallocation", Name = "get_master_functionallocation")]
        //[ValidateAntiForgeryToken]
        public string get_master_functionallocation(LoadMasterPageModel param)
        {
            ClassMasterData cls = new ClassMasterData();
            return cls.get_master_functionallocation(param);

        }
        [HttpPost("set_master_functionallocation", Name = "set_master_functionallocation")]
        //[ValidateAntiForgeryToken]
        public string set_master_functionallocation(SetMasterGuideWordsModel param)
        {
            ClassMasterData cls = new ClassMasterData();
            return cls.set_master_functionallocation(param);
        }

        //guidewords
        [HttpPost("get_master_guidewords", Name = "get_master_guidewords")]
        //[ValidateAntiForgeryToken]
        public string get_master_guidewords(LoadMasterPageModel param)
        {
            ClassMasterData cls = new ClassMasterData();
            return cls.get_master_guidewords(param);

        }
        [HttpPost("set_master_guidewords", Name = "set_master_guidewords")]
        //[ValidateAntiForgeryToken]
        public string set_master_guidewords(SetMasterGuideWordsModel param)
        {
            ClassMasterData cls = new ClassMasterData();
            return cls.set_master_guidewords(param);
        }
        #endregion hazop module 

        #region Manage User
        [HttpPost("get_master_contractlist", Name = "get_master_contractlist")]
        //[ValidateAntiForgeryToken]
        public string get_master_contractlist(LoadMasterPageModel param)
        {
            ClassMasterData cls = new ClassMasterData();
            return cls.get_master_contractlist(param);

        }
        [HttpPost("set_master_contractlist", Name = "set_master_contractlist")]
        //[ValidateAntiForgeryToken]
        public string set_master_contractlist(SetMasterGuideWordsModel param)
        {
            ClassMasterData cls = new ClassMasterData();
            return cls.set_master_contractlist(param);
        }

        #endregion Manage User


        #region hra
        //get_master_sub_area_group
        [HttpPost("get_master_sub_area_group", Name = "get_master_sub_area_group")]
        //[ValidateAntiForgeryToken]
        public string get_master_sub_area_group(LoadMasterPageModel param)
        {
            ClassMasterData cls = new ClassMasterData();
            return cls.get_master_sub_area_group(param);
        }

        [HttpPost("set_master_sub_area_group", Name = "set_master_sub_area_group")]
        //[ValidateAntiForgeryToken]
        public string set_master_sub_area_group(SetMasterGuideWordsModel param)
        {
            ClassMasterData cls = new ClassMasterData();
            return cls.set_master_sub_area_group(param);
        }

        [HttpPost("get_master_sub_area_equipmet", Name = "get_master_sub_area_equipmet")]
        //[ValidateAntiForgeryToken]
        public string get_master_sub_area_equipmet(LoadMasterPageModel param)
        {
            ClassMasterData cls = new ClassMasterData();
            return cls.get_master_sub_area_equipmet(param);
        }

        [HttpPost("set_master_sub_area_equipmet", Name = "set_master_sub_area_equipmet")]
        //[ValidateAntiForgeryToken]
        public string set_master_sub_area_equipmet(SetMasterGuideWordsModel param)
        {
            ClassMasterData cls = new ClassMasterData();
            return cls.set_master_sub_area_equipmet(param);
        }

        //get_master_hazard_type
        [HttpPost("get_master_hazard_type", Name = "get_master_hazard_type")]
        //[ValidateAntiForgeryToken]
        public string get_master_hazard_type(LoadMasterPageModel param)
        {
            ClassMasterData cls = new ClassMasterData();
            return cls.get_master_hazard_type(param);
        }
        [HttpPost("set_master_hazard_type", Name = "set_master_hazard_type")]
        //[ValidateAntiForgeryToken]
        public string set_master_hazard_type(SetMasterGuideWordsModel param)
        {
            ClassMasterData cls = new ClassMasterData();
            return cls.set_master_hazard_type(param);
        }

        //get_master_hazard_riskfactors
        [HttpPost("get_master_hazard_riskfactors", Name = "get_master_hazard_riskfactors")]
        //[ValidateAntiForgeryToken]
        public string get_master_hazard_riskfactors(LoadMasterPageModel param)
        {
            ClassMasterData cls = new ClassMasterData();
            return cls.get_master_hazard_riskfactors(param);
        }
        [HttpPost("set_master_hazard_riskfactors", Name = "set_master_hazard_riskfactors")]
        //[ValidateAntiForgeryToken]
        public string set_master_hazard_riskfactors(SetMasterGuideWordsModel param)
        {
            ClassMasterData cls = new ClassMasterData();
            return cls.set_master_hazard_riskfactors(param);

        }


        //get_master_group_list
        [HttpPost("get_master_group_list", Name = "get_master_group_list")]
        //[ValidateAntiForgeryToken]
        public string get_master_group_list(LoadMasterPageBySectionModel param)
        {
            ClassMasterData cls = new ClassMasterData();
            return cls.get_master_group_list(param);
        }
        [HttpPost("set_master_group_list", Name = "set_master_group_list")]
        //[ValidateAntiForgeryToken]
        public string set_master_group_list(SetMasterGuideWordsModel param)
        {
            ClassMasterData cls = new ClassMasterData();
            return cls.set_master_group_list(param); 
        }

        //get_master_worker_group
        [HttpPost("get_master_worker_group", Name = "get_master_worker_group")]
        //[ValidateAntiForgeryToken]
        public string get_master_worker_group(LoadMasterPageBySectionModel param)
        {
            ClassMasterData cls = new ClassMasterData();
            return cls.get_master_worker_group(param);
        }
        [HttpPost("set_master_worker_group", Name = "set_master_worker_group")]
        //[ValidateAntiForgeryToken]
        public string set_master_worker_group(SetMasterGuideWordsModel param)
        {
            ClassMasterData cls = new ClassMasterData();
            return cls.set_master_worker_group(param); 
        }

        #endregion hra


    }
}
