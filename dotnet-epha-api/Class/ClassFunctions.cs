
using System.Data;
using System.Data.SqlClient;
using System;
using System.IO;


namespace dotnet6_epha_api.Class
{
    public class FileHandler
    {
        // ฟังก์ชันสำหรับตรวจสอบและ sanitize ค่า module_name
        private static string SanitizeModuleName(string moduleName)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                moduleName = moduleName.Replace(c, '_');
            }
            return moduleName;
        }

        // ฟังก์ชันสำหรับสร้างเส้นทางไฟล์ที่ปลอดภัย
        private static string CreateSafePath(string basePath, string moduleName)
        {
            string sanitizedModuleName = SanitizeModuleName(moduleName);
            string combinedPath = Path.Combine(basePath, sanitizedModuleName);
            return Path.GetFullPath(combinedPath);
        }

        public void HandleFilesAttachedFileTemp(string module_name, ref string downloadPath, ref string folderTemplate)
        {
            // ระบุเส้นทางฐานสำหรับการดาวน์โหลดและโฟลเดอร์เทมเพลต
            string baseDownloadPath = @"/AttachedFileTemp/";
            string baseFolderTemplatePath = @"/wwwroot/AttachedFileTemp/";

            // สร้างเส้นทางที่ปลอดภัย
            string _DownloadPath = CreateSafePath(baseDownloadPath, module_name);
            string _FolderTemplate = CreateSafePath(baseFolderTemplatePath, module_name);

            // ตรวจสอบว่าเส้นทางที่สร้างไม่ออกนอกขอบเขตที่กำหนด
            if (!IsPathInsideBasePath(_DownloadPath, baseDownloadPath) || !IsPathInsideBasePath(_FolderTemplate, baseFolderTemplatePath))
            {
                throw new UnauthorizedAccessException("Attempted to access a path outside the allowed directory.");
            }

            // กำหนดค่าให้กับตัวแปรที่เป็น ref
            downloadPath = _DownloadPath;
            folderTemplate = _FolderTemplate;
        }
         
        // ฟังก์ชันสำหรับตรวจสอบว่าเส้นทางที่สร้างไม่ออกนอกขอบเขตที่กำหนด
        private static bool IsPathInsideBasePath(string path, string basePath)
        {
            var fullPath = Path.GetFullPath(path);
            var fullBasePath = Path.GetFullPath(basePath);
            return fullPath.StartsWith(fullBasePath, StringComparison.OrdinalIgnoreCase);
        }
    }

    public class ClassFunctions
    { 
        public static void refAttachedFilePath(string module, ref string downloadPath, ref string folder, ref string path, ref string folderTemplate)
        { 
            downloadPath = $"/AttachedFileTemp/{module}/";
            folder = $"/wwwroot/AttachedFileTemp/{module}/";
            path = Path.Combine(Directory.GetCurrentDirectory(), folder.Replace("~", ""));
            folderTemplate = Path.Combine(Directory.GetCurrentDirectory(), "/wwwroot/AttachedFileTemp/".Replace("~", ""));
        }

     
        #region Utility Functions

        public string ChkSqlNum(object str, string nType)
        {
            if (str == null || Convert.IsDBNull(str) || (str?.ToString() ?? "").ToUpper() == "NULL")
                return "NULL";

            try
            {
                return nType switch
                {
                    "N" => Convert.ToInt64(str).ToString(),
                    "D" => Convert.ToDouble(str).ToString(),
                    _ => "NULL"
                };
            }
            catch
            {
                return "NULL";
            }
        }

        public string ChkSqlNum(object str, string nType, int iLength)
        {
            if (str == null || Convert.IsDBNull(str) || (str?.ToString() ?? "").ToUpper() == "NULL")
                return "NULL";

            try
            {
                double num = Convert.ToDouble(str);
                return nType switch
                {
                    "N" => Convert.ToInt64(num).ToString(),
                    "D" => num.ToString($"F{iLength}"),
                    _ => "NULL"
                };
            }
            catch
            {
                return "NULL";
            }
        }

        public string ChkSqlStr(object str, int length)
        {
            if (str == null || Convert.IsDBNull(str) || string.IsNullOrWhiteSpace(str.ToString()) || (str?.ToString() ?? "").ToLower() == "null")
                return "null";

            string str1 = (str?.ToString() ?? "").Replace("'", "''");

            return $"'{(str1.Length > length ? str1.Substring(0, length) : str1)}'";
        }

        public string ChkSqlDateYYYYMMDD(object sDate)
        {
            if (sDate == null || Convert.IsDBNull(sDate) || string.IsNullOrWhiteSpace(sDate.ToString()))
                return "NULL";

            try
            {
                string[] dateParts = sDate.ToString().Split('-');
                if (dateParts.Length == 3)
                {
                    sDate = $"{dateParts[0]}{dateParts[1].PadLeft(2, '0')}{dateParts[2].PadLeft(2, '0')}";
                }

                DateTime tsDate = DateTime.ParseExact(sDate.ToString(), "yyyyMMdd", null);

                if (tsDate.Year > 2500)
                {
                    tsDate = tsDate.AddYears(-543);
                }
                if (tsDate.Year < 2000)
                {
                    tsDate = tsDate.AddYears(543);
                }

                return $"CONVERT(date, '{tsDate:yyyyMMdd}')";
            }
            catch
            {
                return "NULL";
            }
        }

        #endregion Utility Functions
    }
}

