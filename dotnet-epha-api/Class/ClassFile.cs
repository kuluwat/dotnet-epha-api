

namespace dotnet_epha_api.Class
{
    using System.Data;
    public class ClassFile
    {

        public static DataTable refMsg(string status, string remark, string? seq_new = "")
        {
            DataTable dtMsg = new DataTable();
            dtMsg.Columns.Add("status");
            dtMsg.Columns.Add("remark");
            dtMsg.Columns.Add("seq_new");
            dtMsg.TableName = "msg";
            dtMsg.AcceptChanges();

            dtMsg.Rows.Add(dtMsg.NewRow());
            dtMsg.Rows[0]["status"] = status;
            dtMsg.Rows[0]["remark"] = remark;
            dtMsg.Rows[0]["seq_new"] = seq_new;
            return dtMsg;
        }
        public static DataTable refMsgSave(string status, string remark, string? seq_new = "", string? pha_seq = "", string? pha_no = "", string? pha_status = "")
        {
            DataTable dtMsg = new DataTable();
            dtMsg.Columns.Add("status");
            dtMsg.Columns.Add("remark");
            dtMsg.Columns.Add("seq_new");
            dtMsg.Columns.Add("pha_seq");
            dtMsg.Columns.Add("pha_no");
            dtMsg.Columns.Add("pha_status");
            dtMsg.TableName = "msg";
            dtMsg.AcceptChanges();

            dtMsg.Rows.Add(dtMsg.NewRow());
            dtMsg.Rows[0]["status"] = status;
            dtMsg.Rows[0]["remark"] = remark;
            dtMsg.Rows[0]["seq_new"] = seq_new;
            dtMsg.Rows[0]["pha_seq"] = pha_seq;
            dtMsg.Rows[0]["pha_no"] = pha_no;
            dtMsg.Rows[0]["pha_status"] = pha_status;
            return dtMsg;
        }
        public static DataTable DatatableMsg()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("status");
            return dt;
        }
        public static DataTable DatatableFile()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("ATTACHED_FILE_NAME");
            dt.Columns.Add("ATTACHED_FILE_PATH");
            dt.Columns.Add("ATTACHED_FILE_OF");
            dt.Columns.Add("STATUS");
            dt.Columns.Add("IMPORT_DATA_MSG");
            dt.TableName = "msg";
            return dt;
        }
        public static void AddRowToDataTable(ref DataTable dtdef, string fileName, string filePath, string msgError)
        {
            // ตรวจสอบว่า DataTable ไม่เป็น null
            if (dtdef == null) throw new ArgumentNullException(nameof(dtdef));

            if (dtdef != null &&
                dtdef.Columns.Contains("ATTACHED_FILE_NAME") &&
                dtdef.Columns.Contains("ATTACHED_FILE_PATH") &&
                dtdef.Columns.Contains("IMPORT_DATA_MSG") &&
                dtdef.Columns.Contains("STATUS"))
            {
                // สร้างแถวใหม่
                DataRow newRow = dtdef.NewRow();
                newRow["ATTACHED_FILE_NAME"] = fileName ?? "";
                newRow["ATTACHED_FILE_PATH"] = filePath ?? "";
                newRow["IMPORT_DATA_MSG"] = msgError ?? "";
                newRow["STATUS"] = string.IsNullOrEmpty(msgError) ? "true" : "error";

                // เพิ่มแถวใหม่ลงใน DataTable
                dtdef.Rows.Add(newRow);
            }

        }
        public static string copy_file_data_to_server(
        ref string file_name, ref string file_download_name, ref string _file_fullpath_name,
        IFormFileCollection? files,
        string? folder = "_temp",
        string? file_part = "import",
        string? file_doc = "docno",
        bool tempFile = false,
        bool folderCopyFile = false)
        {
            //"C:\xx\dotnet-epha-api\wwwroot\AttachedFileTemp\_temp\JSEA AttendeeSheet - xxxxx.xlsx"

            // ตรวจสอบค่า files
            if (files == null || files.Count == 0)
            {
                return "Invalid files.";
            }

            // ตรวจสอบค่า folder
            if (string.IsNullOrWhiteSpace(folder) || folder.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            {
                return "Invalid folder.";
            }
            // ตรวจสอบค่า part
            if (string.IsNullOrWhiteSpace(file_part) || file_part.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            {
                return "Invalid part.";
            }

            try
            {
                IFormFile file = files[0];
                if (file != null)
                {
                    string fileTemp = file.FileName ?? "";
                    if (string.IsNullOrEmpty(fileTemp))
                    {
                        return ("Invalid file name.");
                    }

                    // ตรวจสอบและดึงเฉพาะชื่อไฟล์จาก fileTemp
                    string safeFileTemp = Path.GetFileName(fileTemp);
                    if (string.IsNullOrEmpty(safeFileTemp))
                    {
                        return ("Invalid file safe file temp.");
                    }
                    // ตรวจสอบค่า safeFileTemp
                    if (string.IsNullOrWhiteSpace(safeFileTemp) || safeFileTemp.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
                    {
                        return "Invalid file safe file temp(2).";
                    }

                    // ตรวจสอบ Format ชื่อไฟล์
                    // เพิ่มช่วงของตัวอักษรภาษาไทยอยู่ในช่วง 0x0E00 ถึง 0x0E7F (หรือ เ ถึง ๏)
                    char[] AllowedCharacters = "()abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-. ".ToCharArray()
                    .Concat(Enumerable.Range(0x0E00, 0x0E7F - 0x0E00 + 1).Select(c => (char)c)).ToArray();

                    if (safeFileTemp.Any(c => !AllowedCharacters.Contains(c)))
                    {
                        return ("Input file name contains invalid characters.");
                    }

                    // ตรวจสอบนามสกุลไฟล์ว่าถูกต้องหรือไม่
                    string extension = Path.GetExtension(safeFileTemp).ToLowerInvariant();
                    if (string.IsNullOrEmpty(extension))
                    {
                        return ("File does not have a valid extension.");
                    }

                    // อนุญาตเฉพาะไฟล์ที่กำหนด
                    string[] allowedExtensionsExcel = { ".xlsx", ".xls" };
                    string[] allowedExtensions = { ".xlsx", ".xls", ".pdf", ".doc", ".docx", ".png", ".jpg", ".gif", ".eml", ".msg" };

                    if (tempFile && !allowedExtensionsExcel.Contains(extension))
                    {
                        return ("Invalid file type. Only Excel files are allowed.");
                    }
                    else if (!tempFile && !allowedExtensions.Contains(extension))
                    {
                        return ("Invalid file type.");
                    }

                    // สร้างเส้นทางตามไฟล์ต้นทาง
                    string finalRootDir = "";
                    string templatewwwwRootDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot");
                    if (!Directory.Exists(templatewwwwRootDir))
                    {
                        throw new DirectoryNotFoundException("Folder directory not found.");
                    }

                    string templateRootDir = Path.Combine(templatewwwwRootDir, "AttachedFileTemp");
                    if (!Directory.Exists(templateRootDir))
                    {
                        throw new DirectoryNotFoundException("Folder directory not found.");
                    }

                    // ตรวจสอบไดเร็กทอรี file ภายใน module
                    string templateDir = Path.Combine(templateRootDir, folder);
                    templateDir = Path.GetFullPath(templateDir);
                    if (!templateDir.StartsWith(templateRootDir, StringComparison.OrdinalIgnoreCase))
                    {
                        return ("Temp directory is outside of the Template directory.");
                    }

                    if (folderCopyFile)
                    {
                        // กรณีที่เป็นการ copy file จะเป็นอีกชั้น
                        string templateCopyDir = Path.Combine(templateDir, "copy");
                        templateCopyDir = Path.GetFullPath(templateCopyDir);
                        if (!templateCopyDir.StartsWith(templateDir, StringComparison.OrdinalIgnoreCase))
                        {
                            return ("TempCopy directory is outside of the Template directory.");
                        }
                        finalRootDir = templateCopyDir;
                    }
                    else
                    {
                        finalRootDir = templateDir;
                    }

                    if (string.IsNullOrEmpty(finalRootDir))
                    {
                        return ("Invalid finalRootDir.");
                    }

                    if (!Directory.Exists(finalRootDir))
                    {
                        Directory.CreateDirectory(finalRootDir);
                    }

                    // สร้างชื่อไฟล์ใหม่
                    var datetime_run = DateTime.Now.ToString("yyyyMMddHHmm");
                    string retFileName = "";
                    if (!string.IsNullOrEmpty(file_doc) && !string.IsNullOrEmpty(file_part))
                    {
                        retFileName = $"{file_doc}-{file_part}-{datetime_run}";
                    }
                    else if (!string.IsNullOrEmpty(file_doc) && string.IsNullOrEmpty(file_part))
                    {
                        retFileName = $"{file_doc}-{datetime_run}";
                    }
                    else if (string.IsNullOrEmpty(file_doc) && !string.IsNullOrEmpty(file_part))
                    {
                        retFileName = $"{file_part}-{datetime_run}";
                    }
                    else { retFileName = $"{datetime_run}"; }

                    if (string.IsNullOrEmpty(retFileName))
                    {
                        return ("File does not have a valid new file.");
                    }
                    string sourceFile = $"{retFileName}{extension}";
                    if (string.IsNullOrEmpty(sourceFile))
                    {
                        return ("File does not have a valid path source file.");
                    }
                    if (sourceFile.Any(c => !AllowedCharacters.Contains(c)) || string.IsNullOrWhiteSpace(sourceFile) || sourceFile.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || sourceFile.Contains("..") || sourceFile.Contains("\\"))
                    {
                        return ("Invalid file name.");
                    }

                    // สร้างเส้นทางไฟล์ปลายทางแบบสัมบูรณ์และตรวจสอบทีละชั้น
                    string newFileNameFullPath = Path.Combine(finalRootDir, sourceFile);
                    newFileNameFullPath = Path.GetFullPath(newFileNameFullPath);
                    if (string.IsNullOrEmpty(newFileNameFullPath))
                    {
                        return ("File does not have a valid path new file.");
                    }
                    if (!newFileNameFullPath.StartsWith(finalRootDir, StringComparison.OrdinalIgnoreCase))
                    {
                        return ("Attempt to access unauthorized path.");
                    }

                    //สำหรับใช้ใน excel data
                    //"C:\xx\dotnet-epha-api\wwwroot\AttachedFileTemp\_temp\JSEA AttendeeSheet - xxxxx.xlsx"
                    _file_fullpath_name = newFileNameFullPath;
                    if (string.IsNullOrEmpty(_file_fullpath_name))
                    {
                        return ("File does not have a valid full path new file.");
                    }
                    // คัดลอกไฟล์จากต้นทางไปปลายทาง
                    using (var fileStream = new FileStream(newFileNameFullPath, FileMode.Create))
                    {
                        file.CopyTo(fileStream);
                    }

                    //คืนชื่อเดิมกลับไปให้
                    //"JSEA AttendeeSheet - xxxxx.xlsx"
                    file_name = safeFileTemp;
                    if (string.IsNullOrEmpty(file_name))
                    {
                        return ("File does not have a valid file name.");
                    }

                    //"AttachedFileTemp\_temp\JSEA AttendeeSheet - xxxxx.xlsx"
                    file_download_name = Path.GetRelativePath(templatewwwwRootDir, newFileNameFullPath);
                    if (string.IsNullOrEmpty(file_download_name))
                    {
                        return ("File does not have a valid file download name.");
                    }
                    // ตรวจสอบความปลอดภัย
                    if (file_download_name.Contains(".."))
                    {
                        return ("The resulting relative path is attempting to access outside the intended directory.");
                    }


                }
            }
            catch (Exception ex)
            {
                return $"An error occurred while processing your request.{ex.Message.ToString()}";
            }

            return "";
        }

        public static string copy_file_excel_template(
         ref string file_name, ref string file_download_name, ref string _file_fullpath_name,
         string? folder = "other", string? file_part = "Report", string? file_doc = "template")
        {
            //"C:\xx\dotnet-epha-api\wwwroot\AttachedFileTemp\JSEA AttendeeSheet Template.xlsx"
            //"C:\xx\dotnet-epha-api\wwwroot\AttachedFileTemp\jsea\JSEA AttendeeSheet - xxxxx.xlsx"

            // ตรวจสอบค่า folder -> hazop, jsea, whatif, hra
            if (string.IsNullOrWhiteSpace(folder) || folder.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            {
                return "Invalid folder.";
            }
            // ตรวจสอบค่า part
            if (string.IsNullOrWhiteSpace(file_part) || file_part.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            {
                return "Invalid part.";
            }

            try
            {
                #region part check file name
                // ตรวจสอบชื่อไฟล์ -> HAZOP Report Template.xlsx, JSEA Report Template.xlsx 
                string defFileTemp = $"{folder?.ToUpper()} {file_part}";
                if (string.IsNullOrEmpty(defFileTemp))
                {
                    return ("Invalid file file temp.");
                }
                string safeFileTemp = $"{defFileTemp} Template.xlsx";
                if (string.IsNullOrEmpty(safeFileTemp))
                {
                    return ("Invalid file safe file temp.");
                }

                // ตรวจสอบ Format ชื่อไฟล์
                // เพิ่มช่วงของตัวอักษรภาษาไทยอยู่ในช่วง 0x0E00 ถึง 0x0E7F (หรือ เ ถึง ๏)
                char[] AllowedCharacters = "()abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-. ".ToCharArray()
                .Concat(Enumerable.Range(0x0E00, 0x0E7F - 0x0E00 + 1).Select(c => (char)c)).ToArray();

                if (safeFileTemp.Any(c => !AllowedCharacters.Contains(c)))
                {
                    return ("Input file name contains invalid characters.");
                }

                // ตรวจสอบนามสกุลไฟล์ว่าถูกต้องหรือไม่
                string extension = Path.GetExtension(safeFileTemp).ToLowerInvariant();
                if (string.IsNullOrEmpty(extension))
                {
                    return ("File does not have a valid extension.");
                }

                // อนุญาตเฉพาะไฟล์ที่กำหนด
                string[] allowedExtensionsExcel = { ".xlsx", ".xls" };
                if (!allowedExtensionsExcel.Contains(extension))
                {
                    return ("Invalid file type. Only Excel files are allowed.");
                }

                #endregion part check file name

                // สร้างเส้นทางตามไฟล์ต้นทาง
                string finalRootDir = "";
                string templatewwwwRootDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot");
                if (!Directory.Exists(templatewwwwRootDir))
                {
                    throw new DirectoryNotFoundException("Folder directory not found.");
                }
                string templateRootDir = Path.Combine(templatewwwwRootDir, "AttachedFileTemp");
                if (!Directory.Exists(templateRootDir))
                {
                    throw new DirectoryNotFoundException("Folder directory not found.");
                }

                // ตรวจสอบไดเร็กทอรี file template ภายใน folder template
                string fileNameFullPath = Path.Combine(templateRootDir, safeFileTemp);
                fileNameFullPath = Path.GetFullPath(fileNameFullPath);
                if (!fileNameFullPath.StartsWith(templateRootDir, StringComparison.OrdinalIgnoreCase))
                {
                    return ("Temp directory is outside of the Template directory.");
                }

                #region create new file
                // สร้างชื่อไฟล์ใหม่ format ไม่เหมือน function อื่นนะ
                var datetime_run = DateTime.Now.ToString("yyyyMMddHHmm");
                string retFileName = "";
                if (!string.IsNullOrEmpty(file_name))
                {
                    retFileName = $"{file_name}-{file_doc} {datetime_run}";
                }
                else
                {
                    retFileName = $"{defFileTemp}-{file_doc} {datetime_run}";
                }

                if (string.IsNullOrEmpty(retFileName))
                {
                    return ("File does not have a valid new file.");
                }
                string sourceFile = $"{retFileName}{extension}";
                if (string.IsNullOrEmpty(sourceFile))
                {
                    return ("File does not have a valid path source file.");
                }
                if (sourceFile.Any(c => !AllowedCharacters.Contains(c)) || string.IsNullOrWhiteSpace(sourceFile) || sourceFile.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || sourceFile.Contains("..") || sourceFile.Contains("\\"))
                {
                    return ("Invalid file name.");
                }
                #endregion create new file

                // สร้างเส้นทางไฟล์ปลายทางแบบสัมบูรณ์และตรวจสอบทีละชั้น 
                string templateModuleDir = Path.Combine(templateRootDir, $"{folder}");
                if (!Directory.Exists(templateModuleDir))
                {
                    throw new DirectoryNotFoundException("Folder Module directory not found.");
                }

                finalRootDir = templateModuleDir;
                if (string.IsNullOrEmpty(finalRootDir))
                {
                    return ("Invalid finalRootDir.");
                }

                // full path folder ของปลายทาง ว่าถูกต้องหรือป่าว
                if (!Directory.Exists(finalRootDir))
                {
                    Directory.CreateDirectory(finalRootDir);
                }

                // ตรวจสอบไดเร็กทอรี file ใหม่ที่ copy template มาตั้งต้น ภายใน folder module  
                string newFileNameFullPath = Path.Combine(finalRootDir, sourceFile);
                newFileNameFullPath = Path.GetFullPath(newFileNameFullPath);
                if (string.IsNullOrEmpty(newFileNameFullPath))
                {
                    return ("File does not have a valid path new file.");
                }
                if (!newFileNameFullPath.StartsWith(finalRootDir, StringComparison.OrdinalIgnoreCase))
                {
                    return ("Attempt to access unauthorized path.");
                }

                //สำหรับใช้ใน excel data
                //"C:\xx\dotnet-epha-api\wwwroot\AttachedFileTemp\jsea\JSEA AttendeeSheet - xxxxx.xlsx"
                _file_fullpath_name = newFileNameFullPath;
                if (string.IsNullOrEmpty(_file_fullpath_name))
                {
                    return ("File does not have a valid full path new file.");
                }

                // คัดลอกไฟล์จากต้นทางไปปลายทาง
                try
                {
                    File.Copy(fileNameFullPath, newFileNameFullPath, overwrite: true);

                    var checkFileInfo = new FileInfo(newFileNameFullPath);
                    if (!checkFileInfo.Exists || checkFileInfo.IsReadOnly)
                    {
                        return ("File permissions are not correctly set.");
                    }
                }
                catch (IOException ex)
                {
                    // Handle the exception, log it, or return an error message
                    throw new InvalidOperationException("Failed to copy the file.", ex);
                }

                #region retrun ค่ากลับ 

                //คืนชื่อเดิมกลับไปให้
                //JSEA AttendeeSheet - xxxxx.xlsx"
                file_name = sourceFile;
                if (string.IsNullOrEmpty(file_name))
                {
                    return ("File does not have a valid file name.");
                }
                //"AttachedFileTemp\jsea\JSEA AttendeeSheet - xxxxx.xlsx"
                file_download_name = Path.GetRelativePath(templatewwwwRootDir, newFileNameFullPath);
                if (string.IsNullOrEmpty(file_download_name))
                {
                    return ("File does not have a valid file download name.");
                }
                // ตรวจสอบความปลอดภัย
                if (file_download_name.Contains(".."))
                {
                    return ("The resulting relative path is attempting to access outside the intended directory.");
                }
                #endregion retrun ค่ากลับ 

            }
            catch (Exception ex)
            {
                return $"An error occurred while processing your request.{ex.Message.ToString()}";
            }

            return "";
        }


        public static string copy_file_duplicate(string file_name, ref string file_download_name, ref string _file_fullpath_name, string? folder = "other")
        {
            //"C:\xx\dotnet-epha-api\wwwroot\AttachedFileTemp\JSEA AttendeeSheet Template.pdf"
            //"C:\xx\dotnet-epha-api\wwwroot\AttachedFileTemp\jsea\JSEA AttendeeSheet.pdf"
            //to
            //"C:\xx\dotnet-epha-api\wwwroot\AttachedFileTemp\JSEA AttendeeSheet xxxx.pdf"
            //"C:\xx\dotnet-epha-api\wwwroot\AttachedFileTemp\jsea\JSEA AttendeeSheet xxxx.pdf"

            // ตรวจสอบค่า folder -> hazop, jsea, whatif, hra 
            if (string.IsNullOrWhiteSpace(folder) || folder.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            {
                return "Invalid folder.";
            }
            if (string.IsNullOrWhiteSpace(file_name) || file_name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            {
                return "Invalid file name.";
            }
            try
            {
                #region part check file name
                string fileTemp = $"{file_name}";
                if (string.IsNullOrEmpty(fileTemp))
                {
                    return ("Invalid file name.");
                }

                // ตรวจสอบและดึงเฉพาะชื่อไฟล์จาก fileTemp
                string safeFileTemp = Path.GetFileName(fileTemp);
                if (string.IsNullOrEmpty(safeFileTemp))
                {
                    return ("Invalid file safe file temp.");
                }
                // ตรวจสอบค่า safeFileTemp
                if (string.IsNullOrWhiteSpace(safeFileTemp) || safeFileTemp.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
                {
                    return "Invalid file safe file temp(2).";
                }

                // ตรวจสอบ Format ชื่อไฟล์
                // เพิ่มช่วงของตัวอักษรภาษาไทยอยู่ในช่วง 0x0E00 ถึง 0x0E7F (หรือ เ ถึง ๏)
                char[] AllowedCharacters = "()abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-. ".ToCharArray()
                .Concat(Enumerable.Range(0x0E00, 0x0E7F - 0x0E00 + 1).Select(c => (char)c)).ToArray();

                if (safeFileTemp.Any(c => !AllowedCharacters.Contains(c)))
                {
                    return ("Input file name contains invalid characters.");
                }

                // ตรวจสอบนามสกุลไฟล์ว่าถูกต้องหรือไม่
                string extension = Path.GetExtension(safeFileTemp).ToLowerInvariant();
                if (string.IsNullOrEmpty(extension))
                {
                    return ("File does not have a valid extension.");
                }

                // อนุญาตเฉพาะไฟล์ที่กำหนด
                string[] allowedExtensions = { ".xlsx", ".xls", ".pdf", ".doc", ".docx", ".png", ".jpg", ".gif", ".eml", ".msg" };
                if (!allowedExtensions.Contains(extension))
                {
                    return ("Invalid file type. Only Excel files are allowed.");
                }

                #endregion part check file name

                // สร้างเส้นทางตามไฟล์ต้นทาง
                string finalRootDir = "";
                string templatewwwwRootDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot");
                if (!Directory.Exists(templatewwwwRootDir))
                {
                    throw new DirectoryNotFoundException("Folder directory not found.");
                }
                string templateRootDir = Path.Combine(templatewwwwRootDir, "AttachedFileTemp");
                if (!Directory.Exists(templateRootDir))
                {
                    throw new DirectoryNotFoundException("Folder directory not found.");
                }

                // ตรวจสอบไดเร็กทอรี file template ภายใน folder template
                string fileNameFullPath = Path.Combine(templateRootDir, safeFileTemp);
                fileNameFullPath = Path.GetFullPath(fileNameFullPath);
                if (!fileNameFullPath.StartsWith(templateRootDir, StringComparison.OrdinalIgnoreCase))
                {
                    return ("Temp directory is outside of the Template directory.");
                }

                #region create new file
                // สร้างชื่อไฟล์ใหม่ format ไม่เหมือน function อื่นนะ
                var datetime_run = DateTime.Now.ToString("yyyyMMddHHmm");
                string defFileTemp = Path.GetFileNameWithoutExtension(fileTemp);
                if (string.IsNullOrEmpty(defFileTemp))
                {
                    return ("File does not have a valid new file temp.");
                }

                string retFileName = $"{defFileTemp} {datetime_run}";
                if (string.IsNullOrEmpty(retFileName))
                {
                    return ("File does not have a valid new file.");
                }
                string sourceFile = $"{retFileName}{extension}";
                if (string.IsNullOrEmpty(sourceFile))
                {
                    return ("File does not have a valid path source file.");
                }
                if (sourceFile.Any(c => !AllowedCharacters.Contains(c)) || string.IsNullOrWhiteSpace(sourceFile) || sourceFile.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || sourceFile.Contains("..") || sourceFile.Contains("\\"))
                {
                    return ("Invalid file name.");
                }
                #endregion create new file

                // สร้างเส้นทางไฟล์ปลายทางแบบสัมบูรณ์และตรวจสอบทีละชั้น 
                string templateModuleDir = Path.Combine(templateRootDir, $"{folder}");
                if (!Directory.Exists(templateModuleDir))
                {
                    throw new DirectoryNotFoundException("Folder Module directory not found.");
                }

                finalRootDir = templateModuleDir;
                if (string.IsNullOrEmpty(finalRootDir))
                {
                    return ("Invalid finalRootDir.");
                }

                // full path folder ของปลายทาง ว่าถูกต้องหรือป่าว
                if (!Directory.Exists(finalRootDir))
                {
                    Directory.CreateDirectory(finalRootDir);
                }

                // ตรวจสอบไดเร็กทอรี file ใหม่ที่ copy template มาตั้งต้น ภายใน folder module  
                string newFileNameFullPath = Path.Combine(finalRootDir, sourceFile);
                newFileNameFullPath = Path.GetFullPath(newFileNameFullPath);
                if (string.IsNullOrEmpty(newFileNameFullPath))
                {
                    return ("File does not have a valid path new file.");
                }
                if (!newFileNameFullPath.StartsWith(finalRootDir, StringComparison.OrdinalIgnoreCase))
                {
                    return ("Attempt to access unauthorized path.");
                }

                //สำหรับใช้ใน excel data
                //"C:\xx\dotnet-epha-api\wwwroot\AttachedFileTemp\jsea\JSEA AttendeeSheet - xxxxx.pdf"
                _file_fullpath_name = newFileNameFullPath;
                if (string.IsNullOrEmpty(_file_fullpath_name))
                {
                    return ("File does not have a valid full path new file.");
                }

                // คัดลอกไฟล์จากต้นทางไปปลายทาง
                try
                {
                    File.Copy(fileNameFullPath, newFileNameFullPath, overwrite: true);

                    var checkFileInfo = new FileInfo(newFileNameFullPath);
                    if (!checkFileInfo.Exists || checkFileInfo.IsReadOnly)
                    {
                        return ("File permissions are not correctly set.");
                    }
                }
                catch (IOException ex)
                {
                    // Handle the exception, log it, or return an error message
                    throw new InvalidOperationException("Failed to copy the file.", ex);
                }

                #region retrun ค่ากลับ 

                //คืนชื่อเดิมกลับไปให้
                //JSEA AttendeeSheet - xxxxx.xlsx"
                file_name = sourceFile;
                if (string.IsNullOrEmpty(file_name))
                {
                    return ("File does not have a valid file name.");
                }
                //"AttachedFileTemp\jsea\JSEA AttendeeSheet - xxxxx.xlsx"
                file_download_name = Path.GetRelativePath(templatewwwwRootDir, newFileNameFullPath);
                if (string.IsNullOrEmpty(file_download_name))
                {
                    return ("File does not have a valid file download name.");
                }
                // ตรวจสอบความปลอดภัย
                if (file_download_name.Contains(".."))
                {
                    return ("The resulting relative path is attempting to access outside the intended directory.");
                }
                #endregion retrun ค่ากลับ 

            }
            catch (Exception ex)
            {
                return $"An error occurred while processing your request.{ex.Message.ToString()}";
            }

            return "";
        }
        public static string check_file_other(string file_name, ref string _file_fullpath_name, string? folder = "")
        {
            //"RAM Template.png"
            //"C:\xx\dotnet-epha-api\wwwroot\AttachedFileTemp\RAM - xxxxx.xlsx" 
            //if (string.IsNullOrWhiteSpace(file_name) || file_name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            if (string.IsNullOrEmpty(file_name))
            {
                return "Invalid file name.";
            }
            try
            {
                #region part check file name 

                string safeFileTemp = $"{file_name}";
                if (string.IsNullOrEmpty(safeFileTemp))
                {
                    return ("Invalid file safe file temp.");
                }

                // ตรวจสอบ Format ชื่อไฟล์
                // เพิ่มช่วงของตัวอักษรภาษาไทยอยู่ในช่วง 0x0E00 ถึง 0x0E7F (หรือ เ ถึง ๏)
                char[] AllowedCharacters = "()abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-. ".ToCharArray()
                .Concat(Enumerable.Range(0x0E00, 0x0E7F - 0x0E00 + 1).Select(c => (char)c)).ToArray();

                if (safeFileTemp.Any(c => !AllowedCharacters.Contains(c)))
                {
                    return ("Input file name contains invalid characters.");
                }

                // ตรวจสอบนามสกุลไฟล์ว่าถูกต้องหรือไม่
                string extension = Path.GetExtension(safeFileTemp).ToLowerInvariant();
                if (string.IsNullOrEmpty(extension))
                {
                    return ("File does not have a valid extension.");
                }

                // อนุญาตเฉพาะไฟล์ที่กำหนด
                string[] allowedExtensions = { ".xlsx", ".xls", ".pdf", ".doc", ".docx", ".png", ".jpg", ".gif", ".eml", ".msg" };
                if (!allowedExtensions.Contains(extension))
                {
                    return ("Invalid file type. Only Excel files are allowed.");
                }

                #endregion part check file name

                // สร้างเส้นทางตามไฟล์ต้นทาง
                string finalRootDir = "";
                string templatewwwwRootDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot");
                if (!Directory.Exists(templatewwwwRootDir))
                {
                    throw new DirectoryNotFoundException("Folder directory not found.");
                }
                string templateRootDir = Path.Combine(templatewwwwRootDir, "AttachedFileTemp");
                if (!Directory.Exists(templateRootDir))
                {
                    throw new DirectoryNotFoundException("Folder directory not found.");
                }

                if (string.IsNullOrWhiteSpace(folder))
                {
                    // สร้างเส้นทางไฟล์ปลายทางแบบสัมบูรณ์และตรวจสอบทีละชั้น 
                    string templateModuleDir = Path.Combine(templateRootDir, $"{folder}");
                    if (!Directory.Exists(templateModuleDir))
                    {
                        throw new DirectoryNotFoundException("Folder Module directory not found.");
                    }
                    finalRootDir = templateModuleDir;
                }
                else
                {
                    finalRootDir = templateRootDir;
                }
                if (string.IsNullOrEmpty(finalRootDir))
                {
                    return ("Invalid finalRootDir.");
                }
                // full path folder ของปลายทาง ว่าถูกต้องหรือป่าว
                if (!Directory.Exists(finalRootDir))
                {
                    Directory.CreateDirectory(finalRootDir);
                }

                // ตรวจสอบไดเร็กทอรี file template ภายใน folder template
                string fileNameFullPath = Path.Combine(finalRootDir, safeFileTemp);
                fileNameFullPath = Path.GetFullPath(fileNameFullPath);
                if (!fileNameFullPath.StartsWith(finalRootDir, StringComparison.OrdinalIgnoreCase))
                {
                    return ("Temp directory is outside of the Template directory.");
                }

                #region retrun ค่ากลับ 
                //สำหรับใช้ 
                _file_fullpath_name = fileNameFullPath;
                if (string.IsNullOrEmpty(_file_fullpath_name))
                {
                    return ("File does not have a valid full path new file.");
                }
                #endregion retrun ค่ากลับ 

            }
            catch (Exception ex)
            {
                return $"An error occurred while processing your request.{ex.Message.ToString()}";
            }

            return "";
        }
        public static string check_file_on_server(string? folder = "other", string? _file_fullpath_name = "")
        {
            if (string.IsNullOrWhiteSpace(folder) || folder.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            {
                return "Invalid folder.";
            }
            if (string.IsNullOrEmpty(_file_fullpath_name))
            {
                return "Invalid file fullpath.";
            }

            try
            {
                string fileTemp = _file_fullpath_name ?? "";
                if (string.IsNullOrEmpty(fileTemp))
                {
                    return ("Invalid file name.");
                }

                // ตรวจสอบและดึงเฉพาะชื่อไฟล์จาก fileTemp
                string safeFileTemp = Path.GetFileName(fileTemp);
                if (string.IsNullOrEmpty(safeFileTemp))
                {
                    return ("Invalid file name.");
                }

                // ตรวจสอบค่า safeFileTemp
                if (string.IsNullOrWhiteSpace(safeFileTemp) || safeFileTemp.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
                {
                    return "Invalid file name.";
                }

                // ตรวจสอบ Format ชื่อไฟล์
                // เพิ่มช่วงของตัวอักษรภาษาไทยอยู่ในช่วง 0x0E00 ถึง 0x0E7F (หรือ เ ถึง ๏)
                char[] AllowedCharacters = "()abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-. ".ToCharArray()
                .Concat(Enumerable.Range(0x0E00, 0x0E7F - 0x0E00 + 1).Select(c => (char)c)).ToArray();

                if (safeFileTemp.Any(c => !AllowedCharacters.Contains(c)))
                {
                    return ("Input file name contains invalid characters.");
                }

                // ตรวจสอบนามสกุลไฟล์ว่าถูกต้องหรือไม่
                string extension = Path.GetExtension(safeFileTemp).ToLowerInvariant();
                if (string.IsNullOrEmpty(extension))
                {
                    return ("File does not have a valid extension.");
                }

                // อนุญาตเฉพาะไฟล์ที่กำหนด 
                string[] allowedExtensions = { ".xlsx", ".xls", ".pdf", ".doc", ".docx", ".png", ".jpg", ".gif", ".eml", ".msg" };
                if (!allowedExtensions.Contains(extension))
                {
                    return ("Invalid file type.");
                }

                // สร้างเส้นทางตามไฟล์ต้นทาง
                string finalRootDir = "";
                string templatewwwwRootDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot");
                if (!Directory.Exists(templatewwwwRootDir))
                {
                    throw new DirectoryNotFoundException("Folder directory not found.");
                }

                string templateRootDir = Path.Combine(templatewwwwRootDir, "AttachedFileTemp");
                if (!Directory.Exists(templateRootDir))
                {
                    throw new DirectoryNotFoundException("Folder directory not found.");
                }

                // ตรวจสอบไดเร็กทอรี file ภายใน module --> /AttachedFileTemp/folderxxx/
                string templateDir = Path.Combine(templateRootDir, folder);
                templateDir = Path.GetFullPath(templateDir);
                if (!templateDir.StartsWith(templateRootDir, StringComparison.OrdinalIgnoreCase))
                {
                    return ("Temp directory is outside of the Template directory.");
                }

                finalRootDir = templateDir;
                if (string.IsNullOrEmpty(finalRootDir))
                {
                    return ("Invalid finalRootDir.");
                }

            }
            catch (Exception ex)
            {
                return $"An error occurred while processing your request.{ex.Message.ToString()}";
            }

            return "";
        }
        public static string check_format_file_name(string file_name)
        {
            try
            {
                if (string.IsNullOrEmpty(file_name))
                {
                    return ("Invalid file file name.");
                }

                string safeFileTemp = $"{file_name}";
                if (string.IsNullOrEmpty(safeFileTemp))
                {
                    return ("Invalid file safe file temp.");
                }
                // ตรวจสอบ Format ชื่อไฟล์
                // เพิ่มช่วงของตัวอักษรภาษาไทยอยู่ในช่วง 0x0E00 ถึง 0x0E7F (หรือ เ ถึง ๏)
                char[] AllowedCharacters = "()abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-. ".ToCharArray()
                .Concat(Enumerable.Range(0x0E00, 0x0E7F - 0x0E00 + 1).Select(c => (char)c)).ToArray();

                if (safeFileTemp.Any(c => !AllowedCharacters.Contains(c)))
                {
                    return ("Input file name contains invalid characters.");
                }
            }
            catch (Exception ex)
            {
                return $"An error occurred while processing your request.{ex.Message.ToString()}";
            }


            return "";
        }

    }
}
