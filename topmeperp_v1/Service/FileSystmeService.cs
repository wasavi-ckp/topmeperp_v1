using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ICSharpCode.SharpZipLib.Zip;
using log4net;
using System.IO;
using System.IO.Compression;
using topmeperp.Models;
using System.Collections;

namespace topmeperp.Service
{
    /// <summary>
    /// Zip & Unzip 工具
    /// </summary>
    public class ZipFileCreator
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        //讀取目錄下所有檔案
        private static ArrayList GetFiles(string path)
        {
            logger.Debug("get all file in " + path);
            ArrayList files = new ArrayList();

            if (Directory.Exists(path))
            {
                files.AddRange(Directory.GetFiles(path));
            }

            return files;
        }
        //建立目錄
        public static void CreateDirectory(string path)
        {
            if (!Directory.Exists(path))
            {
                logger.Debug("CreateDirectory " + path);
                Directory.CreateDirectory(path);
            }
        }
        public string ZipDirectory(string path)
        {
            string zipPath = path + @"..\" + "空白詢價單.zip";
            ICSharpCode.SharpZipLib.Zip.FastZip z = new ICSharpCode.SharpZipLib.Zip.FastZip();
            z.CreateEmptyDirectories = true;
            z.CreateZip(zipPath, path , true, "");
            return zipPath;
        }
        public static void DelDirectory(string path)
        {
            try
            {
                logger.Debug("clear directory: " + path);
                Directory.Delete(path, true);
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
            }
        }
        public string ZipFiles(string path, string password, string comment)
        {
            logger.Debug("zip file in : " + path);
            string zipPath = "";
            ZipOutputStream zos = null;
            try
            {
                zipPath = path + @"\" + Path.GetFileName(path) + comment + ".zip";
                ArrayList files = GetFiles(path);
                zos = new ZipOutputStream(File.Create(zipPath));
                if (password != null && password != string.Empty) zos.Password = password;
                if (comment != null && comment != "") zos.SetComment(comment);
                zos.SetLevel(9);//Compression level 0-9 (9 is highest)
                byte[] buffer = new byte[4096];

                foreach (string f in files)
                {
                    ZipEntry entry = new ZipEntry(Path.GetFileName(f));
                    entry.DateTime = DateTime.Now;
                    zos.PutNextEntry(entry);
                    FileStream fs = File.OpenRead(f);
                    int sourceBytes;

                    do
                    {
                        sourceBytes = fs.Read(buffer, 0, buffer.Length);
                        zos.Write(buffer, 0, sourceBytes);
                    } while (sourceBytes > 0);

                    fs.Close();
                    fs.Dispose();
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            finally
            {
                if (zos != null)
                {
                    zos.Finish();
                    zos.Close();
                    zos.Dispose();
                }
            }
            return zipPath;
        }

        ///
        /// 解壓縮檔案
        ///
        ///解壓縮檔案目錄路徑
        ///密碼
        public void UnZipFiles(string path, string password)
        {
            ZipInputStream zis = null;

            try
            {
                string unZipPath = path.Replace(".zip", "");
                CreateDirectory(unZipPath);
                zis = new ZipInputStream(File.OpenRead(path));
                if (password != null && password != string.Empty) zis.Password = password;
                ZipEntry entry;

                while ((entry = zis.GetNextEntry()) != null)
                {
                    string filePath = unZipPath + @"\" + entry.Name;

                    if (entry.Name != "")
                    {
                        FileStream fs = File.Create(filePath);
                        int size = 2048;
                        byte[] buffer = new byte[2048];
                        while (true)
                        {
                            size = zis.Read(buffer, 0, buffer.Length);
                            if (size > 0) { fs.Write(buffer, 0, size); }
                            else { break; }
                        }

                        fs.Close();
                        fs.Dispose();
                    }
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            finally
            {
                zis.Close();
                zis.Dispose();
            }
        }
    }
}