using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading.Tasks;

namespace OpenXmlDemo.Api.Help
{
    public static class FileHelper
    {
        private readonly static string _rootPath = Directory.GetCurrentDirectory();

        /// <summary>
        /// 下载图片至本地
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public static async Task<string> Down(string url)
        {
            try
            {
                WebRequest webRequest = WebRequest.Create(url);
                using HttpWebResponse response = (HttpWebResponse)await webRequest.GetResponseAsync();
                Stream stream = response.GetResponseStream();
                Image img = Image.FromStream(stream);
                string localPath = Path.Combine(_rootPath, "wwwroot", "Images");
                if (!Directory.Exists(localPath))
                    Directory.CreateDirectory(localPath);
                localPath = Path.Combine(localPath, Path.GetFileName(url));
                try
                {
                    if (File.Exists(localPath))
                        File.Delete(localPath);
                }
                catch (Exception)
                {
                    return localPath;
                }
                string imgType = url.Split('.').Last().ToLower();
                switch (imgType)
                {
                    case "png":
                        img.Save(localPath, ImageFormat.Png);
                        break;
                    case "gif":
                        img.Save(localPath, ImageFormat.Gif);
                        break;
                    case "icon":
                        img.Save(localPath, ImageFormat.Icon);
                        break;
                    case "jpg" or "jpeg":
                    default:
                        img.Save(localPath, ImageFormat.Jpeg);
                        break;
                }

                await stream.DisposeAsync();
                img.Dispose();
                return localPath;
            }
            catch (Exception)
            {
                return "";
            }

        }
    }
}
