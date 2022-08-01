using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DonetInstall
{
    internal class InstallDotnet
    {
        protected Version version = new Version();
        protected readonly Dictionary<string, string> installedDic = new Dictionary<string, string>();

        /// <summary>
        /// 已安装的.NET CORE版本
        /// </summary>
        protected readonly Dictionary<string, string> installedNetCoreDic = new Dictionary<string, string>();

        public InstallDotnet()
        {
            Console.WriteLine("正在获取本机已安装版本");
            Get1To45VersionFromRegistry(installedDic);
            Get45PlusFromRegistry(installedDic);
            GetNetCore(installedDic);
            ShowInstall();
        }

        /// <summary>
        /// 获取.NetCore的安装情况
        /// </summary>
        /// <param name="dic"></param>
        private void GetNetCore(IDictionary<string, string> dic)
        {
            string text = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            if (string.IsNullOrEmpty(text))
            {
                return;
            }
            text += "\\dotnet\\shared";
            if (!Directory.Exists(text))
            {
                return;
            }
            DirectoryInfo[] directories = new DirectoryInfo(text).GetDirectories();
            for (int i = 0; i < directories.Length; i++)
            {
                DirectoryInfo[] directories2 = directories[i].GetDirectories();
                for (int j = 0; j < directories2.Length; j++)
                {
                    DirectoryInfo directoryInfo = directories2[j];
                    dic[directoryInfo.Name] = directoryInfo.Name;
                    installedNetCoreDic[directoryInfo.Name] = directoryInfo.Name;
                }
            }
        }

        /// <summary>
        /// 获取.NET 4.5以上版本安装情况
        /// </summary>
        /// <param name="dic"></param>
        private void Get45PlusFromRegistry(IDictionary<string, string> dic)
        {
            const string subkey = @"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\";

            using (RegistryKey ndpKey = Registry.LocalMachine.OpenSubKey(subkey))
            {
                if (ndpKey == null)
                    return;
                //First check if there's an specific version indicated
                if (ndpKey.GetValue("Version") != null)
                {
                    dic.Add(ndpKey.GetValue("Version").ToString(), ndpKey.GetValue("Version").ToString());
                }
                else
                {
                    if (ndpKey != null && ndpKey.GetValue("Release") != null)
                    {
                        var temp = CheckFor45PlusVersion(
                                    (int)ndpKey.GetValue("Release")
                                );
                        if (string.IsNullOrEmpty(temp))
                            dic.Add(temp, temp);
                    }
                }

                // Checking the version using >= enables forward compatibility.
                string CheckFor45PlusVersion(int releaseKey)
                {
                    if (releaseKey >= 528040)
                        return "4.8";
                    if (releaseKey >= 461808)
                        return "4.7.2";
                    if (releaseKey >= 461308)
                        return "4.7.1";
                    if (releaseKey >= 460798)
                        return "4.7";
                    if (releaseKey >= 394802)
                        return "4.6.2";
                    if (releaseKey >= 394254)
                        return "4.6.1";
                    if (releaseKey >= 393295)
                        return "4.6";
                    if (releaseKey >= 379893)
                        return "4.5.2";
                    if (releaseKey >= 378675)
                        return "4.5.1";
                    if (releaseKey >= 378389)
                        return "4.5";
                    // This code should never execute. A non-null release key should mean
                    // that 4.5 or later is installed.
                    return "";
                }
            }
        }

        public void ShowInstall()
        {
            foreach (var item in installedDic)
            {
                var value = item.Value;
                Console.WriteLine(value);
            }
            ShowLatestVersion();
            Console.WriteLine("请输入.NetCore 版本进行安装，请勿重复安装");
            InstallNew();
        }

        private void ShowLatestVersion()
        {
            Console.WriteLine("正在获取当前.NetCore 最新版本");
            GetLatestVersionFromUrl("3.1");
            GetLatestVersionFromUrl("6.0");
            GetLatestVersionFromUrl("7.0");
        }

        private void InstallNew()
        {
            try
            {
                var version = Console.ReadLine();
                if (string.IsNullOrEmpty(version))
                    InstallNew();
                InstallNetCore(version);
            }
            catch (WebException webex)
            {
                if (webex.Status == WebExceptionStatus.ProtocolError && webex.Response != null)
                {
                    var resp = (HttpWebResponse)webex.Response;

                    if (resp.StatusCode == HttpStatusCode.NotFound)

                    {
                        Console.WriteLine("版本号不存在，请检查安装信息和版本号后重试");
                        InstallNew();
                    }
                }
                else
                {
                    Console.WriteLine("安装异常，请检查安装信息和版本号后重试");
                    InstallNew();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("安装异常，请检查安装信息和版本号后重试");
                InstallNew();
            }
        }

        /// <summary>
        /// 获取.NET FrameWork1到 4.5 的安装情况
        /// </summary>
        /// <param name="dic"></param>
        private void Get1To45VersionFromRegistry(IDictionary<string, string> dic)
        {
            using (RegistryKey ndpKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\"))
            {
                foreach (string versionKeyName in ndpKey.GetSubKeyNames())
                {
                    // Skip .NET Framework 4.5 version information.
                    if (versionKeyName == "v4")
                    {
                        continue;
                    }

                    if (versionKeyName.StartsWith("v"))
                    {
                        RegistryKey versionKey = ndpKey.OpenSubKey(versionKeyName);
                        // Get the .NET Framework version value.
                        string name = (string)versionKey.GetValue("Version", "");
                        // Get the service pack (SP) number.
                        string sp = versionKey.GetValue("SP", "").ToString();

                        // Get the installation flag, or an empty string if there is none.
                        string install = versionKey.GetValue("Install", "").ToString();
                        if (string.IsNullOrEmpty(install)) // No install info; it must be in a child subkey.
                        { }
                        else
                        {
                            if (!(string.IsNullOrEmpty(sp)) && install == "1")
                            {
                                dic.Add(name, name);
                            }
                        }
                        if (!string.IsNullOrEmpty(name))
                        {
                            continue;
                        }
                        foreach (string subKeyName in versionKey.GetSubKeyNames())
                        {
                            RegistryKey subKey = versionKey.OpenSubKey(subKeyName);
                            name = (string)subKey.GetValue("Version", "");
                            if (!string.IsNullOrEmpty(name))
                                sp = subKey.GetValue("SP", "").ToString();

                            install = subKey.GetValue("Install", "").ToString();
                            if (string.IsNullOrEmpty(install)) //No install info; it must be later.
                                dic.Add(name, name);
                            else
                            {
                                if (!(string.IsNullOrEmpty(sp)) && install == "1")
                                {
                                    dic.Add(name, sp);
                                }
                                else if (install == "1")
                                {
                                    dic.Add(name, name);
                                }
                            }
                        }
                    }
                }
            }
        }

        private string fileName;
        private ProgressBar progressBar;

        /// <summary>
        /// 自动安装.NET CORE
        /// </summary>
        /// <param name="defaultVersion"></param>
        private void InstallNetCore(string defaultVersion = "6.0.7")
        {
            Version version = new Version();
            GetNetCore(installedDic);
            if (installedNetCoreDic.Count > 0)
            {
                if (installedNetCoreDic.ContainsValue(defaultVersion))
                {
                    Console.WriteLine($"已安装版本：{defaultVersion}");
                    return;
                }
            }

            Console.WriteLine($"InstallNetCore: {defaultVersion}");
            string weburl = $"https://dotnet.microsoft.com/zh-cn/download/dotnet/thank-you/runtime-aspnetcore-{defaultVersion}-windows-hosting-bundle-installer";
            Console.WriteLine("正在获取下载地址");
            var urlTextTemp = GetUrlHtml(weburl);

            var downUrl = GetDowloadUrl(urlTextTemp);
            Console.WriteLine($"解析下载地址成功：{downUrl}");

            fileName = Path.Combine(Path.GetTempPath(), Path.GetFileName(downUrl));
            if (!File.Exists(fileName))
            {
                Console.WriteLine("正在下载：{0}", downUrl);
                progressBar = new ProgressBar();
                var client = new WebClient();
                client.DownloadFileCompleted += client_DownloadFileCompleted;
                client.DownloadProgressChanged += client_DownloadProgressChanged;
                client.DownloadFileAsync(new Uri(downUrl), fileName);
            }
            else
            {
                ProcessInstall();
            }
        }

        private void ProcessInstall()
        {
            Console.WriteLine("正在安装：{0}", fileName);
            if (Process.Start(fileName, "/quiet").WaitForExit(40000))
            {
                Console.WriteLine("安装成功！");
                Environment.ExitCode = 0;
                return;
            }
            Console.WriteLine("安装超时！");
            Environment.ExitCode = 1;
        }

        //下载完成
        private void client_DownloadFileCompleted(object sender, AsyncCompletedEventArgs e)
        {
            Console.WriteLine("");
            Console.WriteLine($"{fileName}文件下载完成");
            Console.WriteLine("");
            ProcessInstall();
        }

        //下载进度
        private void client_DownloadProgressChanged(object sender, DownloadProgressChangedEventArgs e)
        {
            progressBar.Dispaly(e.ProgressPercentage);
        }

        private string GetUrlHtml(string url)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.ContentType = "application/json";
            request.Accept = "application/json,text/javascript,*/*,q=0.01";
            request.Headers.Add("Accept-Encoding", "deflate,gzip");
            request.AutomaticDecompression = DecompressionMethods.Deflate | DecompressionMethods.GZip | DecompressionMethods.None;

            using (Stream stream = request.GetResponse().GetResponseStream())
            {
                StreamReader sr = new StreamReader(stream);
                return sr.ReadToEnd();
            }
        }

        private string GetDowloadUrl(string temp)
        {
            var start = "https://download.visualstudio";
            var end = ".exe";
            var reg = new Regex($"(?<=({start}))[.\\s\\S]*?(?=({end}))");
            return start + reg.Match(temp).Value + end;
        }

        private string GetLatestVersionFromUrl(string mainVersion)
        {
            try
            {
                var url = $"https://dotnet.microsoft.com/zh-cn/download/dotnet/{mainVersion}";
                var temp = GetUrlHtml(url);
                var version = GetLatestVersion(temp);
                Console.WriteLine($".NET CORE {mainVersion}最新版本为：{version}");
                return version;
            }
            catch (WebException webex)
            {
                if (webex.Status == WebExceptionStatus.ProtocolError && webex.Response != null)
                {
                    var resp = (HttpWebResponse)webex.Response;

                    if (resp.StatusCode == HttpStatusCode.NotFound)

                    {
                        Console.WriteLine($"获取.NET CORE {mainVersion}最新版本没有成功");
                    }
                }
                else
                {
                    Console.WriteLine($"获取.NET CORE {mainVersion}最新版本没有成功");
                }
                return "";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取.NET CORE {mainVersion}最新版本没有成功");
                return "";
            }
        }

        /// <summary>
        /// 获取最新版本
        /// </summary>
        /// <param name="temp"></param>
        /// <returns></returns>
        private string GetLatestVersion(string temp)
        {
            var start = "#version_0";
            var end = "</button>";
            var reg = new Regex($"(?<=({start}))[.\\s\\S]*?(?=({end}))");
            return reg.Match(temp).Value.Substring(2);
        }
    }
}