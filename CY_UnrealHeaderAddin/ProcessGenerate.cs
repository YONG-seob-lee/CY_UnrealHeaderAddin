using System;
using System.CodeDom.Compiler;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace CY_UnrealHeaderAddin
{
    public class GenerateProgram : Process
    {

        public void Executor(string unrealbuildtoolfolderpath, string uprojectpath)
        {
            ProcessStartInfo StartInfo = new ProcessStartInfo
            {
                FileName = "CMD.exe",
                CreateNoWindow = false,
                UseShellExecute = false,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            Process Process = new Process();

            Process.StartInfo = StartInfo;
            Process.Start();
            
            Process.StandardInput.WriteLine(@"D:" + Environment.NewLine);
            Process.StandardInput.WriteLine(@"cd " + unrealbuildtoolfolderpath + Environment.NewLine);
            Process.StandardInput.WriteLine(@"UnrealBuildTool.exe -projectfiles -project=" + "\"" + uprojectpath + "\" -game -rocket -progress" + Environment.NewLine);
            Process.StandardInput.Close();

            string resultvalue = Process.StandardOutput.ReadToEnd();
            Process.WaitForExit();
            Process.Close();

            MessageBox.Show(resultvalue);
        }

        public string FindFolderPath(string rootFolder)
        {
            string UnrealName = "UE_5.1";
            string folderName = "UnrealBuildTool";
            string found = null;

            if(Directory.Exists(rootFolder) == false)
            {
                return found;
            }

            DirectoryInfo rootDirectory = Directory.CreateDirectory(rootFolder);

            Local_FindDirectoryAbsPath(rootDirectory);
            return found;

            // 내부 재귀 메소드
            void Local_FindDirectoryAbsPath(DirectoryInfo currentDirectory)
            {
                if (found != null) return;

                // 일치하는 폴더명 찾은 경우
                if (currentDirectory.Name == folderName)
                {
                    found = currentDirectory.FullName;
                    return;
                }

                // 하위 폴더들 재귀 탐색
                DirectoryInfo[] subFolders = currentDirectory.GetDirectories();
                foreach (var folder in subFolders)
                {
                    Local_FindDirectoryAbsPath(folder);
                }
            }
        }

        public string FinduprojectPath()
        {
            string fileName = "ProjectCY.uproject";
            AddinFunctionLibrary Lib = new AddinFunctionLibrary();
            string EmplacePath = Lib.GetEmplacePath();

            EmplacePath = CommonUtil.ApartFolder(EmplacePath);
            EmplacePath = CommonUtil.ApartFolder(EmplacePath);

            DirectoryInfo directoryInfo = Directory.CreateDirectory(EmplacePath);
            string found = null;

            Local_FindDirectoryAbsPath(directoryInfo);
            return found;

            void Local_FindDirectoryAbsPath(DirectoryInfo currentDirectory)
            {
                if(found != null)
                {
                    return;
                }

                FileInfo[] files = currentDirectory.GetFiles();
                foreach(var file in files)
                {
                    if(file.Name == fileName)
                    {
                        found = file.FullName;
                        return;
                    }
                }

                DirectoryInfo[] subFolders = currentDirectory.GetDirectories();
                foreach (var folder in subFolders)
                {
                    Local_FindDirectoryAbsPath(folder);
                }
            }
        }
    }
}