﻿using System;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Data;
using System.Collections.Generic;
using System.Linq;

namespace CY_UnrealHeaderAddin
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion

        public void RegistPath(string path)
        {
            PrimitivePath = path;
        }

        public void AddCsvCheckButton()
        {
            bAddCsv = !bAddCsv;
        }

        public void AddHeader()
        { 
            if(PrimitivePath.Length == 0)
            {
                return;
            }
        }

        string PrimitivePath;
        bool bAddCsv = false;
    }

    public partial class AddinFunctionLibrary
    {
        public enum EMessageType
        {
            None = 0,
            AccessRegist = 1,
            NoRegistError = 2,
            MoreThanOneExcelIsOpen = 3,
            BlankData = 4,
        }

        public const string UneralHeaderAddin = "UnrealHeaderAddin";
        public const string RegistPath = "RegistPath";

        public string MakeRegistPath(string OriginPath)
        {
            string ResultPath = OriginPath + "\\" + UneralHeaderAddin;
            Directory.CreateDirectory(ResultPath);

            ResultPath += "\\" + RegistPath + ".txt";
            return ResultPath;
        }
        public string GetEmplacePath()
        {
            StreamReader Reader = new StreamReader(GetRegistFilePath());
            if(Reader == null)
            {
                return string.Empty;
            }
            return Reader.ReadLine();
        }
        public string GetRegistPath()
        {
            string Path = Directory.GetCurrentDirectory() + "\\" + UneralHeaderAddin;
            if(Directory.Exists(Path))
            {
                return Path;
            }

            return string.Empty;
        }
        public string GetRegistFilePath()
        {
            string FilePath = GetRegistPath() + "\\" + RegistPath + ".txt";
            if(File.Exists(FilePath))
            {
                return FilePath;
            }

            return string.Empty;
        }

        public bool CheckRegistDirectory()
        {
            string CurrentDirectory = Directory.GetCurrentDirectory();

            return true;
        }
        private Excel.Workbook GetCurrentWorkbook()
        {
            Process[] ProcessName = Process.GetProcessesByName("excel");
            if (ProcessName.Length != 1)
            {
                return null;
            }

            Excel.Application App = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("excel.application");
            return App.ActiveWorkbook;
        }

        private Dictionary<string, string> GetTablePropertyData()
        {
            Dictionary<string, string> StructData = new Dictionary<string, string>();

            Excel.Workbook CurrentWorkbook = GetCurrentWorkbook();
            Excel.Worksheet WorkSheet = CurrentWorkbook.ActiveSheet;
            Excel.Range ExcelRange = WorkSheet.UsedRange;

            int colCount = ExcelRange.Columns.Count;

            for (int i = 1; i <= colCount; i++)
            {
                StructData.Add(ExcelRange.Cells[1, i].Value2.ToString(), ExcelRange.Cells[2, i].Value2.ToString());
            }
           

            return StructData;
        }
        
        public string ApartExtension(string Name)
        {
            Int32 PointIndex = 0;

            char[] NameChar = Name.ToCharArray();
            for(int i = NameChar.Length - 1; i != 0 ; i--)
            {
                if(NameChar[i] == '.')
                {
                    PointIndex = i;
                    break;
                }
            }

            return Name.Substring(0, PointIndex);
        }

        public void MakeCsv()
        {
            Excel.Workbook Workbook = GetCurrentWorkbook();
            if(Workbook == null)
            {
                ShowMessage(EMessageType.MoreThanOneExcelIsOpen, string.Empty);
                return;
            }

            System.Data.DataTable DataTable = new System.Data.DataTable();
            string SavePath = string.Empty;
            
            try
            {
                Process[] ProcessName = Process.GetProcessesByName("excel");
                if(ProcessName.Length > 1) 
                {
                    ShowMessage(EMessageType.MoreThanOneExcelIsOpen, string.Empty);
                    return;
                }

                SavePath = Workbook.Path;

                Excel.Worksheet WorkSheet = Workbook.ActiveSheet;
                Excel.Range ExcelRange = WorkSheet.UsedRange;

                int rowCount = ExcelRange.Rows.Count;
                int colCount = ExcelRange.Columns.Count;

                if(rowCount == 1 && colCount == 1)
                {
                    ShowMessage(EMessageType.BlankData, string.Empty);
                    return;
                }

                for (int i = 1; i <= rowCount; i++)
                {
                    for (int j = 1; j <= colCount; j++)
                    {
                        DataTable.Columns.Add(ExcelRange.Cells[i, j].Value2.ToString());
                    }
                    break;
                }

                int rowCounter;
                DataRow Row = null;
                for (int i = 2; i < rowCount; i++)
                {
                    Row = DataTable.NewRow();
                    rowCounter = 0;
                    for (int j = 1; j <= colCount; j++)
                    {
                        if (ExcelRange.Cells[i, j] != null && ExcelRange.Cells[i, j].Value2 != null)
                        {
                            if (j == 1)
                            {
                                var Value = ExcelRange.Cells[i, j].Value2;
                                Row[rowCounter] = Value;
                            }
                            else
                            {
                                Row[rowCounter] = ExcelRange.Cells[i, j].Value2.ToString();
                            }
                        }
                        else
                        {
                            Row[i] = "";
                        }
                        rowCounter++;
                    }
                    DataTable.Rows.Add(Row);
                }

                Workbook.SaveAs(SavePath + "\\" + "asdf.csv", XlFileFormat.xlCSV);

            }
            finally
            {
                //Workbook.Close(true);
                //App.Quit();

                //ReleaseExcelObject(Workbook);
                //ReleaseExcelObject(App);
            }
        }

        public void MakeHeader()
        {
            Excel.Workbook Workbook = GetCurrentWorkbook();
            if (Workbook == null)
            {
                ShowMessage(EMessageType.MoreThanOneExcelIsOpen, string.Empty);
                return;
            }

            const string BlankStr = "    ";
            const string UPROPERTY = "   UPROPERTY(EditAnywhere)";
            
            string ExcelName = ApartExtension(Workbook.Name);
            
            String RegistPath = GetEmplacePath();

            FileStream FileStream = new FileStream(RegistPath + "\\" + ExcelName + ".h", FileMode.OpenOrCreate);
            StreamWriter Writer = new StreamWriter(FileStream);

            Writer.WriteLine("#pragma once");
            Writer.WriteLine(BlankStr);
            Writer.WriteLine("include \"CoreMinimal.h\"");
            Writer.WriteLine("include \"Engine/DataTable.h\"");
            Writer.WriteLine("include " + ExcelName + ".generated.h");
            Writer.WriteLine(BlankStr);
            Writer.WriteLine("USTRUCT()");
            Writer.WriteLine("struct F" + ExcelName + ": public FTableRowBase");
            Writer.WriteLine("{");
            Writer.WriteLine(BlankStr + "GENERATED_USTRUCT_BODT()");
            Writer.WriteLine("public:");

            //여기부터 구조체 for문
            Dictionary<string, string> StructData = GetTablePropertyData();

            List<string> Keys = StructData.Keys.ToList();
            List<string> Values = StructData.Values.ToList();
            for(int i = 0; i != StructData.Count; i++)
            {
                Writer.WriteLine(UPROPERTY);

                string InitializeStr = string.Empty;
                switch(Values[i])
                {
                    case "int32":
                        InitializeStr = " = 0;";
                        break;
                    case "Float":
                        InitializeStr = " = 0.f;";
                        break;
                    case "FString":
                        InitializeStr = " = FString();";
                        break;
                    case "FName":
                        InitializeStr = " = FName();";
                        break;
                    case "bool":
                        InitializeStr = " = false;";
                        break;
                    default:
                        InitializeStr = ";";
                        break;
                }

                Writer.WriteLine(BlankStr + Keys[i] + " " + Values[i] + InitializeStr);
            }

            Writer.WriteLine("};");
            Writer.Close();
            FileStream.Close();
        }

        public void ReleaseExcelObject(object Object)
        {
            try
            {
                if(Object != null)
                {
                    Marshal.ReleaseComObject(Object);
                    Object = null;
                }
            }
            catch (Exception ex)
            {
                Object = null;
                throw ex;
            }
            finally 
            { 
                GC.Collect();
            }
        }

        public void ShowMessage(EMessageType Type, string String)
        {
            string Title = string.Empty;
            string Discussion = string.Empty;
            const string Warning = "Warning : ";

            switch (Type)
            {
                case EMessageType.AccessRegist:
                    Title = "Access Regist Path : 경로 설정 완료";
                    Discussion = "경로 :" + String + "\n 해당 경로로 설정 되었습니다." + "\n Save Directory : \"" + Directory.GetCurrentDirectory() + "\"";
                    break;
                case EMessageType.NoRegistError:
                    Title = Warning + "경로 미설정";
                    Discussion = "경로를 설정하지 않으셨습니다.\n경로를 설정해 주세요.\n" + "경로 :" + String + "\n 해당 경로를 확인해주세요.";
                    break;
                case EMessageType.MoreThanOneExcelIsOpen:
                    Title = Warning + "다수의 엑셀 파일 오픈";
                    Discussion = "두개 이상의 엑셀 파일이 활성화 되어있습니다..\n만약 한개만 켜져있는데도 이 메세지가 노출된다면\n컴퓨터를 재부팅해주세요\n" 
                                  + "해결책 : 한개의 엑셀파일만 열려있어야 합니다.";
                    break;
                case EMessageType.BlankData:
                    Title = Warning + "엑셀 내용 공백";
                    Discussion = "최소 한개 이상의 데이터를 넣고\nAddCsv or AddHeader 를 해 주세요.";
                    break;
                default:
                    break;
            }

            MessageBox.Show(Discussion, Title);
        }
    }
}
