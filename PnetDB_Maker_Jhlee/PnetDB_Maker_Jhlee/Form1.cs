using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;                                           //COM for excel file handling

namespace PnetDB_Maker_Jhlee
{
    public partial class Form1 : Form
    {

        bool isFullListBox = false;// 한번 ListBox채우면 더 못채우도록
        string Path = null;
        string InputFile_Path = null;
        object[,] ReadData; // 읽어온  Mtp Io 리스트
        string OutputFile_1 = "\\PNet_Compoints.xlsx";
        string[] Row1_Text = new string[] { "BasicSystemName", "PointName", "StationNum", "PointIndication", "CurrentValue", "Unit", "HistoryLibraryCollectionPeriod", "ProjectCalculationAttribute", "RangeUpperLimit", "RangeBottomLimit" };
        string[] PC_ID = new string[] { "SC1P", "SC1S", "SD1P", "SD1S", "CB2P", "CB2S", "CD2P", "CD2S", "SB2P", "SB2S", "CA2P", "CA2S", "CC2P", "CC2S", "SA2P", "SA2S", "CB1P", "CB1S", "CD1P", "CD1S", "SB1P", "SB1S", "CA1P", "CA1S", "CC1P", "CC1S", "SA1P", "SA1S" };
     
        uint AO_TagCnt = 0;
        uint AI_TagCnt = 0;
        uint DO_TagCnt = 0;
        uint DI_TagCnt = 0;

        Excel.Application Excel_App = null;
        Excel.Workbook Rd_Wb = null;
        Excel.Worksheet Rd_Ws = null;

        Excel.Workbook Wr_Wb = null;

        Excel.Worksheet AO_WS = null;
        Excel.Worksheet AI_WS = null;
        Excel.Worksheet DO_WS = null;
        Excel.Worksheet DI_WS = null;
        Excel.Worksheet GWAI_WS = null;


        public Form1()
        {
            InitializeComponent();
        }

        private void listView1_DragDrop(object sender, DragEventArgs e)// 이벤트 함수는 리스트에 파일의 제목을 추가하는 내용임
        {
            if(false ==isFullListBox)
            {
                string[] Files = (string[])e.Data.GetData(DataFormats.FileDrop);

                foreach (string File in Files)
                {
                   listBox1.Items.Add(File);
                   InputFile_Path += File;
                }

                
                isFullListBox = true;

            }
        }

        private void listView1_DragEnter(object sender, DragEventArgs e)// 이거 안되어있으면, 파일이 리스트에 추가되지않는것처럼 보임 
        {
            e.Effect = DragDropEffects.Copy;

        }

        private void button1_MouseClick(object sender, MouseEventArgs e)// RunButton Inputfile이 없으면 아무것도 실행안된다.
        {
            if (InputFile_Path == null)
            {
                MessageBox.Show("IO List를 드래그해주십시오.");
                return;
            }

            Read_InputFile();
            Write_PnetCompoints();

            //버튼1을 누르면 Excel에 데이터를 읽어들여서(Read_InputFile()-> PnetCompoint.xlxs 형태로 만든다


        }

        private void Write_PnetCompoints()
        {

            //시트만들고
            //1행 :BasicSystemName	PointName	StationNum	PointIndication	CurrentValue	Unit	HistoryLibraryCollectionPeriod	ProjectCalculationAttribute	RangeUpperLimit	RangeBottomLimit

            try
            {
                Path += OutputFile_1;

                Excel_App = new Excel.Application();
                Wr_Wb = Excel_App.Workbooks.Add();
       
                AO_WS = Wr_Wb.Worksheets.Item["Sheet1"];
                AO_WS.Name = "AO";
                int i = 1;

                foreach(string s in Row1_Text)
                {
                    AO_WS.Cells[1, i] = s;
                    ++i;
                }

                int Dest_Index = 2;
                int Sour_Index = 2;
                foreach (string pc_id in PC_ID)
                {
                    while (true)
                    {

                        if (ReadData[Sour_Index, 3] == null)
                            break;

                        string TagName = ReadData[Sour_Index, 3].ToString();

                        if (true != TagName.StartsWith("AI"))
                            break;


                        AO_WS.Cells[Dest_Index, 1] = pc_id;          //BasicSystem Name

                        if (ReadData[Sour_Index, 2] == null)
                        {
                            AO_WS.Cells[Dest_Index, 2] = ""; // Point Name
                        }
                        else
                        {
                            AO_WS.Cells[Sour_Index, 2] = ReadData[Sour_Index, 2]; // Point Name
                        }

                        AO_WS.Cells[Dest_Index, 3] = 1; // StationNum

                        if (ReadData[Sour_Index, 4] == null)
                        {
                            AO_WS.Cells[Dest_Index, 4] = "";
                        }
                        else
                        {
                            AO_WS.Cells[Dest_Index, 4] = ReadData[Sour_Index, 4];//PointIndication
                        }

                        AO_WS.Cells[i, 5] = 0;// currentValue

                        if (ReadData[Sour_Index, 11] == null)
                        {
                            AO_WS.Cells[Dest_Index, 6] = "-";
                        }
                        else
                        {
                            AO_WS.Cells[Dest_Index, 6] = ReadData[Sour_Index, 11];//Unit
                        }

                        AO_WS.Cells[Dest_Index, 7] = 1;//HistoryLibrary
                        AO_WS.Cells[Dest_Index, 8] = 0;//ProjectCalculationAttribute


                        if (ReadData[Sour_Index, 10] == null)
                        {
                            AO_WS.Cells[Dest_Index, 9] = "";
                            AO_WS.Cells[Dest_Index, 10] = "";
                        }
                        else
                        {
                            string[] temp = ReadData[Sour_Index, 10].ToString().Split('~');

                            AO_WS.Cells[Dest_Index, 9] = temp[1];
                            AO_WS.Cells[Dest_Index, 10] =temp[0];
                        }
                        ++Dest_Index;
                        ++Sour_Index;
                    }
                    Sour_Index = 2;

                }
               
                
              
               


                AI_WS = Wr_Wb.Worksheets.Add(After: AO_WS);
                AI_WS.Name = "AI";

                DO_WS = Wr_Wb.Worksheets.Add(After: AI_WS);
                DO_WS.Name = "DO";

                DI_WS = Wr_Wb.Worksheets.Add(After: DO_WS);
                DI_WS.Name = "DI";

                GWAI_WS = Wr_Wb.Worksheets.Add(After: DI_WS);
                GWAI_WS.Name = "GWAI";


        

                //파일 xlsx 로 저장
                Wr_Wb.SaveAs(@Path, Excel.XlFileFormat.xlOpenXMLWorkbook);
                Wr_Wb.Close(true);
                Excel_App.Quit();
            }

            finally
            {
                ReleaseExcelObject(Rd_Ws);
                ReleaseExcelObject(Rd_Wb);
                ReleaseExcelObject(Excel_App);

            }
        }

        private void Read_InputFile()
        {
            try
            {
                Excel_App = new Excel.Application();
                Rd_Wb = Excel_App.Workbooks.Open(@InputFile_Path);// 해당 경로에 워크북 열기
                Rd_Ws = Rd_Wb.Worksheets.get_Item(1) as Excel.Worksheet;
              
                Excel.Range rng = Rd_Ws.UsedRange;
                
                ReadData = rng.Value;

                Rd_Wb.Close(true);
                Excel_App.Quit();
            }

            finally
            {
                ReleaseExcelObject(Rd_Ws);
                ReleaseExcelObject(Rd_Wb);
                ReleaseExcelObject(Excel_App);
            }
        }
        private static void ReleaseExcelObject(object obj)
        {
            try
            {

                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);
                    obj = null;

                }
            }

            catch (Exception ex)
            {
                obj = null;
                throw ex;

            }
            finally
            {
                GC.Collect();

            }

        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            Path = System.IO.Directory.GetCurrentDirectory();// Form이 열리자마자, 폴더 실행경로 받아오기



        }
    }
}
