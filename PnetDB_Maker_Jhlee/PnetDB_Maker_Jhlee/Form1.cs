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
using Excel = Microsoft.Office.Interop.Excel;
using ClosedXML.Excel;


namespace PnetDB_Maker_Jhlee
{
    public partial class PnetDB_Maker : Form
    {

        bool isFullListBox = false;// 한번 ListBox채우면 더 못채우도록
        string Path = null;
        string InputFile_Path = null;

        string OutputFile_Path1 = "\\PNet_Compoints.xlsx";
        string OutputFile_Path2 = "\\ComPoints";
        
        IXLRange ReadRange; // InputFile에서 복사한 셀 내용을 담는 컨테이너


        string[] AOAI_ROW1 = new string[] { "BasicSystemName", "PointName", "StationNum", "PointIndication", "CurrentValue", "Unit", "HistoryLibraryCollectionPeriod", "ProjectCalculationAttribute", "RangeUpperLimit", "RangeBottomLimit" };
        string[] DODI_ROW1 = new string[] {"BasicSystemName","PointName","StationNum","PointIndication","HistoryLibraryCollectionPeriod","ProjectCalculationAttribute"};
        string[] GWAI_ROW1 = new string[] {"BasicSystemName","PointName","StationNum","PointIndication" };
        string[] PC_ID = new string[] { "SC1P", "SC1S", "SD1P", "SD1S", "CB2P", "CB2S", "CD2P", "CD2S", "SB2P", "SB2S", "CA2P", "CA2S", "CC2P", "CC2S", "SA2P", "SA2S", "CB1P", "CB1S", "CD1P", "CD1S", "SB1P", "SB1S", "CA1P", "CA1S", "CC1P", "CC1S", "SA1P", "SA1S" };
        string[] GWAI_NAME = new string[] { "GWAI_BUF_A1",
                                            "GWAI_BUF_B1",
                                            "GWAI_BUF_A2",
                                            "GWAI_BUF_B2",
                                            "GWAI_BUF_A3",
                                            "GWAI_BUF_B3",
                                            "GWAI_BUF_A4",
                                            "GWAI_BUF_B4" };

        string[] GWAI_DESCRIPTION = new string[] { "Y1 Before 800EA",
                                                   "Y1 After 800EA",
                                                   "Y2 Before 800EA",
                                                   "Y2 After 800EA",
                                                   "Y3 Before 800EA",
                                                   "Y3 After 800EA",
                                                   "Y4 Before 800EA",
                                                   "Y4 After 800EA" };
        const int Compoints_ROW1_1 = 6;
        string[] Compoints_ROW2 = new string[] { "PIN", "SN", "RP", "SN", "PT", "DT", "DL", "FC", "AD", "PRE", "PRS" };
        string[] Compoints_ROW3 = new string[] { "CommunicationPointItemName", "StationNum", "RodPosition", "StepNumber", "PointType", "DataType", "DataLength", "FunctionCode", "Addr", "PysicalRangeEnd", "PysicalRangeStart" };

      

        public PnetDB_Maker()
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
            System.IO.FileInfo fi = new System.IO.FileInfo(Path + OutputFile_Path1);
            if (fi.Exists == false)
            {
                MessageBox.Show("PNet_Compoints.xlsx does not Exist!! ");
            }
            else
            {
                Write_Compoints_1to28();
            }
            Compoints_Saveas_CSV();

            MessageBox.Show("Completed");

        }
        private void Compoints_Saveas_CSV()
        {
            Excel.Application Excel_App = new Excel.Application();

            Excel_App.DisplayAlerts = false;

            for (int i=1; i<=28;++i)
            {
                var WorkBook = Excel_App.Workbooks.Open(@Path + OutputFile_Path2 + i.ToString() + ".xlsx");
                
                WorkBook.SaveAs(Path + OutputFile_Path2 + i.ToString() + ".csv", Excel.XlFileFormat.xlCSV);
                WorkBook.Close(false);
                Excel_App.Quit();
                System.IO.File.Delete(Path + OutputFile_Path2 + i.ToString() + ".xlsx");// Compoints 1~28 . xlsx Delete -> 휴지통에는 있음
            }
         

        }
        private void Write_Compoints_1to28()
        {
            var WorkBook = new XLWorkbook(@Path + OutputFile_Path1); //PNet_Compoints.xlsx  열기

            var AO_Sheet = WorkBook.Worksheet("AO");
            var AI_Sheet = WorkBook.Worksheet("AI");
            var DO_Sheet = WorkBook.Worksheet("DO");
            var DI_Sheet = WorkBook.Worksheet("DI");


            int PC_Num = 1;
            int Sour_Start_DI_Row = 2;
            int Sour_Start_DO_Row = 2;
            int Sour_Start_AI_Row = 2;
            int Sour_Start_AO_Row = 2;
            foreach (string pcid in PC_ID)
            {
                //DI-> DO -> AI-> AO 순으로 복사
                var Compoints_WB = new XLWorkbook();
                var Compoint_WS = Compoints_WB.Worksheets.Add("ComPoints" + PC_Num.ToString()); // Dest Work Sheet생성
                //행만들기 
                int Col = 1;
               
                Compoint_WS.Cell(1, 1).Value = Compoints_ROW1_1;
                foreach(string s in Compoints_ROW2)
                {
                    Compoint_WS.Cell(2, Col).Value = s;
                    ++Col;
                }
                Col = 1;
                foreach (string s in Compoints_ROW3)
                {

                    Compoint_WS.Cell(3, Col).Value = s;
                    ++Col;
                }

                //DI 복사
                //
                int Addr = 1;
                int Dest_Start_Row = 4;

                while (true)
                {
                    if (DI_Sheet.Cell(Sour_Start_DI_Row, 1).Value == null)
                    {
                        DI_Sheet.Cell(Sour_Start_DI_Row, 1).Value = "";
                        //MessageBox.Show(Sour_Start_Row.ToString() + "," + 1.ToString()+"Error");
                    }
                    if (DI_Sheet.Cell(Sour_Start_DI_Row, 1).Value.ToString() == pcid)
                    {
                        Compoint_WS.Cell(Dest_Start_Row, 1).Value = DI_Sheet.Cell(Sour_Start_DI_Row, 2).Value.ToString() + ".DV";
                        Compoint_WS.Cell(Dest_Start_Row, 2).Value = PC_Num;
                        Compoint_WS.Cell(Dest_Start_Row, 3).Value = 0;
                        Compoint_WS.Cell(Dest_Start_Row, 4).Value = 0;
                        Compoint_WS.Cell(Dest_Start_Row, 5).Value = 1;
                        Compoint_WS.Cell(Dest_Start_Row, 6).Value = 4;
                        Compoint_WS.Cell(Dest_Start_Row, 7).Value = 2;
                        Compoint_WS.Cell(Dest_Start_Row, 8).Value = 2;
                        Compoint_WS.Cell(Dest_Start_Row, 9).Value = Addr;
                        Compoint_WS.Cell(Dest_Start_Row, 10).Value = 1;
                        Compoint_WS.Cell(Dest_Start_Row, 11).Value = 0;


                        Dest_Start_Row++;
                        Sour_Start_DI_Row++;
                        Addr++;

                    }
                    else
                    {
                        break;
                    }
                }

                //DO 복사 : FunctionCode 이상
              
                Addr = 1;
                int[]DO_Fucntion_Code=  new int[] { 1, 5, 15 };


                while (true)
                {
                    if (DO_Sheet.Cell(Sour_Start_DO_Row, 1).Value == null)
                    {
                        DO_Sheet.Cell(Sour_Start_DO_Row, 1).Value = "";
                        //MessageBox.Show(Sour_Start_Row.ToString() + "," + 1.ToString()+"Error");
                    }
                    if (DO_Sheet.Cell(Sour_Start_DO_Row, 1).Value.ToString() == pcid)
                    {
                        for(int i= 0;i< 3;++i)
                        {
                            Compoint_WS.Cell(Dest_Start_Row, 1).Value = DO_Sheet.Cell(Sour_Start_DO_Row, 2).Value.ToString() + ".DV";
                            Compoint_WS.Cell(Dest_Start_Row, 2).Value = PC_Num;
                            Compoint_WS.Cell(Dest_Start_Row, 3).Value = 0;
                            Compoint_WS.Cell(Dest_Start_Row, 4).Value = 0;
                            Compoint_WS.Cell(Dest_Start_Row, 5).Value = 1;
                            Compoint_WS.Cell(Dest_Start_Row, 6).Value = 4;
                            Compoint_WS.Cell(Dest_Start_Row, 7).Value = 2;
                            Compoint_WS.Cell(Dest_Start_Row, 8).Value = DO_Fucntion_Code[i]; // Function Code 임시로 3으로해놓음
                            Compoint_WS.Cell(Dest_Start_Row, 9).Value = Addr;
                            Compoint_WS.Cell(Dest_Start_Row, 10).Value = 1;
                            Compoint_WS.Cell(Dest_Start_Row, 11).Value = 0;
                            Dest_Start_Row++;
                        }

                      
                        Sour_Start_DO_Row++;
                        Addr++;

                    }
                    else
                    {
                        break;
                    }
                }


                //AI 복사
                Addr = 1;

                while (true)
                {
                    if (AI_Sheet.Cell(Sour_Start_AI_Row, 1).Value == null)
                    {
                        AI_Sheet.Cell(Sour_Start_AI_Row, 1).Value = "";
                        //MessageBox.Show(Sour_Start_Row.ToString() + "," + 1.ToString()+"Error");
                    }
                    if (AI_Sheet.Cell(Sour_Start_AI_Row, 1).Value.ToString() == pcid)
                    {
                        Compoint_WS.Cell(Dest_Start_Row, 1).Value = AI_Sheet.Cell(Sour_Start_AI_Row, 2).Value.ToString() + ".AV";
                        Compoint_WS.Cell(Dest_Start_Row, 2).Value = PC_Num;
                        Compoint_WS.Cell(Dest_Start_Row, 3).Value = 0;
                        Compoint_WS.Cell(Dest_Start_Row, 4).Value = 0;
                        Compoint_WS.Cell(Dest_Start_Row, 5).Value = 1;
                        Compoint_WS.Cell(Dest_Start_Row, 6).Value = 6;
                        Compoint_WS.Cell(Dest_Start_Row, 7).Value = 2;
                        Compoint_WS.Cell(Dest_Start_Row, 8).Value = 4;
                        Compoint_WS.Cell(Dest_Start_Row, 9).Value = Addr;
                        Compoint_WS.Cell(Dest_Start_Row, 10).Value = AI_Sheet.Cell(Sour_Start_AI_Row, 9);
                        Compoint_WS.Cell(Dest_Start_Row, 11).Value = AI_Sheet.Cell(Sour_Start_AI_Row, 10);


                        Dest_Start_Row++;
                        Sour_Start_AI_Row++;
                        Addr++;

                    }
                    else
                    {
                        break;
                    }
                }

                //AO 복사
                Addr = 1;


                int[] AO_Fucntion_Code = new int[] { 3, 6, 16 };

                while (true)
                {
                    if (AO_Sheet.Cell(Sour_Start_AO_Row, 1).Value == null)
                    {
                        AO_Sheet.Cell(Sour_Start_AO_Row, 1).Value = "";
                        //MessageBox.Show(Sour_Start_Row.ToString() + "," + 1.ToString()+"Error");
                    }
                    if (AO_Sheet.Cell(Sour_Start_AO_Row, 1).Value.ToString() == pcid)
                    {
                        for(int i=0; i<3;++i)
                        {
                            Compoint_WS.Cell(Dest_Start_Row, 1).Value = AO_Sheet.Cell(Sour_Start_AO_Row, 2).Value.ToString() + ".AV";
                            Compoint_WS.Cell(Dest_Start_Row, 2).Value = PC_Num;
                            Compoint_WS.Cell(Dest_Start_Row, 3).Value = 0;
                            Compoint_WS.Cell(Dest_Start_Row, 4).Value = 0;
                            Compoint_WS.Cell(Dest_Start_Row, 5).Value = 1;
                            Compoint_WS.Cell(Dest_Start_Row, 6).Value = 6;
                            Compoint_WS.Cell(Dest_Start_Row, 7).Value = 2;
                            Compoint_WS.Cell(Dest_Start_Row, 8).Value = AO_Fucntion_Code[i];
                            Compoint_WS.Cell(Dest_Start_Row, 9).Value = Addr;
                            Compoint_WS.Cell(Dest_Start_Row, 10).Value = AO_Sheet.Cell(Sour_Start_AO_Row, 9);
                            Compoint_WS.Cell(Dest_Start_Row, 11).Value = AO_Sheet.Cell(Sour_Start_AO_Row, 10);
                            Dest_Start_Row++;
                        }

                       
                        Sour_Start_AO_Row++;
                        Addr++;

                    }
                    else
                    {
                        break;
                    }
                }

                Compoints_WB.SaveAs(@Path + OutputFile_Path2+PC_Num.ToString()+".xlsx");
                ++PC_Num;
            }

        }
        
        private void Write_PnetCompoints()
        {

           //
           var WorkBook = new XLWorkbook(); //PnetCompoint 엑셀파일 만들기
            
            //AO AI DO DI GWAI 5개의 Sheet 생성
            var AO_Sheet = WorkBook.Worksheets.Add("AO");
            var AI_Sheet = WorkBook.Worksheets.Add("AI");
            var DO_Sheet = WorkBook.Worksheets.Add("DO");
            var DI_Sheet = WorkBook.Worksheets.Add("DI");
            var GWAI_Sheet = WorkBook.Worksheets.Add("GWAI");
            int Col = 1;
        
            //1행 채우기
            foreach(string s in AOAI_ROW1)
            {
                AO_Sheet.Cell(1, Col).Value =s;
                AI_Sheet.Cell(1, Col).Value = s;
                ++Col;
            }
            Col = 1;
            foreach (string s in DODI_ROW1)
            {
                DO_Sheet.Cell(1, Col).Value = s;
                DI_Sheet.Cell(1, Col).Value = s;
                ++Col;
            }
            Col = 1;
            foreach (string s in GWAI_ROW1)
            {
                GWAI_Sheet.Cell(1, Col).Value = s;
                ++Col;
            }
            //AI 복사
            bool isFirstLoop = true;

            int Source_Row = 2;
            int Dest_Row = 2;
            const int AI_StartSourceRow = 2;

            int AO_StartSourceRow = 0;
            int DI_StartSourceRow = 0;
            int DO_StartSourceRow = 0;

            foreach (string pcid in PC_ID)
           {
                while(true)
                {
                    if (ReadRange.Cell(Source_Row, 3).Value == null)
                        ReadRange.Cell(Source_Row, 3).Value = "";
                    if(ReadRange.Cell(Source_Row, 3).Value.ToString().StartsWith("AI") !=true)
                    {
                       
                        break;
                    }
                    AI_Sheet.Cell(Dest_Row, 1).Value = pcid; 
                    AI_Sheet.Cell(Dest_Row, 2).Value = ReadRange.Cell(Source_Row,3).Value ;
                    if(AI_Sheet.Cell(Dest_Row, 2).Value!=null)
                    {
                        AI_Sheet.Cell(Dest_Row, 2).Value = AI_Sheet.Cell(Dest_Row, 2).Value.ToString().Replace("CA1P", pcid);
                    }
                    AI_Sheet.Cell(Dest_Row, 3).Value = 1; 
                    AI_Sheet.Cell(Dest_Row, 4).Value = ReadRange.Cell(Source_Row, 4).Value; 
                    AI_Sheet.Cell(Dest_Row, 5).Value = 0; 
                    AI_Sheet.Cell(Dest_Row, 6).Value = ReadRange.Cell(Source_Row, 11).Value;
                    AI_Sheet.Cell(Dest_Row, 7).Value = 1; 
                    AI_Sheet.Cell(Dest_Row, 8).Value = 0;


                   string[] Temp = ReadRange.Cell(Source_Row, 10).Value.ToString().Split('~');

                    AI_Sheet.Cell(Dest_Row, 9).Value = Temp[1]; 
                    AI_Sheet.Cell(Dest_Row, 10).Value = Temp[0];
                    
                    Source_Row++;
                    Dest_Row++;
                }
                if (isFirstLoop == true)
                {
                    AO_StartSourceRow = ++Source_Row; // 첫루프에 한번만 AO Start 지점 저장
                    isFirstLoop = false;
                }
           
                Source_Row = AI_StartSourceRow;

          }
            //AO 복사
            isFirstLoop = true;
            Dest_Row = 2;
            Source_Row = AO_StartSourceRow;
            foreach (string pcid in PC_ID)
            {
                while (true)
                {
                    if (ReadRange.Cell(Source_Row, 3).Value == null)
                        ReadRange.Cell(Source_Row, 3).Value = "";
                    if (ReadRange.Cell(Source_Row, 3).Value.ToString().StartsWith("AO") != true)
                    {
                        break;
                    }
                    AO_Sheet.Cell(Dest_Row, 1).Value = pcid;
                    AO_Sheet.Cell(Dest_Row, 2).Value = ReadRange.Cell(Source_Row, 3).Value;
                    if (AO_Sheet.Cell(Dest_Row, 2).Value != null)
                    {
                        AO_Sheet.Cell(Dest_Row, 2).Value = AO_Sheet.Cell(Dest_Row, 2).Value.ToString().Replace("CA1P", pcid);
                    }
                    AO_Sheet.Cell(Dest_Row, 3).Value = 1;
                    AO_Sheet.Cell(Dest_Row, 4).Value = ReadRange.Cell(Source_Row, 4).Value;
                    AO_Sheet.Cell(Dest_Row, 5).Value = 0;
                    AO_Sheet.Cell(Dest_Row, 6).Value = ReadRange.Cell(Source_Row, 11).Value;
                    AO_Sheet.Cell(Dest_Row, 7).Value = 1;
                    AO_Sheet.Cell(Dest_Row, 8).Value = 0;


                    string[] Temp = ReadRange.Cell(Source_Row, 10).Value.ToString().Split('~');

                    AO_Sheet.Cell(Dest_Row, 9).Value = Temp[1];
                    AO_Sheet.Cell(Dest_Row, 10).Value = Temp[0];

                    Source_Row++;
                    Dest_Row++;
                }
                if (isFirstLoop == true)
                {
                    DI_StartSourceRow = ++Source_Row; // 첫루프에 한번만 AO Start 지점 저장
                    isFirstLoop = false;
                }
                Source_Row = AO_StartSourceRow;

            }

            //DI 복사
            isFirstLoop = true;
            Dest_Row = 2;
            Source_Row = DI_StartSourceRow;
            foreach (string pcid in PC_ID)
            {
                while (true)
                {
                    if (ReadRange.Cell(Source_Row, 3).Value == null)
                        ReadRange.Cell(Source_Row, 3).Value = "";
                    if (ReadRange.Cell(Source_Row, 3).Value.ToString().StartsWith("DI") != true)
                    {
                        break;
                    }
                    DI_Sheet.Cell(Dest_Row, 1).Value = pcid;
                    DI_Sheet.Cell(Dest_Row, 2).Value = ReadRange.Cell(Source_Row, 3).Value;
                    if (DI_Sheet.Cell(Dest_Row, 2).Value != null)
                    {
                        DI_Sheet.Cell(Dest_Row, 2).Value = DI_Sheet.Cell(Dest_Row, 2).Value.ToString().Replace("CA1P", pcid);
                    }
                    DI_Sheet.Cell(Dest_Row, 3).Value = 1;
                    DI_Sheet.Cell(Dest_Row, 4).Value = ReadRange.Cell(Source_Row, 4).Value;
                    DI_Sheet.Cell(Dest_Row, 5).Value = 1;
                    DI_Sheet.Cell(Dest_Row, 6).Value = 0;
                    DI_Sheet.Cell(Dest_Row, 7).Value = 1;
                    DI_Sheet.Cell(Dest_Row, 8).Value = 0;

                    Source_Row++;
                    Dest_Row++;
                }
                if (isFirstLoop == true)
                {
                    DO_StartSourceRow = ++Source_Row; // 첫루프에 한번만 DO Start 지점 저장
                    isFirstLoop = false;
                }
                Source_Row = DI_StartSourceRow;

            }


            //DO 복사
            Dest_Row = 2;
            Source_Row = DO_StartSourceRow;
            foreach (string pcid in PC_ID)
            {
                while (true)
                {
                    if (ReadRange.Cell(Source_Row, 3).Value == null)
                        ReadRange.Cell(Source_Row, 3).Value = "";
                    if (ReadRange.Cell(Source_Row, 3).Value.ToString().StartsWith("DO") != true)
                    {
                        break;
                    }
                    DO_Sheet.Cell(Dest_Row, 1).Value = pcid;
                    DO_Sheet.Cell(Dest_Row, 2).Value = ReadRange.Cell(Source_Row, 3).Value;
                    if (DO_Sheet.Cell(Dest_Row, 2).Value != null)
                    {
                        DO_Sheet.Cell(Dest_Row, 2).Value = DO_Sheet.Cell(Dest_Row, 2).Value.ToString().Replace("CA1P", pcid);
                    }
                    DO_Sheet.Cell(Dest_Row, 3).Value = 1;
                    DO_Sheet.Cell(Dest_Row, 4).Value = ReadRange.Cell(Source_Row, 4).Value;
                    DO_Sheet.Cell(Dest_Row, 5).Value = 1;
                    DO_Sheet.Cell(Dest_Row, 6).Value = 0;
                    DO_Sheet.Cell(Dest_Row, 7).Value = 1;
                    DO_Sheet.Cell(Dest_Row, 8).Value = 0;

                    Source_Row++;
                    Dest_Row++;
                }
             
                Source_Row = DO_StartSourceRow;

            }


            //GWAI 복사
            Dest_Row = 2;
            int i = 0;
            foreach (string gwainame in GWAI_NAME)
            {
                GWAI_Sheet.Cell(Dest_Row, 2).Value = gwainame;
                GWAI_Sheet.Cell(Dest_Row, 3).Value = 1;
                GWAI_Sheet.Cell(Dest_Row, 4).Value = GWAI_DESCRIPTION[i];
                ++i;
                ++Dest_Row;
            }

            WorkBook.SaveAs(@Path + OutputFile_Path1);//PnetCompoints.xlsx 로저장




        }

        private void Read_InputFile()
        {
            
                var WorkBook = new XLWorkbook(InputFile_Path); // 입력받은 파일 엑셀 Open

                var WorkSheet = WorkBook.Worksheet("Sheet1"); // 워크시트가져오기

                var First_Cell = WorkSheet.FirstCellUsed();
                var Last_Cell = WorkSheet.LastCellUsed();
                ReadRange = WorkSheet.Range(First_Cell.Address, Last_Cell.Address);

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
            label1.Text = "1.Drag and Drop the MTP I/O List\r\n" + "2.Run Button Click";

        }

      

      
    }
}
