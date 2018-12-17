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
        string Path=null;
        string InputFile_Path=null;

        object[,] ReadData;

        Excel.Application Excel_App = null;
        Excel.Workbook Rd_Wb = null;
        Excel.Worksheet Rd_Ws = null;

        Excel.Workbook Wr_Wb = null;
        Excel.Worksheet Wr_Ws = null;

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
                return;

            Read_InputFile();

            //버튼1을 누르면 Excel에 데이터를 읽어들여서(Read_InputFile()-> PnetCompoint.xlxs 형태로 만든다


        }

        private void Read_InputFile()
        {

            try
            {
                Excel_App = new Excel.Application();
                Rd_Wb = Excel_App.Workbooks.Open("@" + InputFile_Path);// 해당 경로에 워크북 열기
                Rd_Ws = Rd_Wb.Worksheets.Add();//Work Sheet추가 

                Excel.Range rng = Rd_Ws.UsedRange;
                ReadData = rng.Value;


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
