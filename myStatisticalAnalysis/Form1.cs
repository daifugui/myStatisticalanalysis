using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;
namespace myStatisticalAnalysis
{
    
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public DataSet RawDataSet=new DataSet();
        public DataTable RawDataTable = new DataTable();
        public DataSet readExcel(string excelpath)
        {
            string strCon = " Provider = Microsoft.Jet.OLEDB.4.0 ; Data Source = "+excelpath +";Extended Properties=Excel 8.0";
            OleDbConnection myConn = new OleDbConnection(strCon);
            myConn.Open();

            DataTable dtSheetName = myConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" });

            string[] strTableNames = new string[dtSheetName.Rows.Count];
            for (int k = 0; k < dtSheetName.Rows.Count; k++)
            {
                strTableNames[k] = dtSheetName.Rows[dtSheetName.Rows.Count - k - 1]["TABLE_NAME"].ToString();
            }

           // string strCom = " SELECT * FROM [Sheet1$] ";
              RawDataSet.Clear();
              for (int k = 0; k < strTableNames.Length; k++)
              {
                  string strCom = " SELECT * FROM [" + strTableNames[k] + "] ";

                  OleDbDataAdapter myCommand = new OleDbDataAdapter(strCom, myConn);
                  myCommand.Fill(RawDataSet, "[" + strTableNames[k] + "]");
              }
  
          //  myDataSet = new DataSet();
           
           // myCommand.Fill(RawDataSet, "[Sheet1$]");      
            
            //  myCommand.Fill(
           // RawDataSet.t
            myConn.Close();
            return RawDataSet;
        }
        private void openpToolStripMenuItem_Click(object sender, EventArgs e)
        {
          //  this.openFileDialog1.Filter = "*.xls";
            this.openFileDialog1.Filter = "(*.xls)|*.xls";
            if(this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string FileName = this.openFileDialog1.FileName;
                readExcel(FileName);
               // RawDataTable = RawDataSet.Tables["[demo1]"];
                RawDataTable = RawDataSet.Tables["[Sheet1$]"];
               // dataGridView1.DataMember = "[Sheet1$]";
                this.dataGridView1.DataSource = RawDataTable;
                int rowNumber = 1;
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    row.HeaderCell.Value = rowNumber.ToString();
                    rowNumber++;
                }
                for (int i = 0; i < this.dataGridView1.Columns.Count; i++)
                {
                    this.dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                }
               //this.dataGridView1.Refresh();
                // 你的 处理文件路径代码 
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        int i = 10;
        public Bitmap createChartImage(int Width, int Height)
        {
            System.Drawing.Bitmap image1 = new Bitmap(Width, Height);
            Graphics g = Graphics.FromImage(image1);
            g.Clear(Color.Black);
            g.DrawLine(Pens.Gold, 10 + i, 10, 110, 10 + i);
            Point[] pp ={ new Point(1, 2), new Point(2, 3), new Point(5, 7), new Point(9, 7), new Point(9, 1) };

            g.DrawLines(Pens.Gold, pp);
            i += 10;
            Font font = new Font("宋体", 30f); //字是什么样子的？

            Brush brush = Brushes.Red; //用红色涂上我的字吧；
            string str = "Baidu"; //写什么字？

//Font font =new  Font("宋体",30f); //字是什么样子的？

//Brush brush = Brushes.Red; //用红色涂上我的字吧；

PointF point = new PointF(10f,10f); //从什么地方开始写字捏？

 

//横着写还是竖着写呢？

System.Drawing.StringFormat sf = new System.Drawing.StringFormat();

//还是竖着写吧

sf.FormatFlags = StringFormatFlags.DirectionVertical;

 

//开始写咯

g.DrawString(str,font,brush,point,sf);

 
g.DrawString("hello", font, brush,20f, 20f);
            g.Dispose();
            return image1;
        }
        private void sPCToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("hello");
            this.pictureBox1.Image = createChartImage(pictureBox1.Width, pictureBox1.Height);
        }
    }
}