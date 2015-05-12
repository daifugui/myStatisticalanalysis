using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
//using Microsoft.Reporting.WinForms;
using ZedGraph;
using System.Collections;
using System.IO.Ports;
using System.Threading;
using System.IO;
using System.Data.SqlClient;
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
        public DataSet readExcel(string fileName)
        {
             string connStr ;
             if (fileName.EndsWith(".xls"))
               connStr = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + fileName + ";" + ";Extended Properties=\"Excel 8.0;HDR=NO;IMEX=1\"";
           else
                 connStr = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + fileName + ";" + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";



        //    if (fileType == ".xls")
        //        connStr = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + fileName + ";" + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\"";
       //     else
          //      connStr = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + fileName + ";" + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";

      //      string strCon = " Provider = Microsoft.Jet.OLEDB.4.0 ; Data Source = " + excelpath + ";Extended Properties=Excel 8.0;HDR=NO";
           OleDbConnection myConn = new OleDbConnection(connStr);
            myConn.Open();

            DataTable dtSheetName = myConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" });

            string[] strTableNames = new string[dtSheetName.Rows.Count];
            for (int k = 0; k < dtSheetName.Rows.Count; k++)
            {
                strTableNames[k] = dtSheetName.Rows[k]["TABLE_NAME"].ToString();
            }

           // string strCom = " SELECT * FROM [Sheet1$] ";
              RawDataSet.Tables.Clear();

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
            this.openFileDialog1.Filter = "(*.xls)|*.xls|(*.xlsx)|*.xlsx";
            if(this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string FileName = this.openFileDialog1.FileName;
                readExcel(FileName);
               // RawDataTable = RawDataSet.Tables["[demo1]"];
                RawDataTable = RawDataSet.Tables[0];
               // dataGridView1.DataMember = "[Sheet1$]";
               // this.dataGridView1.Rows.Clear();
               // DataTable dt = (DataTable)this.dataGridView1.DataSource;
              //  dt.Rows.Clear();
                //foreach (DataRow Dr in RawDataTable.Rows)
               // {
                    //MessageBox.Show(Dr[0].ToString());
                    //string ss = Dr[0].ToString();
              // }
                this.dataGridView1.Columns.Clear();
                this.dataGridView1.DataSource = RawDataTable;
              
               // dataGridView1.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
                dataGridView1.RowHeadersWidth = 80;
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
        public Bitmap createChartImage1(int Width, int Height, double[] showDataY,double minY,double maxY)
        {
            System.Drawing.Bitmap image1 = new Bitmap(Width, Height);
            Graphics g = Graphics.FromImage(image1);
            g.Clear(Color.Black);
            double xd = (Width - 20) / (showDataY.Length - 1);
            for (int i = 0; i < showDataY.Length-1;i++)
            {
                double  py0= (double)Height * 0.9 - (double)(showDataY[i] - minY) * (Height * 0.9 - Height * 0.1) / (maxY - minY);
                double  py1= (double)Height * 0.9 - (double)(showDataY[i+1] - minY) * (Height * 0.9 - Height * 0.1) / (maxY - minY);
                g.DrawLine(Pens.Gold, 10 + (int)xd * i, (int)py0, 10 + (int)xd * (i + 1), (int)py1);
            }
        //    Point[] pp = { new Point(1, 2), new Point(2, 3), new Point(5, 7), new Point(9, 7), new Point(9, 1) };

        //    g.DrawLines(Pens.Gold, pp);
 
            Font font = new Font("宋体", 30f); //字是什么样子的？

            Brush brush = Brushes.Red; //用红色涂上我的字吧；
            string str = "Baidu"; //写什么字？


            PointF point = new PointF(10f, 10f); //从什么地方开始写字捏？



            //横着写还是竖着写呢？

            System.Drawing.StringFormat sf = new System.Drawing.StringFormat();

            //还是竖着写吧

            sf.FormatFlags = StringFormatFlags.DirectionVertical;



            //开始写咯

            g.DrawString(str, font, brush, point, sf);


            g.DrawString("hello", font, brush, 20f, 20f);
            g.Dispose();
            return image1;
        }
        public void computeRangeMean(double[] testdata, int subgroupsize, ArrayList R, ArrayList mean)
        {
            int kk=0;
            double maxX=0;
            double minX=0;
           // int m=0;
            R.Clear();
            double sum = 0;
            foreach (double x in testdata)
            {
                if (kk == 0)
                {
                    maxX = x;
                    minX = x;
                    sum = 0;
                }
                if (x > maxX) maxX = x;
                else if (x < minX) minX = x;
                sum += x;
                kk++;
                if (kk == subgroupsize)
                {
                   // R[m++] = maxX - minX;
                    R.Add(maxX - minX);
                    mean.Add(sum / subgroupsize);
                    sum = 0;
                    kk = 0;
                }
            }
       //     int nn=testdata.Length;
         //   for(int i=0;i<nn;i++)
           // {

         // }
 
        }
        /*
         25 0.600 0.153 0.606 0.9896 1.0105 0.565 1.435 0.559 1.420 3.931 0.2544 0.708 1.806 6.056 0.459 1.541
         24 0.612 0.157 0.619 0.9892 1.0109 0.555 1.445 0.549 1.429 3.895 0.2567 0.712 1.759 6.031 0.451 1.548
         23 0.626 0.162 0.633 0.9887 1.0114 0.545 1.455 0.539 1.438 3.858 0.2592 0.716 1.710 6.006 0.443 1.557
         22 0.640 0.167 0.647 0.9882 1.0119 0.534 1.466 0.528 1.448 3.819 0.2618 0.720 1.659 5.979 0.434 1.566
         21 0.655 0.173 0.663 0.9876 1.0126 0.523 1.477 0.516 1.459 3.778 0.2647 0.724 1.605 5.951 0.425 1.575
         20 0.671 0.180 0.680 0.9869 1.0133 0.510 1.490 0.504 1.470 3.735 0.2677 0.729 1.549 5.921 0.415 1.585
        19 0.688 0.187 0.698 0.9862 1.0140 0.497 1.503 0.490 1.483 3.689 0.2711 0.734 1.487 5.891 0.403 1.597
        18 0.707 0.194 0.718 0.9854 1.0148 0.482 1.518 0.475 1.496 3.640 0.2747 0.739 1.424 5.856 0.391 1.608
        17 0.728 0.203 0.739 0.9845 1.0157 0.466 1.534 0.458 1.511 3.588 0.2787 0.744 1.356 5.820 0.378 1.622
        16 0.750 0.212 0.763 0.9835 1.0168 0.448 1.552 0.440 1.526 3.532 0.2831 0.750 1.282 5.782 0.363 1.637
        15 0.775 0.223 0.789 0.9823 1.0180 0.428 1.572 0.421 1.544 3.472 0.2880 0.756 1.203 5.741 0.347 1.653
        14 0.802 0.235 0.817 0.9810 1.0194 0.406 1.594 0.399 1.563 3.407 0.2935 0.763 1.118 5.696 0.328 1.672
        13 0.832 0.249 0.850 0.9794 1.0210 0.382 1.618 0.374 1.585 3.336 0.2998 0.770 1.025 5.647 0.307 1.693
        12 0.866 0.266 0.886 0.9776 1.0229 0.354 1.646 0.346 1.610 3.258 0.3069 0.778 0.922 5.594 0.283 1.717
        11 0.905 0.285 0.927 0.9754 1.0252 0.321 1.679 0.313 1.637 3.173 0.3152 0.787 0.811 5.535 0.256 1.744
        10 0.949 0.308 0.975 0.9727 1.0281 0.284 1.716 0.276 1.669 3.078 0.3249 0.797 0.687 5.469 0.223 1.777
        9 1.000 0.337 1.032 0.9693 1.0317 0.239 1.761 0.232 1.707 2.970 0.3367 0.808 0.547 5.393 0.184 1.816
        8 1.061 0.373 1.099 0.9650 1.0363 0.185 1.815 0.179 1.751 2.847 0.3512 0.820 0.388 5.306 0.136 1.864
        7 1.134 0.419 1.182 0.9594 1.0423 0.118 1.882 0.113 1.806 2.704 0.3698 0.833 0.204 5.204 0.076 1.924
        6 1.225 0.483 1.287 0.9515 1.0510 0.030 1.970 0.029 1.874 2.534 0.3946 0.848 0 5.078 0 2.004
        5 1.342 0.577 1.427 0.9400 1.0638 0 2.089 0 1.964 2.326 0.4299 0.864 0 4.918 0 2.114
        4 1.500 0.729 1.628 0.9213 1.0854 0 2.266 0 2.088 2.059 0.4857 0.880 0 4.698 0 2.282
        3 1.732 1.023 1.954 0.8862 1.1284 0 2.568 0 2.276 1.693 0.5907 0.888 0 4.358 0 2.574
        2 2.121 1.880 2.659 0.7979 1.2533 0 3.267 0 2.606 1.128 0.8865 0.853 0 3.686 0 3.267
         
         * */

        double[] SPC_A = { 2.12100000000000, 1.73200000000000, 1.50000000000000, 1.34200000000000, 1.22500000000000, 1.13400000000000, 1.06100000000000, 1, 0.949000000000000, 0.905000000000000, 0.866000000000000, 0.832000000000000, 0.802000000000000, 0.775000000000000, 0.750000000000000, 0.728000000000000, 0.707000000000000, 0.688000000000000, 0.671000000000000, 0.655000000000000, 0.640000000000000, 0.626000000000000, 0.612000000000000, 0.600000000000000 };
        double[] SPC_A2 = { 1.8800, 1.0230, 0.7290, 0.5770, 0.4830, 0.4190, 0.3730, 0.3370, 0.3080, 0.2850, 0.2660, 0.2490, 0.2350, 0.2230, 0.2120, 0.2030, 0.1940, 0.1870, 0.1800, 0.1730, 0.1670, 0.1620, 0.1570, 0.1530 };
        double[] SPC_D3={ 0 , 0 ,0, 0 ,   0 , 0.0760, 0.1360,0.1840 , 0.2230,0.2560 ,0.2830,0.3070,0.3280,0.3470, 0.3630,0.3780,0.3910,0.4030 ,0.4150 ,0.4250 ,0.4340, 0.4430,0.4510,0.4590};
        double[] SPC_D4 = { 3.2670,2.5740,2.2820,2.1140 ,2.0040,1.9240 ,1.8640,1.8160,1.7770 ,1.7440,1.7170,1.6930,1.6720,1.6530,1.6370 ,1.6220,1.6080 ,1.5970,1.5850,1.5750,1.5660,1.5570,1.5480,1.5410};
        public bool computeLimit(double[] R, double[] mean, int subgroupsize,  ref double  Ave_R, ref double  Ave_mean, ref double UCL, ref double LCL, ref double UCLR, ref double LCLR)
         {
             if (subgroupsize <= 1) return false;
             int m = R.Length;
            Ave_R=0;
            Ave_mean=0;
            for (int i = 0; i < m; i++)
            {
                Ave_R += R[i];
                Ave_mean += mean[i]; 
            }
            Ave_R /= m;
           // Ave_R = Math.Sqrt(Ave_R);
            Ave_mean /= m;
            UCL = Ave_mean + SPC_A2[subgroupsize - 2] * Ave_R;
            LCL = Ave_mean - SPC_A2[subgroupsize - 2] * Ave_R;
           
            UCLR = SPC_D4[subgroupsize - 2] * Ave_R;
            LCLR = SPC_D3[subgroupsize - 2] * Ave_R;
            return true; 
        }
        public void ShowXbarRChart(DataTable RawDataTable, int subgroupsize, ZedGraph.ZedGraphControl ZedxbarChart, ZedGraph.ZedGraphControl ZedRchart)
        {
            ZedRchart.GraphPane.CurveList.Clear();
            ZedRchart.GraphPane.GraphItemList.Clear();
            ZedxbarChart.GraphPane.CurveList.Clear();
            ZedxbarChart.GraphPane.GraphItemList.Clear();      
            int nn = RawDataTable.Rows.Count;
            double[] dataY = new double[nn];
            //double[] dataX=new double[nn];
            for (int i = 0; i < nn; i++)
            {
                dataY[i] = (double)System.Convert.ToDouble(RawDataTable.Rows[i][0]);
            }

            //   double[] showdataY = {1,2,4,3,2,3,4,4,1,23,12,12,12,12,1,56,323};
            //     double[] showdataX = { 0,1, 2, 3, 4, 5, 6, 7, 8 };
            //    this.pictureBox1.Image = createChartImage1(pictureBox1.Width, pictureBox1.Height,showdataY,0,4000);
            // this.reportViewer1.LocalReport.DataSources.Add(new ReportDataSource(RawDataTable.TableName, RawDataTable));
            //  this.reportViewer1.RefreshReport();
            // subgroupsize = 5;
            ArrayList R = new ArrayList();
            ArrayList mean = new ArrayList();
            computeRangeMean(dataY, subgroupsize, R, mean);
            double[] RR = (double[])R.ToArray(Type.GetType("System.Double"));
            double[] mm = (double[])mean.ToArray(Type.GetType("System.Double"));
            double Ave_R = 0;
            double Ave_mean = 0;
            double UCL = 0;
            double LCL = 0;
            double UCLR = 0;
            double LCLR = 0;

            //defAxisBorderPenWidth=0.1;
            computeLimit(RR, mm, subgroupsize, ref Ave_R, ref Ave_mean, ref UCL, ref LCL, ref  UCLR, ref LCLR);
            string printparameter_ss = string.Format("subgroupsize:{0}\nAve_R:{1}\nAve_mean:{2}\nUCL:{3}\nLCL:{4}\nUCLR:{5}\nLCLR:{6}", subgroupsize, Ave_R, Ave_mean,UCL, LCL, UCLR, LCLR);
         //   printparameter_ss
            this.label1.Text=printparameter_ss;
            //     zedGraphControl1.=0.1;
            //      zedGraphControl1.GraphPane.AddCurve("line", showdataX, showdataY, Color.Red, SymbolType.None);
            //   zedGraphControl1.GraphPane.AddCurve();
            nn = R.Count;
            double[] showdataX = new double[nn];
            for (int i = 0; i < nn; i++)
                showdataX[i] = i;
            ZedRchart.GraphPane.AddCurve("Data", null, (double[])R.ToArray(Type.GetType("System.Double")), Color.Blue, SymbolType.Circle);

            PointPairList lineAveR = new PointPairList();
            lineAveR.Add(0, Ave_R);
            lineAveR.Add(nn, Ave_R);
            ZedRchart.GraphPane.AddCurve("Ave_R", lineAveR, Color.Green, SymbolType.None);

            PointPairList lineUCLR = new PointPairList();
            lineUCLR.Add(0, UCLR);
            lineUCLR.Add(nn, UCLR);
            //   zedGraphControl1.GraphPane.ScaledPenWidth(3, 1);

            ZedRchart.GraphPane.AddCurve("UCLR", lineUCLR, Color.Red, SymbolType.None);

            PointPairList lineLCLR = new PointPairList();
            lineLCLR.Add(0, LCLR);
            lineLCLR.Add(nn, LCLR);
            ZedRchart.GraphPane.AddCurve("LCLR", lineLCLR, Color.Red, SymbolType.None);
            ZedRchart.GraphPane.Title = "R chart";
            //  zedGraphControl1.GraphPane.
            ZedRchart.GraphPane.YAxis.IsZeroLine = false;

            //    zedGraphControl1.GraphPane.AxisBorder.IsVisible = false;
            // zedGraphControl1.GraphPane.PaneBorder.IsVisible = false;
            double minR = (double)R[0];
            double maxR = (double)R[0];
            foreach(double Ri in R)
            {
                if (Ri < minR) minR = Ri;
                else if (Ri > maxR) maxR = Ri;
            }
            double Ymin = minR < LCLR ? minR : LCLR - (maxR - minR) * 0.2;
            if (minR >= 0 && LCLR >= 0 && Ymin < 0)
                ZedRchart.GraphPane.YAxis.Min = Ymin;
            //ZedRchart.GraphPane.YAxis.Min = minR< LCLR?minR:LCLR - (maxR - minR) * 0.2;
            ZedRchart.AxisChange();
         //   double Ymin = ZedRchart.GraphPane.YAxis.Min;
          //  double Ymax = ZedRchart.GraphPane.YAxis.Max;
           
            //    zedGraphControl1.GraphPane.AddCurve("LCLR", lineLCLR, Color.Red, SymbolType.None);
            //   zedGraphControl1.Refresh();
            ZedRchart.Invalidate();
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            nn = mean.Count;
         //   double[] showdataX = new double[nn];
            for (int i = 0; i < nn; i++)
                showdataX[i] = i;
            ZedxbarChart.GraphPane.AddCurve("Data", null, (double[])mean.ToArray(Type.GetType("System.Double")), Color.Blue, SymbolType.Circle);

            PointPairList linemean = new PointPairList();
            linemean.Add(0, Ave_mean);
            linemean.Add(nn, Ave_mean);
            ZedxbarChart.GraphPane.AddCurve("X_bar", linemean, Color.Green, SymbolType.None);

            PointPairList lineUCL = new PointPairList();
            lineUCL.Add(0, UCL);
            lineUCL.Add(nn, UCL);
            //   zedGraphControl1.GraphPane.ScaledPenWidth(3, 1);

            ZedxbarChart.GraphPane.AddCurve("UCL", lineUCL, Color.Red, SymbolType.None);

            PointPairList lineLCL = new PointPairList();
            lineLCL.Add(0, LCL);
            lineLCL.Add(nn, LCL);
            ZedxbarChart.GraphPane.AddCurve("LCL", lineLCL, Color.Red, SymbolType.None);
            ZedxbarChart.GraphPane.Title = "X_Bar chart";
            //  zedGraphControl1.GraphPane.
            ZedxbarChart.GraphPane.YAxis.IsZeroLine = false;

            //    zedGraphControl1.GraphPane.AxisBorder.IsVisible = false;
            // zedGraphControl1.GraphPane.PaneBorder.IsVisible = false;
           // double minM= (double)mean[0];
          //  double maxM = (double)mean[0];
           // foreach (double meani in mean)
           // {
              //  if (meani < minM) minM = meani;
              //  else if (meani > maxM) maxM = meani;
           // }
            //  Ymin=
            ZedxbarChart.GraphPane.YAxis.IsZeroLine = false;
            ZedxbarChart.AxisChange();
          //  ZedRchart.GraphPane.YAxis.MinAuto = false;
        //    ZedxbarChart.GraphPane.YAxis.Min = minM < LCL ? minM : LCL - (maxM - minM) * 0.2 ;
           // ZedxbarChart.AxisChange();
            //
            
            //   double Ymin = ZedRchart.GraphPane.YAxis.Min;
            //  double Ymax = ZedRchart.GraphPane.YAxis.Max;

            //    zedGraphControl1.GraphPane.AddCurve("LCLR", lineLCLR, Color.Red, SymbolType.None);
            //   zedGraphControl1.Refresh();
            ZedxbarChart.Invalidate();
             
        }
        public DataTable GetDgvToTable(DataGridView dgv)
        {
            DataTable dt = new DataTable();

            // 列强制转换
            for (int count = 0; count < dgv.Columns.Count; count++)
            {
                DataColumn dc = new DataColumn(dgv.Columns[count].Name.ToString());
                dt.Columns.Add(dc);
            }

            // 循环行
            for (int count = 0; count < dgv.Rows.Count; count++)
            {
                DataRow dr = dt.NewRow();
                for (int countsub = 0; countsub < dgv.Columns.Count; countsub++)
                {
                    dr[countsub] = Convert.ToString(dgv.Rows[count].Cells[countsub].Value);
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }
        private void sPCToolStripMenuItem_Click(object sender, EventArgs e)
        {
          //  ThreadRun = false;
        //    this.sPort.Close();
            DataTable dtSource = GetDgvToTable(this.dataGridView1);
            ShowXbarRChart(dtSource, 5, zedGraphControl1, zedGraphControl2);
            ShowDistributionChart(dtSource, zedGraphControl3);
        }

        public void ShowDistributionChart(DataTable RawDataTable, ZedGraph.ZedGraphControl ZedChart0)
        {
            ZedChart0.GraphPane.CurveList.Clear();
            ZedChart0.GraphPane.GraphItemList.Clear();
            int nn = RawDataTable.Rows.Count;
            double[] dataY = new double[nn];
          //  double[] dataX=new double[nn];
            double maxY = System.Convert.ToDouble(RawDataTable.Rows[0][0]);
            double minY = maxY;
            double sumY = 0;
            double sumY2 = 0;
            for (int i = 0; i < nn; i++)
            {
                dataY[i] = System.Convert.ToDouble(RawDataTable.Rows[i][0]);
                if (dataY[i] > maxY) 
                    maxY = dataY[i];
                else if (dataY[i] < minY) 
                    minY = dataY[i];

                sumY += dataY[i];
                sumY2 += (dataY[i] * dataY[i]);
            }
            double meanY = sumY / nn;
            double VarY = (sumY2 - nn * meanY * meanY) / (nn - 1);
             double StdY=Math.Sqrt(VarY);
            double [] dataYn =new double[7]{0,0,0,0,0,0,0};
            double[] dataX0 =new double[7]{30,30.5,31,31.5,32,32.2,32.4};
            for (int i = 0; i < 7; i++)
            {
                dataX0[i] =minY+ i * (maxY - minY) / 7 + (maxY - minY)/14;
            }
                for (int i = 0; i < nn; i++)
                {
                    if (dataY[i] == maxY)
                        dataYn[6]++;
                    else
                        dataYn[(int)(7 * (dataY[i] - minY) / (maxY - minY))]++;

                }



               ZedChart0.GraphPane.AddBar("Data Distribution", dataX0, dataYn, Color.Blue);
               // Axis aa=ZedChart0.GraphPane.BarBaseAxis();
              //  ZedChart0.GraphPane.MinBarGap = (float)0.02;
                ZedChart0.GraphPane.ClusterScaleWidth = (float)(maxY-minY)/(7);

            PointPairList datanorm=new PointPairList();
            //    double[] datanorm = new double[100];
                for (int i = 0; i < 100; i++)
                {
                    double xi=0;
                    double yi=0;
                    xi=meanY - 4* StdY+ (8 * StdY) *i/ 100;
                    yi = Math.Exp(-(xi - meanY) * (xi - meanY) / VarY / 2) / Math.Sqrt(2 * Math.PI) / StdY ;
                    datanorm.Add(xi, yi * nn);
 
                }
          //      ZedChart0.GraphPane.XAxis.Min = meanY - 6 * StdY;
              //  ZedChart0.GraphPane.XAxis.Max=meanY + 6 * StdY;

                ZedChart0.GraphPane.AddCurve("norm", datanorm, Color.Green,SymbolType.None);
                  //  for (double x = meanY - 6 * StdY; x <= meanY + 6 * StdY; x += (12 * StdY) / 100)
                //    {
                  //  }

                    //   ZedChart0.GraphPane.MinClusterGap = (float)0.002;
                    //  ZedChart0.GraphPane.AddBar
                    ZedChart0.AxisChange();
                      ZedChart0.Invalidate();
       
 
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            sPort = new SerialPort();
            if (sPort.IsOpen)
            {
                sPort.Close();
            }
            Control.CheckForIllegalCrossThreadCalls = false;
          
         //   sPort.DataReceived += new SerialDataReceivedEventHandler(port_DataReceived); 
             
            
        }
        SerialPort sPort;
        string Valuess=null;
        private void port_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            if (sPort.IsOpen)
            {
               // byte[] data = Convert.FromBase64String(sPort.ReadLine());
              //  this.label1.Text = Encoding.Unicode.GetString(data);
                
           //     int count = sPort.BytesToRead;
           //     byte[] data = new byte[count];
                // sPort.Read(data,0,count);
            // string ss=sPort.ReadTo("\n");
           // this.label1.Text = Encoding.Unicode.GetString(a);
            //this.label1.Text = ss;
           //   this.label1.Text = data.ToString();
               // this.label1.Text = Encoding.Unicode.GetString(data);
           //   this.label1.Text = Encoding.ASCII.GetString(data);

                try
                {
                    sPort.NewLine = "\r";

                    string ss = sPort.ReadLine();
                    if (ss == "S")
                    {
                        DataGridViewRow aa = new DataGridViewRow();
                        // this.dataGridView1.Rows.Add();
                        int index = this.dataGridView1.Rows.Add();
                        this.dataGridView1.Rows[index].Cells[0].Value = Valuess;
                        this.dataGridView1.Refresh();
                    }
                    // this.label1.Text = Encoding.Unicode.GetString(a);
                    else
                        Valuess = ss;
                    this.label1.Text = ss;
                }
                catch
                {
                    ;
                }
            }
            ;
        }
        
        private void monitorToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
            this.dataGridView1.DataSource = null;
            this.dataGridView1.Rows.Clear();
            this.dataGridView1.Columns.Clear();
            this.dataGridView1.Columns.Add("Data", "Data");
            this.dataGridView1.Columns.Add("Time", "Time");
            this.dataGridView1.Columns["Time"].Width = 120;
            dataGridView1.RowHeadersWidth = 80;
            if (ThreadRun)
                return;
            if (sPort.IsOpen) return;
     
           // dataGridView1.Controls.
         //   this.dataGridView1.Columns.
            sPort.PortName = "com1";//串口的portname 
            sPort.BaudRate = 9600;//串口的波特率
            sPort.DataBits = 8;
            //两个停止位
            sPort.StopBits = System.IO.Ports.StopBits.One;
            //无奇偶校验位
            sPort.Parity = System.IO.Ports.Parity.None;
            sPort.ReadTimeout = -1;
            sPort.WriteTimeout = -1;
            try
            {
                sPort.Open();
            }catch
                {
                    sPort.Open();
                }

                t = new Thread(WriteY);
                ThreadRun = true;
                t.Start();     
      
        }
        Thread t;
        bool ThreadRun=false;
        private delegate void InvokeHandler();
        private void WriteY()
        {
            while (ThreadRun)
            {
                
                if (sPort.IsOpen)
                {
                    // byte[] data = Convert.FromBase64String(sPort.ReadLine());
                    //  this.label1.Text = Encoding.Unicode.GetString(data);

                    //     int count = sPort.BytesToRead;
                    //     byte[] data = new byte[count];
                    // sPort.Read(data,0,count);
                    // string ss=sPort.ReadTo("\n");
                    // this.label1.Text = Encoding.Unicode.GetString(a);
                    //this.label1.Text = ss;
                    //   this.label1.Text = data.ToString();
                    // this.label1.Text = Encoding.Unicode.GetString(data);
                    //   this.label1.Text = Encoding.ASCII.GetString(data);
                    try
                    {
                        sPort.NewLine = "\r";
                        string ss = sPort.ReadLine();
                        if (ss == "S")
                        {
                            // DataGridViewRow aa = new DataGridViewRow();
                            // this.dataGridView1.Rows.Add();
                            this.Invoke(new InvokeHandler(delegate()
                           {

                               int index = this.dataGridView1.Rows.Add();
                               this.dataGridView1.Rows[index].Cells[0].Value = Valuess;
                               this.dataGridView1.Rows[index].Cells[1].Value = DateTime.Now.ToString();
                                this.dataGridView1.Rows[index].HeaderCell.Value = (index+1).ToString();      
                               if (dataGridView1.RowCount > 4)
                                   this.dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView1.RowCount - 1;
                               // this.dataGridView1.DisplayedRowCount();
                           }));

                            //  this.dataGridView1.Refresh();
                        }
                        // this.label1.Text = Encoding.Unicode.GetString(a);
                        else
                            Valuess = ss;
                        this.label1.Text = ss;
                    }
                    catch
                    {
                        ;
                    }
                    //System.Threading.Thread.Sleep(100);
                }
            }
        }
        private void testGetdataToolStripMenuItem_Click(object sender, EventArgs e)
        {
          //  byte[] data = Convert.FromBase64String(sPort.ReadLine());
            sPort.NewLine ="\r";
            string ss = sPort.ReadLine();
           // this.label1.Text = Encoding.Unicode.GetString(a);
            this.label1.Text = ss;
        }

        private void Form1_Deactivate(object sender, EventArgs e)
        {
            //MessageBox.Show("close");
        }

        private void formclosing(object sender, FormClosingEventArgs e)
        {
           // MessageBox.Show("close");
            ThreadRun = false;
            sPort.Close();         
           
        }
            public void DataToExcel(DataGridView m_DataView)
        {
           SaveFileDialog kk = new SaveFileDialog(); 
            kk.Title = "保存EXECL文件"; 
            kk.Filter = "EXECL文件(*.xls) |*.xls |所有文件(*.*) |*.*"; 
          kk.FilterIndex = 1;
            if (kk.ShowDialog() == DialogResult.OK) 
            { 
                 string FileName = kk.FileName;
                 if (File.Exists(FileName))
                     File.Delete(FileName);
                 FileStream objFileStream; 
                StreamWriter objStreamWriter; 
                string strLine = ""; 
                objFileStream = new FileStream(FileName, FileMode.OpenOrCreate, FileAccess.Write); 
                objStreamWriter = new StreamWriter(objFileStream, System.Text.Encoding.Unicode);
              //   for (int i = 0; i  < m_DataView.Columns.Count; i++) 
              //   { 
                //    if (m_DataView.Columns[i].Visible == true) 
                //     { 
                    //    strLine = strLine + m_DataView.Columns[i].HeaderText.ToString() + Convert.ToChar(9); 
                 //    } 
             //    } 
            //    objStreamWriter.WriteLine(strLine); 
                strLine = ""; 

             for (int i = 0; i  < m_DataView.Rows.Count; i++) 
               { 
                   if (m_DataView.Columns[0].Visible == true) 
                    { 
                       if (m_DataView.Rows[i].Cells[0].Value == null) 
                          strLine = strLine + " " + Convert.ToChar(9); 
                       else 
                           strLine = strLine + m_DataView.Rows[i].Cells[0].Value.ToString() + Convert.ToChar(9); 
                   } 
                    for (int j = 1; j  < m_DataView.Columns.Count; j++) 
                    { 
                        if (m_DataView.Columns[j].Visible == true) 
                        { 
                           if (m_DataView.Rows[i].Cells[j].Value == null) 
                                strLine = strLine + " " + Convert.ToChar(9); 
                            else 
                            { 
                                string rowstr = ""; 
                               rowstr = m_DataView.Rows[i].Cells[j].Value.ToString(); 
                                if (rowstr.IndexOf("\r\n") >  0) 
                                   rowstr = rowstr.Replace("\r\n", " "); 
                               if (rowstr.IndexOf("\t") >  0) 
                                    rowstr = rowstr.Replace("\t", " "); 
                                strLine = strLine + rowstr + Convert.ToChar(9); 
                            } 
                        } 
                    } 
                   objStreamWriter.WriteLine(strLine); 
                    strLine = ""; 
               } 
                objStreamWriter.Close(); 
                objFileStream.Close();
               MessageBox.Show(this,"保存EXCEL成功","提示",MessageBoxButtons.OK,MessageBoxIcon.Information); 
            }
        }
        void saveexcel(DataGridView m_DataView)
        {
            DataTable DT1=GetDgvToTable(m_DataView);
            DT1.TableName="data";
            String sConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
     "Data Source=K:\\work111\\excel1.xls;" +
     "Extended Properties=Excel 8.0;";
            //实例化一个Oledbconnection类(实现了IDisposable,要using)
                using (OleDbConnection ole_conn = new OleDbConnection(sConnectionString))
                {
                    ole_conn.Open();
                 using (OleDbCommand ole_cmd = ole_conn.CreateCommand())
                   {
                  //     ole_cmd.CommandText = "CREATE TABLE data ([Data] VarChar,[Time] VarChar)";
                   //   ole_cmd.ExecuteNonQuery();
                     //   ole_cmd.CommandText = "insert into data values('DJ001','点击科技')";
                     //   ole_cmd.ExecuteNonQuery();
                       // ole_cmd.up
                       // MessageBox.Show("生成Excel文件成功并写入一条数据......");
                    }

                    string strCom = " SELECT * FROM data";

                  OleDbDataAdapter myDA = new OleDbDataAdapter(strCom, ole_conn); ;
                   // SqlDataAdapter myDA = new SqlDataAdapter(strCom, sConnectionString);
                   // SqlCommandBuilder cbUpdate = new SqlCommandBuilder(myDA);
                  OleDbCommandBuilder cb = new OleDbCommandBuilder(myDA);
                 // myDA.UpdateCommand = cb.GetUpdateCommand(); 
                    myDA.Update(DT1);
                }

        }
        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
           // DataToExcel(this.dataGridView1);
            saveexcel(this.dataGridView1);
        }

        private void fileFToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
    }
}