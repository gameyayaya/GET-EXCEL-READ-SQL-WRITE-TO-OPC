using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using OPCAutomation;
using System.Data.SqlClient;





namespace NPOI_DEMO
{

    public partial class Form1 : Form
    {

        OPCServer KepServer;
        OPCGroups KepGroups;
        OPCGroup KepGroup;
        OPCItems KepItems;
        DataTable dt;
        ISheet mysheet;

        public Form1()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
        }

        // 模版【LOG UI 使用textbox1】
        private void SendMsg(string m)
        {
            this.textBox1.Text += "\r\n\r\n" + System.DateTime.Now.ToString() + " : " + m;
            this.textBox1.SelectionStart = this.textBox1.Text.Length;
            this.textBox1.ScrollToCaret();
        }


        //  模版【读取excel D:\demo.xlsx to datagridview1】
        public void readexcel()
        {
            try
            {
                string thefilefullpath = @"D:\demo.xlsx";

                // 开一个虚拟excel
                IWorkbook myexcel;

                SendMsg("开始读取xlsx文件:  " + thefilefullpath + "  到记忆体myexcel");
                // 把excel  读取到记忆体myexcel
                using (FileStream R = new FileStream(thefilefullpath, FileMode.Open, FileAccess.Read))
                {
                    myexcel = new XSSFWorkbook(R);
                }

                SendMsg("开始把myexcel的第0个sheet:  " + myexcel.GetSheetAt(0).SheetName + "  读取到记忆体mysheet");
                // 设定记忆体里的excel , 指定起始位置
                mysheet = myexcel.GetSheetAt(0); // zero-based index of your target sheet


                SendMsg("New一个空的DataTable:dt来装这个sheet");
                // 把myexcel的sheet  读取到记忆体dt
                dt = new DataTable(mysheet.SheetName);

                SendMsg("把UI上的dataGridView1资料来源指定为dt , 这样dt有异动就会在UI上及时显示");
                // 把显示表格的资料来源设定为dt
                dataGridView1.DataSource = dt;

                SendMsg("开始把excel刷到dt里面 , 先刷标题栏");
                // 读取excel标题栏  写到datatable
                IRow headerRow = mysheet.GetRow(0);
                foreach (ICell headerCell in headerRow)
                {
                    dt.Columns.Add(headerCell.ToString());
                }

                SendMsg("标题刷完了 , 开始从第一行开始把资料刷到dt");
                // 读取excel内容  一行一行写到datatable , 写的是++  所以内容从第一行开始读 而不是第零行
                int rowIndex = 0;
                foreach (IRow row in mysheet)
                {
                    // skip header row
                    if (rowIndex++ == 0) continue;
                    DataRow dataRow = dt.NewRow();
                    dataRow.ItemArray = row.Cells.Select(c => c.ToString()).ToArray();
                    dt.Rows.Add(dataRow);
                }

                SendMsg("完成 : excel已经成功读取到记忆体里");
            }
            catch (Exception err)
            {
                SendMsg( "错误提示 : " + err.Message);
                SendMsg("程式已经停止运行 , 请记录错误讯息后关闭程式 , 启动程式之前  请确认已排除上列错误");
                //throw;
            }
        }


        //  模版【OPC 更新event 模版】
        private void KepGroup_DataChange(int TransactionID, int NumItems, ref Array ClientHandles, ref Array ItemValues, ref Array Qualities, ref Array TimeStamps)
        {
            try
            {
                string s = "共有" + NumItems + "个数据回传";

                foreach (DataRow row in dt.Rows)
                {

                    //SendMsg("Add opcItem  ID[" + row["Num"].ToString() + "] : " + row["ItemID"].ToString());
                    //KepItems.AddItem(row["ItemID"].ToString().Replace(" ", ""), Convert.ToInt32(row["Num"]));

                    for (int i = 1; i <= NumItems; i++)
                    {
                        if (Convert.ToInt32(row["Num"]) == Convert.ToInt32(ClientHandles.GetValue(i)))
                        {
                            row.SetField("Value", ItemValues.GetValue(i));
                            row.SetField("UpdateTime" , System.DateTime.Now.ToString());
                            //writesql(row["ItemID"].ToString(), row["Value"].ToString(), Convert.ToDateTime(row["UpdateTime"]));
                            s = s + "\r\n" + "ID : " + ClientHandles.GetValue(i) + "\t" + "CurrentValue : " + ItemValues.GetValue(i);
                        }

                    }
                }

                
                SendMsg(s);
            }
            catch (Exception e)
            {
                SendMsg("KepGroup_DataChange" + e.Message);
            }
        }

        //  模版【连接OPC 建立GROUP 从dt塞Item】
        public void readopc()
        {
            try
            {

                SendMsg("开始连接OPC SERVER");
                KepServer = new OPCServer();
                Object servers = KepServer.GetOPCServers("localhost");
                KepServer.Connect("KEPware.KEPServerEx.V4", "localhost");

                SendMsg("连接完成 , 开始建立 OPC Client的GROUP  用来装OPCITEM");
                KepGroups = KepServer.OPCGroups;
                KepGroup = KepGroups.Add("OpcGroup");
                KepServer.OPCGroups.DefaultGroupIsActive = true;
                KepServer.OPCGroups.DefaultGroupDeadband = 0;
                KepGroup.UpdateRate = 1000;
                

                SendMsg("GROUP建立成功 , 开始从dt把ITEMID刷到OPC GROUP物件里面");
                KepItems = KepGroup.OPCItems;
                foreach (DataRow row in dt.Rows)
                {
                    
                    SendMsg("Add opcItem  ID[" + row["Num"].ToString()  + "] : "+ row["ItemID"].ToString());
                    KepItems.AddItem(row["ItemID"].ToString().Replace(" ", ""), Convert.ToInt32(row["Num"]));

                }

                foreach (OPCItem item in KepItems)
                {

                    foreach (DataRow row in dt.Rows)
                    {
                        if (item.ClientHandle.ToString() == row["Num"].ToString())
                        {
                            row.SetField("Serverhandle",item.ServerHandle);
                        }
                       

                        
                    }

                }

                

                SendMsg("OPC物件刷完了 , 开始刷资料到dt");
                KepGroup.IsSubscribed = true;
                KepGroup.IsActive = true;

                //System.Threading.Thread.Sleep(3000);

                KepGroup.DataChange += new DIOPCGroupEvent_DataChangeEventHandler(KepGroup_DataChange);

                
                KepGroup.AsyncWriteComplete += KepGroup_AsyncWriteComplete;
                //Task.Factory.StartNew(button1_Click);
                
            }
            catch (Exception err)
            {
                SendMsg("错误提示 : " + err.Message);
                //throw e;
            }
        }

        private void KepGroup_AsyncWriteComplete(int TransactionID, int NumItems, ref Array ClientHandles, ref Array Errors)
        {

            SendMsg("KepGroup_AsyncCancelComplete : 非同步写入已完成");
            //throw new NotImplementedException();
        }


        private void KepGroup_AsyncReadComplete(int TransactionID, int NumItems, ref Array ClientHandles, ref Array ItemValues, ref Array Qualities, ref Array TimeStamps, ref Array Errors)
        {
            try
            {
                string s = "共有" + NumItems + "个数据回传";

                foreach (DataRow row in dt.Rows)
                {

                    //SendMsg("Add opcItem  ID[" + row["Num"].ToString() + "] : " + row["ItemID"].ToString());
                    //KepItems.AddItem(row["ItemID"].ToString().Replace(" ", ""), Convert.ToInt32(row["Num"]));

                    for (int i = 1; i <= NumItems; i++)
                    {
                        if (Convert.ToInt32(row["Num"]) == Convert.ToInt32(ClientHandles.GetValue(i)))
                        {
                            row.SetField("Value", ItemValues.GetValue(i));
                            row.SetField("UpdateTime", System.DateTime.Now.ToString());
                            //writesql(row["ItemID"].ToString(), row["Value"].ToString(), Convert.ToDateTime(row["UpdateTime"]));
                            s = s + "\r\n" + "ID : " + ClientHandles.GetValue(i) + "\t" + "CurrentValue : " + ItemValues.GetValue(i);
                        }

                    }
                }


                SendMsg(s);
            }
            catch (Exception e)
            {
                SendMsg("KepGroup_DataChange" + e.Message);
            }
        }


        
        //  模版【SQL insert】
        public void writesql(string ItemID , string Value , DateTime UpdateTime)
        {

            try
            {
                string connString = "Data Source=fmcsweb;Initial Catalog=FMCSDB;Integrated Security=False;User ID=sa;Password=P@ssw0rd;Connect Timeout=30;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False";
                string query = "INSERT INTO dbo.P1FMCS_TXLOG(ItemID, Value, UpdateTime) VALUES(@1, @2, @3)";
               
                using (SqlConnection conn = new SqlConnection(connString))
                using (SqlCommand command = new SqlCommand(query, conn))
                {
                    command.Parameters.Add("@1", SqlDbType.NVarChar).Value = ItemID;
                    command.Parameters.Add("@2", SqlDbType.NVarChar).Value = Value;
                    command.Parameters.Add("@3", SqlDbType.DateTime).Value = UpdateTime;

                    //make sure you open and close(after executing) the connection
                    conn.Open();
                    command.ExecuteNonQuery();
                    conn.Close();
                }
            }
            catch (Exception err)
            {

                SendMsg("错误提示 : " + err.Message);
                //throw e;
            }
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            readexcel();
            readopc();
            //Task.Factory.StartNew(readopc);
            //Task.Factory.StartNew(readopc);
            //readopc();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {

                string s = "开始把dt写到opc server";

                //把datatable特定栏位  搞出来成为array的方法LINQ
                string[] a1 = dt.AsEnumerable().Select(r => r.Field<string>("Serverhandle")).ToArray();
                int[] a2 = a1.Select(int.Parse).ToArray();
                object[] a3 = dt.AsEnumerable().Select(r => r.Field<string>("Value")).ToArray();

                int[] k1 = new int[1] { 0 };
                object[] k2 = new object[1] { "" };


                Array a5 = (Array)k1.Concat(a2).ToArray();
                Array a6 = (Array)k2.Concat(a3).ToArray();

              
                SendMsg("写入开始");
                // AsyncWrite这个方法有几个坑 , 简单的说 他就是吃 两个array , 把array的东西更新到 OPC server那边
                // 首先他array处理是base 1的 , 然后c#的toarray生出来的是base 0 , 直接刷进去会有bug , 所以toarray弄出来的要自己前面多塞一个 , 可是这样length会多一个 , 记得要减掉
                // 参数1 : 重要__有几个 , 直接取array长度就好
                // 参数2 : 重要__要给他 object[] , 实际上是要塞int array , 内容是每个 OpcItem 的 Serverhandle属性 , serverhandle是系统自己生的 , 跟clienthandle不一漾 
                // 参数3 : 重要__要给他 object[] , 实际上是要塞string array , 内容就是 要更新的数据 , itemvalue
                // 参数4 : 随便塞个空array
                // 参数5 : 随便塞个数字
                // 参数6 : 随便塞个int变数
                KepGroup.AsyncWrite(a5.Length -1 ,ref a5,ref a6,out Array a7,2009,out int i2);
                GC.Collect();

               
             

                
                //下面是不使用asyncwrite 泻入方法 , 扫一次table , 把所有item写一次 , 比较慢
                //foreach (DataRow row in dt.Rows)
                //{
                //    KepItems.GetOPCItem(Convert.ToInt32(row["Serverhandle"])).Write(row["Value"].ToString());
                //    
                //}

                SendMsg("写入结束");
            }
            catch (Exception err)
            {
                
                SendMsg("KepGroup_DataChange" + err.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            foreach (OPCItem item in KepItems)
            {

                foreach (DataRow row in dt.Rows)
                {
                    if (item.ClientHandle.ToString() == row["Num"].ToString())
                    {
                        row.SetField("Serverhandle", item.ServerHandle);
                    }



                }

            }
        }
    }
}
