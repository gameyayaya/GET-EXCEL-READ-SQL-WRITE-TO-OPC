// pull test desu
// change
// 要加参考 OPCDAAUTO
// yoyoyo

// dddddddddddddddddddddddddddddddddddddddddddd

// AsyncWrite这个方法有几个坑 , 简单的说 他就是吃 两个array , 把array的东西更新到 OPC server那边
// 首先他array处理是base 1的 , 然后c#的toarray生出来的是base 0 , 直接刷进去会有bug , 所以toarray弄出来的要自己前面多塞一个 , 可是这样length会多一个 , 记得要减掉
// 参数1 : 重要__有几个 , 直接取array长度就好
// 参数2 : 重要__要给他 object[] , 实际上是要塞int array , 内容是每个 OpcItem 的 Serverhandle属性 , serverhandle是系统自己生的 , 跟clienthandle不一漾 
// 参数3 : 重要__要给他 object[] , 实际上是要塞string array , 内容就是 要更新的数据 , itemvalue
// 参数4 : 随便塞个空array
// 参数5 : 随便塞个数字
// 参数6 : 随便塞个int变数


//下面是不使用asyncwrite 泻入方法 , 扫一次table , 把所有item写一次 , 比较慢  慢很多 不建议使用
//foreach (DataRow row in dt.Rows)
//{
//    KepItems.GetOPCItem(Convert.ToInt32(row["Serverhandle"])).Write(row["Value"].ToString());
//    
//}


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

        // 用来存 key : ItemID 跟 Value : Serverhandle 的 , 主要是会快很多 , 比启用foreach刷到dt里面而言 , 同时避免双层for
        Dictionary<string, int> MyDic = new Dictionary<string, int>();

        OPCServer KepServer;
        OPCGroups KepGroups;
        OPCGroup KepGroup;
        OPCItems KepItems;
        ISheet mysheet;

        DataTable dt;
        DataTable sqldt = new DataTable();
        private Timer timer1;

        public Form1()
        {
            InitializeComponent();

            //跨线程使用变数要加这一条才不会出错 , 官方不建议
            CheckForIllegalCrossThreadCalls = false;

            Initial_mynotify();
            readexcel();
            readopc();
            InitTimer();
            
        }










        // 模版【timer trigger】, 在 Form 呼叫一次后持续LOOP
        public void InitTimer()
        {
            try
            {
                timer1 = new Timer();
                timer1.Tick += new EventHandler(timer1_Tick);
                timer1.Interval = 3000; // in miliseconds
                timer1.Start();
            }
            catch (Exception err)
            {

                SendMsg("err【InitTimer()】:" + err.Message);
            }
            
        }

        // 模版【timer trigger】, trigger by timer1
        private void timer1_Tick(object sender, EventArgs e)
        {
            try
            {
                // 用task开线程 , 这样UI才不会卡住
                Task.Factory.StartNew(readsqltoopc);
            }
            catch (Exception err)
            {

                SendMsg("err【timer1_Tic()】:" + err.Message);
            }
            
            //readsqltoopc();
        }

        // 模版【LOG 输出 】使用textbox1
        private void SendMsg(string m)
        {
            this.textBox1.Text += "\r\n\r\n" + System.DateTime.Now.ToString() + " : " + m;
            this.textBox1.SelectionStart = this.textBox1.Text.Length;
            this.textBox1.ScrollToCaret();
        }

        //  模版【excel 读取 】"D:\demo.xlsx" to datagridview1
        public void readexcel()
        {
            try
            {
                //绝对路径的方法
                //string thefilefullpath = @"D:\demo.xlsx";

                //相对路径的方法 , 把东西跟执行档放一起
                string thefilefullpath = Application.StartupPath + "\\data.xlsx";
                
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
                    //SendMsg(rowIndex.ToString());
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



// OPC 模版 ---------------------------------------------------------------------------------------------------


        //  模版【OPC 初始化 OPC连接 建立GROUP 塞Item @dt 建字典 启动OPC更新event】
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
                    
                    //SendMsg("Add opcItem  ID[" + row["Num"].ToString()  + "] : "+ row["ItemID"].ToString());
                    KepItems.AddItem(row["ItemID"].ToString().Replace(" ", ""), Convert.ToInt32(row["Num"]));

                }
                SendMsg("共加入" + dt.Rows.Count.ToString() + "个OPCITEM");

                //建字典 , 内容是所有excel里面的点 , 之后可以查字典确认有没有点表里面 , 确保 SQL 上是点表有的数据 , 刷到OPC server时才不会出错
                foreach (OPCItem item in KepItems)
                {
                    MyDic.Add(item.ClientHandle.ToString() , item.ServerHandle);
                }

                SendMsg("OPC物件刷完了 , 开始刷资料到dt");
                KepGroup.IsSubscribed = true;
                KepGroup.IsActive = true;
                KepGroup.DataChange += new DIOPCGroupEvent_DataChangeEventHandler(KepGroup_DataChange);
                KepGroup.AsyncWriteComplete += KepGroup_AsyncWriteComplete;        
            }
            catch (Exception err)
            {
                SendMsg("错误提示 : " + err.Message);
            }
        }

        // 模版 【OPC 写入完成 event】
        private void KepGroup_AsyncWriteComplete(int TransactionID, int NumItems, ref Array ClientHandles, ref Array Errors)
        {

            SendMsg("KepGroup_AsyncCancelComplete : 非同步写入已完成");
            //throw new NotImplementedException();
        }

        //  模版【OPC 更新事件  】把资料刷到datatable , 让UI上看的到 
        private void KepGroup_DataChange(int TransactionID, int NumItems, ref Array ClientHandles, ref Array ItemValues, ref Array Qualities, ref Array TimeStamps)
        {
            try
            {

                // 用来存 key : ClientHandles 跟 Value : ItemValues 的 
                Dictionary<string, string> MyDicValue = new Dictionary<string, string>();
                SendMsg("共有" + NumItems + "个数据回传");

                //建这一组更新array的专用字典
                for (int i = 1; i <= NumItems; i++)
                {
                    MyDicValue.Add(ClientHandles.GetValue(i).ToString(), ItemValues.GetValue(i).ToString());
                }

                //然后每个row查看有在字典里面的就更新 , 会比不用字典的用两层for快
                foreach (DataRow row in dt.Rows)
                {
                    //i += 1;
                    if (MyDicValue.ContainsKey(row["Num"].ToString()))
                    {
                        row.SetField("Value", MyDicValue[row["Num"].ToString()]);
                        row.SetField("UpdateTime", System.DateTime.Now.ToString());
                    }
                }

                //第一次会刷所有的数值近dt , 所以刷完后在给 dataGridView , 不然冷启动刷资料进来会时会超级慢
                dataGridView1.DataSource = dt;

            }
            catch (Exception e)
            {
                SendMsg("error【KepGroup_DataChange】:" + e.Message);
            }
        }

        // 模版【OPC 写入方法 数据来源是datatable】
        private void Read_dt_to_opc(object sender, EventArgs e)
        {
            try
            {



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
                KepGroup.AsyncWrite(a5.Length - 1, ref a5, ref a6, out Array a7, 2009, out int i2);
                //GC.Collect();

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



// SQL 模版 ---------------------------------------------------------------------------------------------------
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

        //  模版【SQL select * to datatable】
        public void readsqltoopc()
        {
            try
            {
                string connString = "Data Source=fmcsweb;Initial Catalog=DOOR;Integrated Security=False;User ID=sa;Password=P@ssw0rd;Connect Timeout=30;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False";
                string query = "select * from  [DOOR].[dbo].[LSS]";

                int[] k1 = new int[1] { 0 };
                object[] k2 = new object[1] { "" };

                SqlConnection conn = new SqlConnection(connString);
                SqlCommand command = new SqlCommand(query, conn);
                SqlDataAdapter da = new SqlDataAdapter(command);

                conn.Open();
                    command.ExecuteNonQuery();
                    sqldt.Clear();
                    da.Fill(sqldt);
                conn.Close();

                SendMsg("SQL读取完成 : " + sqldt.Rows.Count.ToString() + " 组数据已读取 ");
                string[] a2 = sqldt.AsEnumerable().Select(r => r.Field<string>("Num")).ToArray();
                object[] a3 = sqldt.AsEnumerable().Select(r => r.Field<bool>("Value").ToString()).ToArray();
                int[] aa = new int[a2.Length];

                //把 Num 字串转换成 serverhandle给AsyncWrite用
                for (int i = 1; i <= a2.Length; i++)
                {
                    aa[i-1] = MyDic[a2.GetValue(i-1).ToString()];
                }
                
                Array a5 = (Array)k1.Concat(aa).ToArray();
                Array a6 = (Array)k2.Concat(a3).ToArray();
                KepGroup.AsyncWrite(a5.Length - 1, ref a5, ref a6, out Array a7, 2009, out int i2);
            }
            catch (Exception err)
            {
                SendMsg("错误提示 : " + err.Message);
            }
        }

        // 模版【SQL update from dt】
        public void writesqlserverhandle()
        {
            try
            {
                string connString = "Data Source=fmcsweb;Initial Catalog=DOOR;Integrated Security=False;User ID=sa;Password=P@ssw0rd;Connect Timeout=30;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False";
                //string query = "UPDATE dbo.LSS(ItemID, Value, UpdateTime) VALUES(@1, @2, @3)";
                string query = "UPDATE LSS SET Serverhandle = @1  Where Num = @2";
                using (SqlConnection conn = new SqlConnection(connString))
                    foreach (DataRow row in dt.Rows)
                    {
                        using (SqlCommand command = new SqlCommand(query, conn))
                        {
                            command.Parameters.Add("@1", SqlDbType.NVarChar).Value = row["Serverhandle"].ToString();
                            command.Parameters.Add("@2", SqlDbType.NVarChar).Value = row["Num"].ToString();
                            conn.Open();
                            command.ExecuteNonQuery();
                            conn.Close();
                        }
                    }
            }
            catch (Exception err)
            {

                SendMsg("错误提示 : " + err.Message);
                //throw e;
            }
        }




// 右下 Icon化 防呆关闭 模版---------------------------------------------------------------------------------------

        //右下Icon , Driver预设不要show Taskbar , 容易被误关 , 等于是实现背景执行
        private System.Windows.Forms.NotifyIcon notifyIcon1;
        private System.Windows.Forms.ContextMenu contextMenu1;
        private System.Windows.Forms.MenuItem menuItem1;

        //右下小图示:初始化模版 , Form1 要引用
        private void Initial_mynotify()
        {
            this.components = new System.ComponentModel.Container();
            this.contextMenu1 = new System.Windows.Forms.ContextMenu();
            this.menuItem1 = new System.Windows.Forms.MenuItem();

            // Initialize contextMenu1
            this.contextMenu1.MenuItems.AddRange(
                        new System.Windows.Forms.MenuItem[] { this.menuItem1 });

            // Initialize menuItem1
            // 定义 : 右下哐哐 , 右键选单
            this.menuItem1.Index = 0;
            this.menuItem1.Text = "E&xit";
            this.menuItem1.Click += new System.EventHandler(this.menuItem1_Click);

            // Create the NotifyIcon.
            this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);

            // 定义ico图 , 要放在debug目录
            notifyIcon1.Icon = new Icon("Icon.ico");

            // The ContextMenu property sets the menu that will
            // appear when the systray icon is right clicked.
            notifyIcon1.ContextMenu = this.contextMenu1;

            // The Text property sets the text that will be displayed,
            // in a tooltip, when the mouse hovers over the systray icon.
            notifyIcon1.Text = "Status : Driver is Running Background";
            notifyIcon1.Visible = true;

            // Handle the DoubleClick event to activate the form.
            notifyIcon1.DoubleClick += new System.EventHandler(this.notifyIcon1_DoubleClick);
        }

        //右下小图示:点两下弹出
        private void notifyIcon1_DoubleClick(object Sender, EventArgs e)
        {
            // Show the form when the user double clicks on the notify icon.

            // Set the WindowState to normal if the form is minimized.
            if (this.WindowState == FormWindowState.Minimized)
                this.WindowState = FormWindowState.Normal;

            // Activate the form.
            this.Activate();
        }

        //右下小图示:右键选单关闭
        private void menuItem1_Click(object Sender, EventArgs e)
        {
            // Close the form, which closes the application.
            this.Close();
        }

        //____样板 : 误关闭防呆机制
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (MessageBox.Show("Are you want to close Driver ?", "", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                string input = Microsoft.VisualBasic.Interaction.InputBox("Please enter 'exit' to close driver ", "Double check", "", 0, 0);
                if (input == "exit")
                {
                    MessageBox.Show("Shutdown OK!!");
                }
                else
                {
                    MessageBox.Show("keyword is not match 'exit' ");
                    e.Cancel = true;
                }
            }
            else
            {
                e.Cancel = true;
            }
        }


    }

}
