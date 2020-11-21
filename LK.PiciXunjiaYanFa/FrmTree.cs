using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using fuzhu;

namespace LKU8.shoukuan
{
    public partial class FrmTree : Form
    {
        private DataTable dtTree = null;

        private DataView dv = null;
        string cInvcode, cInvname, cInvstd, cInvaddcode;
        public FrmTree()
        {
            InitializeComponent();
        }
        public FrmTree(string cinvcode,string cinvname,string cinvstd)
        {
            InitializeComponent();
            cInvcode = cinvcode;
            cInvname = cinvname ;
            cInvstd = cinvstd;
            label1.Text = cInvcode;
            label3.Text = cInvname;
            label5.Text = cInvstd;
            ExtensionMethods.DoubleBuffered(dataGridView1, true);
            dataGridView1.AutoGenerateColumns = false;
        }

        public FrmTree(string cinvcode, string cinvname, string cinvstd,string cinvaddcode)
        {
            InitializeComponent();
            cInvcode = cinvcode;
            cInvname = cinvname;
            cInvstd = cinvstd;
            cInvaddcode = cinvaddcode;
            label1.Text = cInvcode;
            label3.Text = cInvname;
            label5.Text = cInvstd;
            ExtensionMethods.DoubleBuffered(dataGridView1, true);
            dataGridView1.AutoGenerateColumns = false;
        }

        #region 加载
        private void FrmTree_Load(object sender, EventArgs e)
        {
            try
            {
                //zdy_zd_sp_xzqs](@dt1 datetime ,@dt2 datetime ,@lx varchar(20))

                //            string sql =string.Format(@"select a.ddate,a.cmaker,b.gys,a.fpxjrxm,a.cinvcode,a.cinvstd,a.cinvname,a.cqty1,a.cqty2,a.cqty3,b.xunjia1,b.xunjia2,b.xunjia3,b.bz1,b.bz2,b.bz3  from zdy_lk_xunjia a,zdy_lk_xunjias b 
                //               where a.id = b.id and a.id in (select top 5 id from zdy_lk_xunjia where  cinvcode = '{0}' or cinvaddcode = '{1}' or cinvcode = '{1}' or cinvaddcode = '{0}' order by id desc)",cInvcode,cInvaddcode);

                //201903013 更改为全部显示
                string sql = string.Format(@"select a.ddate,a.cmaker,b.gys,a.fpxjrxm,a.cinvcode,a.cinvstd,a.cinvname,a.cqty1,a.cqty2,a.cqty3,b.xunjia1,b.xunjia2,b.xunjia3,b.bz1,b.bz2,b.bz3  from zdy_lk_xunjia a,zdy_lk_xunjias b 
                           where a.byanfa = 1 and a.id = b.id and (cinvcode = '{0}' or cinvaddcode = '{1}' or cinvcode = '{1}' or cinvaddcode = '{0}' )", cInvcode, cInvaddcode);


                DataTable dtmx = DbHelper.ExecuteTable(sql);
                dataGridView1.DataSource = dtmx;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;

            }

        }
        #endregion

    

   
    }
}
