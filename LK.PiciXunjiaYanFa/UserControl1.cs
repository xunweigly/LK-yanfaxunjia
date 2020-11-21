using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using fuzhu;
using ADODB;
using MSXML2;
using UFIDA.U8.U8APIFramework;
using UFIDA.U8.U8MOMAPIFramework;
using UFIDA.U8.U8APIFramework.Parameter;
using System.Threading;
using System.Data.SqlClient;
using Process;
using UFIDA.U8.Pub.FileManager;
using System.IO;
using System.Net;

namespace LKU8.shoukuan
{
    public partial class UserControl1 : UserControl
    {


        DataTable dtXunjia, dtXunjiaPu;
        string sColname;
        int iks = 0;
        string Lujing;
        string cFtp;
        public UserControl1()
        {
            InitializeComponent();

            ExtensionMethods.DoubleBuffered(dataGridView1, true);
            ExtensionMethods.DoubleBuffered(dataGridView2, true);
        }







        #region 加载
        private void UserControl1_Load(object sender, EventArgs e)
        {
            try
            {
                dataGridView2.AutoGenerateColumns = false;
                dataGridView1.AutoGenerateColumns = false;

                string sql = "select 0 as xz,* from zdy_lk_xunjia where byanfa = 1 and cyanfazt = '已提交' ";
                dtXunjia = DbHelper.ExecuteTable(sql);
                dataGridView1.DataSource = dtXunjia;

                comboBox1.Text = "已提交";

                sql = "select cvalue from zdy_lk_para where lx = '101'";
                cFtp = DbHelper.GetDbString(DbHelper.ExecuteScalar(sql));

                Dgvfuzhu.BindReadDataGridViewStyle(this.Name, dataGridView1); // 初始化布局
                Dgvfuzhu.BindReadDataGridViewStyle(this.Name, dataGridView2); // 初始化布局


                //MessageBox.Show(cFtp);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            
            }

           



            //if (canshu.cQx == "1" || canshu.cQx == "2" || canshu.userName == "demo")
            //{
            //    dataGridView1.Columns["fpxjr"].Visible = true;

            //}

        }
        #endregion


        #region 写序号
        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            var dgv = sender as DataGridView;
            if (dgv != null)
            {
                Rectangle rect = new Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, dgv.RowHeadersWidth - 4, e.RowBounds.Height);
                TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), dgv.RowHeadersDefaultCellStyle.Font, rect, dgv.RowHeadersDefaultCellStyle.ForeColor, TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
            }
        }

        #endregion


        #region 布局设置
        public void SaveBuju()
        {
            Dgvfuzhu.SaveDataGridViewStyle(this.Name, dataGridView1);
            Dgvfuzhu.SaveDataGridViewStyle(this.Name, dataGridView2);
            MessageBox.Show("布局保存成功！");
        }

        public void DelBuju()
        {
            Dgvfuzhu.deleteDataGridViewStyle(this.Name, dataGridView1);
            Dgvfuzhu.deleteDataGridViewStyle(this.Name, dataGridView2);
            //Dgvfuzhu.BindReadDataGridViewStyle(this.Name, dataGridView1);
            MessageBox.Show("请关掉界面重新打开，恢复初始布局！");
        }
        #endregion



        #region 增加
        public void Add()
        {
            try
            {
                dataGridView2.EndEdit();
                dataGridView2.AllowUserToAddRows = true;
                if (label6.Text == "")
                {
                    MessageBox.Show("没有选中询价行，无法增加");
                    return;

                }
                DataRow dr = dtXunjiaPu.NewRow();
                dr["ddate"] = DateTime.Now.ToString("yyyy-MM-dd");
                dr["id"] = label6.Text;
                dr["gys"] = "研发";
                dr["bYanfaQuery"] = "True";
                dtXunjiaPu.Rows.Add(dr);
                dataGridView2.AllowUserToAddRows = false;
                iks = 1;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }
        #endregion

        #region 删除
        public void Del()
        {
            DialogResult result = CommonHelper.MsgQuestion("确认要删除选中行吗？");
            if (result == DialogResult.Yes)
            {
                try
                {
                    int i = dataGridView2.CurrentRow.Index;
                    if (i < 0)
                    {
                        MessageBox.Show("没有选择删除行");
                        return;
                    }

                    string sId = DbHelper.GetDbString(dataGridView2.Rows[i].Cells["autoid"].Value);
                    if (sId != "")
                    {
                        string sql = " delete from zdy_lk_xunjias where autoid=@Id ";
                        DbHelper.ExecuteNonQuery(sql, new SqlParameter[] { new SqlParameter("@Id", sId) });
                    }
                    else
                    {
                        dtXunjiaPu.Rows.RemoveAt(i);
                    }




                }

                catch (Exception ex)
                {
                    //DbHelper.RollbackAndCloseConnection(tran);
                    CommonHelper.MsgError("删除失败！原因：" + ex.Message);
                }
                CommonHelper.MsgInformation("删除完成！");
                dtXunjiaPu = DbHelper.ExecuteTable("select id,autoid,ddate,gys,xunjia1,xunjia2,xunjia3,bz1,bz2,bz3 from zdy_lk_xunjias where id ='" + label6.Text + "'");
                dataGridView2.DataSource = dtXunjiaPu;
            }

        }
        #endregion

        #region 保存
        public void Save2()
        {
            Save();
            MessageBox.Show("保存完成");
            return;

        }

        public void Save()
        {
            txtcinvcode.Focus();
            //dataGridView2.CommitEdit((DataGridViewDataErrorContexts)123);

            //MessageBox.Show(dataGridView2.Rows[1].Cells["xunjia3"].Value.ToString());
            //MessageBox.Show(dataGridView2.Rows[1].Cells["autoid"].Value.ToString());
            //return;
            try
            {

                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {


                    //没id，自动保存。有id，判断是否modifyed，如果更改了，update
                    string autoId = DbHelper.GetDbString(dataGridView2.Rows[i].Cells["autoid"].Value);
                    if (string.IsNullOrEmpty(autoId) || autoId == "")
                    {
                        string sql = @" Insert Into zdy_lk_xunjias(id,ddate,gys,xunjia1,xunjia2,xunjia3,bz1,bz2,bz3,bYanfaQuery,ccost) 
                    values(@id,@ddate,@gys,@xunjia1,@xunjia2,@xunjia3,@bz1,@bz2,@bz3,@bYanfaQuery,@ccost) select @@identity ";
                        object obj = DbHelper.GetSingle(sql, new SqlParameter[]{ 
                              new SqlParameter("@id", dataGridView2.Rows[i].Cells["id2"].Value), 
                             new SqlParameter("@ddate", dataGridView2.Rows[i].Cells["ddate2"].Value),
                             new SqlParameter("@gys", dataGridView2.Rows[i].Cells["gys"].Value),
                             new SqlParameter("@xunjia1", dataGridView2.Rows[i].Cells["xunjia1"].Value),
                             new SqlParameter("@xunjia2", dataGridView2.Rows[i].Cells["xunjia2"].Value),
                             new SqlParameter("@xunjia3", dataGridView2.Rows[i].Cells["xunjia3"].Value),
                             new SqlParameter("@bz1", dataGridView2.Rows[i].Cells["bz1"].Value),
                             new SqlParameter("@bz2",dataGridView2.Rows[i].Cells["bz2"].Value),
                             new SqlParameter("@bz3", dataGridView2.Rows[i].Cells["bz3"].Value),
                             new SqlParameter("@bYanfaQuery", dataGridView2.Rows[i].Cells["bYanfaQuery"].Value),
                             new SqlParameter("@ccost", dataGridView2.Rows[i].Cells["ccost"].Value)

                            });
                        //数据表赋值
                        dtXunjiaPu.Rows[i]["autoid"] = Convert.ToInt32(obj);
                        // 设置为非更改状态

                    }
                    else
                    {


                        string sql = @" update zdy_lk_xunjias
                        set gys = @gys,ddate =@ddate,xunjia1 = @xunjia1 ,xunjia2 = @xunjia2,xunjia3 =@xunjia3,bz1=@bz1,bz2= @bz2,bz3= @bz3,
bYanfaQuery=@bYanfaQuery,ccost = @ccost
                         where autoid = @autoid  ";
                        DbHelper.ExecuteNonQuery(sql, new SqlParameter[]{ 
                             new SqlParameter("@gys", dataGridView2.Rows[i].Cells["gys"].Value), 
                             new SqlParameter("@ddate", dataGridView2.Rows[i].Cells["ddate2"].Value),
                             new SqlParameter("@xunjia1", dataGridView2.Rows[i].Cells["xunjia1"].Value),
                             new SqlParameter("@xunjia2", dataGridView2.Rows[i].Cells["xunjia2"].Value),
                             new SqlParameter("@xunjia3", dataGridView2.Rows[i].Cells["xunjia3"].Value),
                             new SqlParameter("@bz1", dataGridView2.Rows[i].Cells["bz1"].Value),
                             new SqlParameter("@bz2", dataGridView2.Rows[i].Cells["bz2"].Value),
                             new SqlParameter("@bz3", dataGridView2.Rows[i].Cells["bz3"].Value),
                              new SqlParameter("@bYanfaQuery", dataGridView2.Rows[i].Cells["bYanfaQuery"].Value),
                             new SqlParameter("@ccost", dataGridView2.Rows[i].Cells["ccost"].Value),
                             new SqlParameter("@autoid", autoId)
                            });

                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }

        }
        #endregion



        #region 提交
        public void Tijiao()
        {

            dataGridView2.Update();
            dataGridView2.EndEdit();

            try
            {

                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {

                    //判断是否输入cas和数量，没输入的不允许进行提交
                    if (DbHelper.GetDbString(dataGridView2.Rows[i].Cells["xunjia1"].Value) == "")
                    {
                        MessageBox.Show("第" + i.ToString() + "行没有输入询价结果，无法提交");
                        continue;
                    }

                    //没id，提示保存
                    string autoId = DbHelper.GetDbString(dataGridView2.Rows[i].Cells["autoid"].Value);
                    if (string.IsNullOrEmpty(autoId))
                    {
                        MessageBox.Show("第" + (i + 1).ToString() + "没有保存，请先保存");
                        return;
                    }
                }



                string cZt = dataGridView1.CurrentRow.Cells["cyanfazt"].Value.ToString();

                if (cZt == "已提交")
                {
                    string sql = @" update zdy_lk_xunjia
                        set cyanfazt='询价完成'
                         where id = @id  ";
                    DbHelper.ExecuteNonQuery(sql, new SqlParameter[]{ 
                                                           new SqlParameter("@id",label6.Text) });

                    string cXs = dataGridView1.CurrentRow.Cells["cmaker"].Value.ToString();
                    string cPersoncode = DbHelper.ExecuteScalar("select cUser_Id  from ufsystem..UA_User where cUser_Name ='" + cXs + "'").ToString();

                    string sqlmsg = @"
  INSERT INTO UFSystem..ua_message (cmsgid,nmsgtype,cmsgtitle,cmsgcontent,csender,creceiver,dsend,nvalidday,bhasread,nurgent, account,[year],cmsgpara)
VALUES(newid(),313552,'询价完成提醒,id号'+@id,'研发询价完成提醒，id号'+@id,@sender,@rec,getdate(),4,0,0,'001','2016',null)";
                    DbHelper.ExecuteNonQuery(sqlmsg, new SqlParameter[]{ 
                                                           new SqlParameter("@id",label6.Text),
                              new SqlParameter("@sender","研发询价"),
                                new SqlParameter("@rec",cPersoncode),
                            });





                }


                MessageBox.Show("提交完成");
                Cx();
            }

            catch (Exception ex)
            {
                CommonHelper.MsgInformation(ex.Message);
                return;
            }

        }
        #endregion

        #region 联查
        public void Liancha()
        {
            try
            {
                this.Validate();
                this.Update();

                string cInvcode = DbHelper.GetDbString(dataGridView1.CurrentRow.Cells["cinvcode"].Value);
                string cInvaddcode = DbHelper.GetDbString(dataGridView1.CurrentRow.Cells["cinvaddcode"].Value);
                string cInvname = DbHelper.GetDbString(dataGridView1.CurrentRow.Cells["cinvname"].Value);
                string cInvstd = DbHelper.GetDbString(dataGridView1.CurrentRow.Cells["cinvstd"].Value);
                if (string.IsNullOrEmpty(cInvcode) == false || string.IsNullOrEmpty(cInvaddcode) == false)
                {

                    FrmTree frm = new FrmTree(cInvcode, cInvname, cInvstd, cInvaddcode);
                    frm.ShowDialog();
                }
                else
                {

                    MessageBox.Show("选中行存货编码为空，不能联查！");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }
        #endregion

        #region 查询
        public void Cx()
        {
            SearchCondition searchObj = new SearchCondition();
            //searchObj.AddCondition("cinvcode", txtcinvcode.Text, SqlOperator.Like);
            searchObj.AddCondition("cmaker", txtcCusname.Text, SqlOperator.Like);
            searchObj.AddCondition("ddate", dateTimePicker1.Value.ToString("yyyy-MM-dd"), SqlOperator.MoreThanOrEqual, dateTimePicker1.Checked == false);
            searchObj.AddCondition("ddate", dateTimePicker2.Value.ToString("yyyy-MM-dd"), SqlOperator.LessThanOrEqual, dateTimePicker2.Checked == false);
            searchObj.AddCondition("cyanfazt", comboBox1.Text, SqlOperator.Equal);


            string conditionSql = searchObj.BuildConditionSql(2);

            if (!string.IsNullOrEmpty(txtcinvcode.Text))
            {

                conditionSql += string.Format(" and (cinvcode like '%{0}%' or cinvaddcode like '%{0}%')", txtcinvcode.Text);
            }

            if (cbxYanfa.Checked)
            {

                conditionSql += " and byanfa = 1";
            }

            StringBuilder strb = new StringBuilder(@"SELECT 0 as xz, [id]
      ,[xmbm]
      ,[cinvcode]
      ,[cinvaddcode]
      ,[cinvname]
      ,[cinvstd]
      ,[cqty1]
      ,[cmemo1]
      ,[czt]
      ,[cmaker]
      ,[ddate]
      ,[bimportant]
      ,[burgent]
      ,[byanfa]
      ,[cyanfazt]
      ,[cEnglishname] from zdy_lk_xunjia where 1=1  ");
            strb.Append(conditionSql);


            dtXunjia = DbHelper.ExecuteTable(strb.ToString());
            dataGridView1.DataSource = dtXunjia;
        }
        #endregion

        #region 输入条件 参照
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {

                U8RefService.IServiceClass obj = new U8RefService.IServiceClass();
                obj.RefID = "Inventory_AA";
                obj.Mode = U8RefService.RefModes.modeRefing;
                //obj.FilterSQL = " bbomsub =1";
                obj.FillText = txtcinvcode.Text;
                obj.Web = false;
                obj.MetaXML = "<Ref><RefSet   bMultiSel='0'  /></Ref>";
                obj.RememberLastRst = false;
                ADODB.Recordset retRstGrid = null, retRstClass = null;
                string sErrMsg = "";
                obj.GetPortalHwnd((int)this.Handle);

                Object objLogin = canshu.u8Login;
                if (obj.ShowRefSecond(ref objLogin, ref retRstClass, ref retRstGrid, ref sErrMsg) == false)
                {
                    MessageBox.Show(sErrMsg);
                }
                else
                {
                    if (retRstGrid != null)
                    {

                        this.txtcinvcode.Text = DbHelper.GetDbString(retRstGrid.Fields["cinvcode"].Value);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("参照失败，原因：" + ex.Message);
            }
        }
        #endregion

        #region 研发询价
        public void YanfaXunjia()
        {
            dataGridView1.Update();
            dataGridView1.EndEdit();

            try
            {

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    //是否选中
                    int ixz = DbHelper.GetDbInt(dataGridView1.Rows[i].Cells["xz"].Value);
                    if (ixz == 1)
                    {
                        //没id，提示保存
                        string id = DbHelper.GetDbString(dataGridView1.Rows[i].Cells["id"].Value);


                        string cZt = DbHelper.GetDbString(dataGridView1.Rows[i].Cells["byanfa"].Value);
                        if (cZt != "True")
                        {
                            string sql = @" update zdy_lk_xunjia
                        set byanfa = 1,cyanfazt='已提交'
                         where id = @id  ";
                            DbHelper.ExecuteNonQuery(sql, new SqlParameter[]{ 
                                                           new SqlParameter("@id",dataGridView1.Rows[i].Cells["id"].Value)
                            });

                            dataGridView1.Rows[i].Cells["cyanfazt"].Value = "已提交";
                        }
                        else
                        {
                            MessageBox.Show("第" + (i + 1).ToString() + "行已提交研发询价");
                            continue;

                        }
                    }


                }


                MessageBox.Show("提交完成");
                Cx();
            }

            catch (Exception ex)
            {
                CommonHelper.MsgInformation(ex.Message);
                return;
            }

        }
        #endregion

        #region
        public void YanfaXunjia2()
        {
            dataGridView1.Update();
            dataGridView1.EndEdit();

            try
            {

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {   //是否选中
                    int ixz = DbHelper.GetDbInt(dataGridView1.Rows[i].Cells["xz"].Value);
                    if (ixz == 1)
                    {
                        //没id，提示保存
                        string id = DbHelper.GetDbString(dataGridView1.Rows[i].Cells["id"].Value);


                        string cZt = DbHelper.GetDbString(dataGridView1.Rows[i].Cells["byanfa"].Value);
                        if (cZt == "True")
                        {
                            string sql = @" update zdy_lk_xunjia
                        set byanfa =0 ,cyanfazt='未提交'
                         where id = @id  ";
                            DbHelper.ExecuteNonQuery(sql, new SqlParameter[]{ 
                                                           new SqlParameter("@id",dataGridView1.Rows[i].Cells["id"].Value)
                            });

                            dataGridView1.Rows[i].Cells["cyanfazt"].Value = "未提交";
                        }
                        else
                        {
                            MessageBox.Show("第" + (i + 1).ToString() + "行未提交研发询价");
                            continue;

                        }

                    }

                }


                MessageBox.Show("撤销完成");
                Cx();
            }

            catch (Exception ex)
            {
                CommonHelper.MsgInformation(ex.Message);
                return;
            }

        }
        #endregion

        #region 进入当前行，获得状态
        private void dataGridView1_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                label6.Text = dataGridView1.Rows[e.RowIndex].Cells["id"].Value.ToString();

                //dtXunjiaPu = DbHelper.ExecuteTable("select id,autoid,ddate,gys,xunjia1,xunjia2,xunjia3,bz1,bz2,bz3,bYanfaQuery,cCost,cattch from zdy_lk_xunjias where gys = '研发' and id ='" + label6.Text + "'");
                dtXunjiaPu = DbHelper.ExecuteTable("select id,autoid,ddate,gys,xunjia1,xunjia2,xunjia3,bz1,bz2,bz3,bYanfaQuery,cCost,cattch from zdy_lk_xunjias where id ='" + label6.Text + "'");
                dataGridView2.DataSource = dtXunjiaPu;


                string cInvcode = DbHelper.GetDbString(dataGridView1.Rows[e.RowIndex].Cells["cinvcode"].Value);
                string cInvaddcode = DbHelper.GetDbString(dataGridView1.Rows[e.RowIndex].Cells["cinvaddcode"].Value);
                string cFtp2;
                if (string.IsNullOrEmpty(cFtp))
                {
                    cFtp2 = "ftp://192.168.0.121/";
                }
                else
                {
                    cFtp2 = cFtp;
                }
                string sql = string.Format("SELECT cas,cpic FROM dbo.zdy_lk_inventory where cas = '{0}' or cas='{1}' ", cInvcode, cInvaddcode);
                DataTable dtcas = DbHelper.ExecuteTable(sql);
                if (dtcas.Rows.Count > 0)
                {
                    string cCas = DbHelper.GetDbString(dtcas.Rows[0]["cas"]);
                    string cLujing = DbHelper.GetDbString(dtcas.Rows[0]["cpic"]);
                    if (string.IsNullOrEmpty(cLujing))
                    {
                        cLujing = string.Format("{0}.png", cCas);

                    }
                    Image img = Image.FromStream(Info(cFtp2, cLujing));
                    pictureBox1.Image = img;
                    Lujing = cLujing;
                }
                else
                {
                    pictureBox1.Image = null;
                }
            }



        }
        #endregion

       

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ftpUrl">FTP地址</param>
        /// <param name="fileName">路径+文件名</param>
        /// <returns></returns>
        public System.IO.Stream Info(string ftpUrl, string fileName)
        {
            try
            {
                string sql = "select cvalue from zdy_lk_para where lx = '102'";
                string cName2 = DbHelper.GetDbString(DbHelper.ExecuteScalar(sql));
                sql = "select cvalue from zdy_lk_para where lx = '103'";
                string cPass = DbHelper.GetDbString(DbHelper.ExecuteScalar(sql));

                FtpWebRequest reqFtp = (FtpWebRequest)FtpWebRequest.Create(new Uri(ftpUrl + "" + fileName));
                reqFtp.Credentials = new NetworkCredential(cName2, cPass);
                reqFtp.UseBinary = true;
                FtpWebResponse respFtp = (FtpWebResponse)reqFtp.GetResponse();
                System.IO.Stream stream = respFtp.GetResponseStream();
                return stream;
            }
            catch (Exception ex)
            {
                
                MessageBox.Show( ex.Message+":"+ftpUrl+fileName);
                throw;
            }
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string pictureName;
            //string houzhui="png";
            //string[] str = Lujing.Split('.');
            //if(str.Length==2)
            //{
            //pictureName = str[0];
            // houzhui = str[1];
            //}
            // 设置保存文件的类型，即文件的扩展名
            saveFileDialog1.Filter = "jpg图片|*.JPG|gif图片|*.GIF|png图片|*.PNG|jpeg图片|*.JPEG";
            saveFileDialog1.FilterIndex = 3;//设置默认文件类型显示顺序 
            // 设置默认的文件名。注意！文件扩展名须与Filter匹配
            saveFileDialog1.FileName = Lujing;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {

                pictureName = saveFileDialog1.FileName;


                if (pictureBox1.Image != null)
                {

                    ////********************照片另存*********************************

                    using (MemoryStream mem = new MemoryStream())
                    {

                        //这句很重要，不然不能正确保存图片或出错（关键就这一句）

                        Bitmap bmp = new Bitmap(pictureBox1.Image);

                        //保存到内存

                        //bmp.Save(mem, pictureBox1.Image.RawFormat );

                        //保存到磁盘文件

                        bmp.Save(@pictureName, pictureBox1.Image.RawFormat);

                        bmp.Dispose();



                        MessageBox.Show("图片另存成功！", "系统提示");

                    }

                    ////********************照片另存*********************************

                }



            }
        }


        //定时刷新
        private void timer1_Tick(object sender, EventArgs e)
        {
            //MessageBox.Show("44");
            Cx();
        }

        private void dataGridView2_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView2.CommitEdit((DataGridViewDataErrorContexts)123);
            dataGridView2.BindingContext[dataGridView2.DataSource].EndCurrentEdit();
            iks++;
        }

        private void dataGridView1_CellParsing(object sender, DataGridViewCellParsingEventArgs e)
        {

        }

        //设置颜色
        private void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                string cImportant = this.dataGridView1.Rows[e.RowIndex].Cells["重要"].Value.ToString();
                string cUrgent = this.dataGridView1.Rows[e.RowIndex].Cells["紧急"].Value.ToString();
                if (cImportant == "True" || cUrgent == "True")
                {
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Red;
                }
            }
        }

        private void dataGridView2_Leave(object sender, EventArgs e)
        {
            dataGridView2.CommitEdit((DataGridViewDataErrorContexts)123);
            dataGridView2.BindingContext[dataGridView2.DataSource].EndCurrentEdit();

            Save();


        }

        #region 双击复制
        private void dataGridView1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1 != null)
            {
                string sv = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                Clipboard.SetData(DataFormats.Text, sv);
            }
        }
        #endregion

        private void dataGridView2_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView2 != null)
            {
                string sv = dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                Clipboard.SetData(DataFormats.Text, sv);
            }
        }









        #region 附件操作
        /// <summary>
        /// 下载附件
        /// </summary>
        /// <param name="fuJian">附件fileid</param>
        /// <param name="FileName">下载的路径</param>
        /// <returns></returns>
        private static string DownloadAtt(string fuJian, string FileName)
        {
            try
            {
                string ls = Environment.CurrentDirectory;
                FileManagerClient client = new FileManagerClient();
                client.FileOperator = "manager";
                client.OperatorPassWord = "manager";
                client.HostUrl = canshu.serverName;
                client.Port = 80;
                client.ProtocolType = "HTTP";
                client.IsWeb = true;
                client.ReadFile(fuJian, FileName);
                return FileName;
                //sel.InlineShapes.AddPicture(FileName);

            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.ToString());
                return "false";
            }
        }
        /// <summary>
        /// 上传附件
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        private static string Upload(string filename)
        {
            try
            {
                //string ls = Environment.CurrentDirectory;
                //string fuJian = "";
                //string FileName = ls + @"\tempcode.bmp";//图片所在路径
                FileManagerClient client = new FileManagerClient();
                client.FileOperator = "manager";
                client.OperatorPassWord = "manager";
                client.HostUrl = canshu.serverName;
                client.Port = 80;
                client.ProtocolType = "HTTP";
                client.IsWeb = true;
                string cFileId = client.AddFile(filename, "test", 60000000, canshu.acc, canshu.acc, canshu.u8Login.CurDate.Year, true);
                //sel.InlineShapes.AddPicture(FileName);
                return cFileId;

            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.ToString());
                return "false";
            }
        }

        /// <summary>
        /// 删除附件
        /// </summary>
        /// <param name="filename"></param>
        private static void DelAtt(string filename)
        {
            try
            {
                //string ls = Environment.CurrentDirectory;
                //string fuJian = "";
                //string FileName = ls + @"\tempcode.bmp";//图片所在路径
                FileManagerClient client = new FileManagerClient();
                client.FileOperator = "manager";
                client.OperatorPassWord = "manager";
                client.HostUrl = canshu.serverName;
                client.Port = 80;
                client.ProtocolType = "HTTP";
                client.IsWeb = true;
                client.DeleteFile(filename);
                //string cFileId = client.AddFile(filename, "test", 60000000, canshu.acc, canshu.acc, canshu.u8Login.CurDate.Year, true);
                ////sel.InlineShapes.AddPicture(FileName);
                //return cFileId;

            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.ToString());

            }
        }

        #endregion


        #region 右键上传附件
        private void 添加附件_Click(object sender, EventArgs e)
        {
            try
            {
                // 判断有没有id，没有id，先进行保存再上传
                int iRow = dataGridView2.CurrentCell.RowIndex;
                string id = DbHelper.GetDbString(dataGridView2.Rows[iRow].Cells["autoid"].Value);
                if (string.IsNullOrEmpty(id))
                {
                    Save();
                    //保存后重新进行获取
                    id = DbHelper.GetDbString(dataGridView2.Rows[iRow].Cells["autoid"].Value);

                }


                string cattch = DbHelper.GetDbString(dataGridView2.Rows[iRow].Cells["cattch"].Value);
                string cattchid = DbHelper.GetDbString(dataGridView2.Rows[iRow].Cells["cattch_fileid"].Value);
                //判断之前有没有上传的图片，有的话，先删除
                //object oRes = DbHelper.ExecuteScalar(string.Format(@"SELECT cattchstr_fileid FROM zdy_lk_projectinquiry WHERE id ={0}", id));

                if (!string.IsNullOrEmpty(cattchid))
                {
                    DialogResult dre = CommonHelper.MsgQuestion("已有附件，是否替换？");
                    if (dre == DialogResult.Yes)
                    {
                        DelAtt(cattchid);
                    }
                    else
                    {
                        return;

                    }
                }

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {

                    string sPic = openFileDialog1.FileName;
                    string filename = System.IO.Path.GetFileName(sPic);//文件名 “Default.aspx”
                    string cId = Upload(sPic);
                    string sql = string.Format(@"update zdy_lk_xunjias set  cattch ='{0}',
cattch_fileid ='{1}' where autoid = '{2}'", filename, cId, id);
                    int i = DbHelper.ExecuteNonQuery(sql);
                    if (i == 1)
                    {
                        MessageBox.Show("附件上传成功！");
                        dataGridView2.Rows[iRow].Cells["cattch"].Value = filename;
                        dataGridView2.Rows[iRow].Cells["cattch_fileid"].Value = cId;

                    }
                }



            }
            catch (Exception ex)
            {
                MessageBox.Show("上传失败：" + ex.Message);
                return;

            }
        }
        #endregion

        #region 删除附件
        private void 删除附件_Click(object sender, EventArgs e)
        {
            int iRow = dataGridView2.CurrentCell.RowIndex;
            string id = DbHelper.GetDbString(dataGridView2.Rows[iRow].Cells["autoid"].Value);
            string cattch = DbHelper.GetDbString(dataGridView2.Rows[iRow].Cells["cattch"].Value);
            string cattchid = DbHelper.GetDbString(dataGridView2.Rows[iRow].Cells["cattch_fileid"].Value);


            if (!string.IsNullOrEmpty(cattchid))
            {

                DelAtt(cattchid);
                string sql = string.Format(@"update zdy_lk_xunjias set  cattch =null,
cattch_fileid =null where autoid = '{0}'", id);
                int i = DbHelper.ExecuteNonQuery(sql);
                if (i == 1)
                {
                    MessageBox.Show("删除成功！");
                    dataGridView2.Rows[iRow].Cells["cattch"].Value = null;
                    dataGridView2.Rows[iRow].Cells["cattch_fileid"].Value = null;

                }
            }
            else
            {
                MessageBox.Show("没有附件，无需删除！");
                return;

            }








        }

        #endregion
        #region 下载附件
        private void 下载附件_Click(object sender, EventArgs e)
        {

            string cFname = DbHelper.GetDbString(dataGridView2.CurrentRow.Cells["cattch"].Value);
            string cFid = DbHelper.GetDbString(dataGridView2.CurrentRow.Cells["cattch_fileid"].Value);


            if (string.IsNullOrEmpty(cFname) == false)
            {
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.FileName = cFname;

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string pictureName = saveFileDialog1.FileName;

                    DownloadAtt(cFid, pictureName);

                    MessageBox.Show("下载完成！", "系统提示");





                }

            }
            else
            {

                MessageBox.Show("选中行没有附件,无法下载！");
            }



        }
        #endregion

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int CIndex = e.ColumnIndex;
            if (dataGridView1.Columns[CIndex].Name == "图片")
            {
                string cInvcode = DbHelper.GetDbString(dataGridView1.Rows[e.RowIndex].Cells["cinvcode"].Value);
                string cInvaddcode = DbHelper.GetDbString(dataGridView1.Rows[e.RowIndex].Cells["cinvaddcode"].Value);

                string sql = string.Format("SELECT cas,cpic FROM dbo.zdy_lk_inventory where cas = '{0}' or cas='{1}' ", cInvcode, cInvaddcode);
                DataTable dtcas = DbHelper.ExecuteTable(sql);
                if (dtcas.Rows.Count > 0)
                {
                    string cCas = DbHelper.GetDbString(dtcas.Rows[0]["cas"]);
                    string cLujing = DbHelper.GetDbString(dtcas.Rows[0]["cpic"]);
                    if (string.IsNullOrEmpty(cLujing))
                    {
                        cLujing = string.Format("{0}.png", cCas);

                    }
                    FrmPic frm = new FrmPic(cLujing, cCas);
                    frm.Show();
                }
                else
                {
                    MessageBox.Show("找不到图片！");

                }

            }

        }

        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            this.dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.Disable;
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            this.dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
        }

     
      

    }
}

