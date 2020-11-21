using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using fuzhu;
using System.IO;
using System.Net;

namespace LKU8.shoukuan
{
    public partial class FrmPic : Form
    {
        string Lujing;
        string Cas;
        public FrmPic()
        {
            InitializeComponent();
        }
        public FrmPic(string cLujing, string cCAS)
        {
            InitializeComponent();

            label1.Text = cCAS;
            Lujing = cLujing;
            Cas = cCAS;

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string sql = "select cvalue from zdy_lk_para where lx = '101'";
            string cFtp = DbHelper.GetDbString(DbHelper.ExecuteScalar(sql));

       

            Image img = Image.FromStream(Info(cFtp, Lujing));
            pictureBox1.Image = img;

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ftpUrl">FTP地址</param>
        /// <param name="fileName">路径+文件名</param>
        /// <returns></returns>
        public Stream Info(string ftpUrl, string fileName)
        {
            try
            {
              string  sql = "select cvalue from zdy_lk_para where lx = '102'";
                string cName2 = DbHelper.GetDbString(DbHelper.ExecuteScalar(sql));
                sql = "select cvalue from zdy_lk_para where lx = '103'";
                string cPass = DbHelper.GetDbString(DbHelper.ExecuteScalar(sql));




              


                FtpWebRequest reqFtp = (FtpWebRequest)FtpWebRequest.Create(new Uri(ftpUrl + "" + fileName));
                reqFtp.Credentials = new NetworkCredential(cName2, cPass);
                reqFtp.UseBinary = true;
                FtpWebResponse respFtp = (FtpWebResponse)reqFtp.GetResponse();
                Stream stream = respFtp.GetResponseStream();
                return stream;
            }
            catch (Exception)
            {
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


    }
}
