using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;

 
namespace armiya
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            
        }

        string folderInfo="";
        private void button1_Click(object sender, EventArgs e)
        {

          
        }
        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }
         //con = new SqlConnection("server=.; Initial Catalog=deneme;User ID=sa;Password=25802580");
        public List<String> liste;
        public void tablogetir(int a) {
            string databasename = textBox1.Text;
            string username = textBox3.Text;
            string password = textBox4.Text;
            string connetionString = null;
            SqlConnection connection;
            SqlCommand command;
            SqlDataAdapter adapter = new SqlDataAdapter();
            DataSet ds = new DataSet();
            int i = 0;
            string sql = null;

            connetionString = "server=.; Initial Catalog='"+databasename+"';User ID='"+username+"';Password='"+password+"'";
            sql = "Select DISTINCT(name) FROM sys.Tables";

            connection = new SqlConnection(connetionString);

            try
            {
                connection.Open();
                command = new SqlCommand(sql, connection);
                adapter.SelectCommand = command;
                adapter.Fill(ds);
                adapter.Dispose();
                command.Dispose();
                
                connection.Close();
                try
                {
                    for (i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
                    {
                        listBox1.Items.Add(ds.Tables[0].Rows[i].ItemArray[0].ToString());

                    }
                    if (a == 1) { excel(); } else if (a == 3) { html(); }
                }
                catch (Exception ex)
                {
                    
                }
                
            }
            catch
            {
                MessageBox.Show("Girdiğiniz bilgileri kontrol ediniz veritabanına bağlanılamadı. ");
                
            }
           
            
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox3.Text == "" || textBox4.Text == "")
            {


                MessageBox.Show("Tüm alanları eksiksiz doldurunuz");
            }
            else { tablogetir(1); }
           
        }

        private void label2_Click_1(object sender, EventArgs e)
        {

        }
        string path;
        string veri;
        public void excel()
        {
           
            string conString = "server=.; Initial Catalog='" + textBox1.Text + "';User ID='" + textBox3.Text + "';Password='" + textBox4.Text + "'";
            SqlConnection baglanti = new SqlConnection(conString);
            baglanti.Open();
            SqlCommand komut;
            SqlDataAdapter da;
            DataTable dt;
            
            string exc="excel";
            Directory.CreateDirectory("" + folderInfo + "\\" + textBox2.Text + "");
            path = folderInfo + "\\" + textBox2.Text;
            Directory.CreateDirectory("" + path + "\\" + exc + "");
            button3.Enabled = false;
            string dos;
            dos = path + "\\" + exc;
                // creating Excel Application  
                Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                // creating new WorkBook within Excel application  
                Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
                // creating new Excelsheet in workbook  
                Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
                // see the excel sheet behind the program  
                app.Visible = false;
                // get the reference of first sheet. By default its name is Sheet1.  
                // store its reference to worksheet  

                string ad = "a";
                for (int x = 0; x < listBox1.Items.Count; x++)
                {
                    veri = listBox1.Items[x].ToString();
                    string kayit = "exec sp_columns " + veri + "";
                    komut = new SqlCommand(kayit, baglanti);


                    da = new SqlDataAdapter(komut);
                    dt = new DataTable();
                    da.Fill(dt);
                    dataGridView1.DataSource = dt;
                    List<string> tut = new List<string>();


                    if (veri.Length > 31)
                    {
                        int pos = veri.LastIndexOf("", 25);
                        veri = veri.Substring(0, pos);
                        tut.Add(veri);

                    }
                    else { veri = veri.Substring(0, veri.Length - 4);
                    tut.Add(veri);
                    }

                    progressBar1.Minimum = 0;
                    progressBar1.Maximum = listBox1.Items.Count;
                    progressBar1.Value = 0;
                    progressBar1.Value += x;  

                    

                    Random rastgele = new Random();
                    int sayi = rastgele.Next(1000);
                    if (tut.Contains(veri))
                    {
                        ad = veri + sayi.ToString();

                    }
                    else
                    {
                        ad = veri + sayi.ToString();
                    }
                   


                    workbook.Sheets.Add();
                    worksheet = workbook.ActiveSheet;
                    // changing the name of active sheet  
                    worksheet.Name = ad;
                    // storing header part in Excel 
                  
                    for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                    {
                        worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                    }
                    // storing Each row and column value to excel sheet  
                    for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                    {
                        try
                        {
                            for (int k = 0; k < dataGridView1.Columns.Count; k++)
                            {
                                worksheet.Cells[i + 2, k + 1] = dataGridView1.Rows[i].Cells[k].Value.ToString();
                            }
                        }catch (ArgumentException e)
                         {
   
                            }
                        
                    }
                    
                }
                
                // save the application  
                workbook.SaveAs(""+dos+"\\output.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                // Exit from the application  
                //app.Quit();  
                MessageBox.Show("Excel'e aktarım işlemi bitmiştir.");
                button3.Enabled = true;
                progressBar1.Value = 0;
            baglanti.Close();
        }   

        private void timer1_Tick(object sender, EventArgs e)
        {
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            timer1.Start();
        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox3.Text == "" || textBox4.Text == "")
            {


                MessageBox.Show("Tüm alanları eksiksiz doldurunuz");
            }
            else { tablogetir(2); }
           
        }
        string dos;
        public void html() {

            string conString = "server=.; Initial Catalog='" + textBox1.Text + "';User ID='" + textBox3.Text + "';Password='" + textBox4.Text + "'";
            SqlConnection baglanti = new SqlConnection(conString);
            baglanti.Open();
            SqlCommand komut;
            SqlDataAdapter da;
            DataTable dt;
            button2.Enabled = false;
            string veri;
            string htm="html";
            Directory.CreateDirectory("" + folderInfo + "\\" + textBox2.Text + "");
            path = folderInfo + "\\" + textBox2.Text;
            Directory.CreateDirectory("" + path + "\\" + htm + "");
            
            dos = path + "\\" + htm;
            for (int x = 0; x < listBox1.Items.Count; x++)
            {
                veri = listBox1.Items[x].ToString();
                string kayit = "exec sp_columns " + veri + "";
                komut = new SqlCommand(kayit, baglanti);


                da = new SqlDataAdapter(komut);
                dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                StringBuilder strHTMLBuilder = new StringBuilder();
                strHTMLBuilder.Append("<html >");
                strHTMLBuilder.Append("<head>");
                strHTMLBuilder.Append("</head>");
                strHTMLBuilder.Append("<body>");
                strHTMLBuilder.Append("<table border='1px' cellpadding='1' cellspacing='1' bgcolor='lightyellow' style='font-family:Garamond; font-size:smaller'>");
                strHTMLBuilder.Append("<tr >");
                foreach (DataColumn myColumn in dt.Columns)
                {
                    strHTMLBuilder.Append("<td >");
                    strHTMLBuilder.Append(myColumn.ColumnName);
                    strHTMLBuilder.Append("</td>");
                }
                strHTMLBuilder.Append("</tr>");
                foreach (DataRow myRow in dt.Rows)
                {
                    strHTMLBuilder.Append("<tr >");
                    foreach (DataColumn myColumn in dt.Columns)
                    {
                        strHTMLBuilder.Append("<td >");
                        strHTMLBuilder.Append(myRow[myColumn.ColumnName].ToString());
                        strHTMLBuilder.Append("</td>");
                    }
                    strHTMLBuilder.Append("</tr>");
                }
                progressBar1.Minimum = 0;
                progressBar1.Maximum = listBox1.Items.Count;
                progressBar1.Value = 0;
                progressBar1.Value += x;      
                //Close tags.ş
                strHTMLBuilder.Append("</table>");
                strHTMLBuilder.Append("</body>");
                strHTMLBuilder.Append("</html>");
                string Htmltext = strHTMLBuilder.ToString();
                System.IO.File.WriteAllText(""+dos+"\\"+veri+".html", strHTMLBuilder.ToString());
                
        
            }
            tumhtml();
            MessageBox.Show("Html'e aktarım işlemi bitmiştir.");
            button2.Enabled = true;
            progressBar1.Value = 0;
        }      

        private void button3_Click(object sender, EventArgs e)
        {
            tablogetir(3);
        }
      
        private void button1_Click_2(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox3.Text == "" || textBox4.Text == "")
            {


                MessageBox.Show("Tüm alanları eksiksiz doldurunuz");
            }
            else
            {
                folderBrowserDialog1.ShowDialog();

                folderInfo = folderBrowserDialog1.SelectedPath;
                
                Directory.CreateDirectory(folderInfo);
                
                
            }
        }

        private void timer1_Tick_1(object sender, EventArgs e)
        {
            if (folderInfo == "")
            {
                button2.Enabled = false;
                button3.Enabled = false;
            }
            else
            {
                button2.Enabled = true;
                button3.Enabled = true;
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            { textBox2.Enabled = true; }
            else { textBox2.Enabled = false; }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            textBox2.Text = textBox1.Text;
        }

        public void tumhtml() {
            string da = "s";
         string renk= "color:black";
            string tb = "tb";
            string a = "4";
            string[] array1 = Directory.GetFiles(dos);
            StringBuilder strHTMLBuilder = new StringBuilder();
            strHTMLBuilder.Append("<html >");
            strHTMLBuilder.Append("<head>");
            strHTMLBuilder.Append("</head>");
            strHTMLBuilder.Append("<body>");

            strHTMLBuilder.Append("<style> .tb { border-collapse: collapse; }.tb th, .tb td { padding: 5px; border: solid 1px #777; }.tb th { background-color: lightblue; }");
            strHTMLBuilder.Append("</style>");

            strHTMLBuilder.Append("<table class="+tb+">");
            strHTMLBuilder.Append("<tr>");
            strHTMLBuilder.Append("<tr >");
            strHTMLBuilder.Append("<th colspan="+a+">TABLO SAYISI :"+listBox1.Items.Count.ToString()+"</th>");  
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("<td>");
            int say = listBox1.Items.Count / 4;
            for (int i = 0; i < listBox1.Items.Count/4; i++)
            {
                
                veri = listBox1.Items[i].ToString();
                da = "html/"+veri + ".html";
                strHTMLBuilder.Append("<tr >");
                strHTMLBuilder.Append("<td >");
                strHTMLBuilder.Append("<a href=" + da + " style=" + renk + ">");
                strHTMLBuilder.Append(veri.ToString());
                strHTMLBuilder.Append("</a>");
                strHTMLBuilder.Append("</td>");
                veri = listBox1.Items[say+i].ToString();
                da = "html/" + veri + ".html";
                strHTMLBuilder.Append("<td >");
                strHTMLBuilder.Append("<a href=" + da + " style=" + renk + ">");
                strHTMLBuilder.Append(veri.ToString());
                strHTMLBuilder.Append("</a>");
                strHTMLBuilder.Append("</td>");

                veri = listBox1.Items[say + i+say].ToString();
                da = "html/" + veri + ".html";
                strHTMLBuilder.Append("<td >");
                strHTMLBuilder.Append("<a href=" + da + " style=" + renk + ">");
                strHTMLBuilder.Append(veri.ToString());
                strHTMLBuilder.Append("</a>");
                strHTMLBuilder.Append("</td>");

                veri = listBox1.Items[say + i + say+say].ToString();
                da = "html/" + veri + ".html";
                strHTMLBuilder.Append("<td >");
                strHTMLBuilder.Append("<a href=" + da + " style=" + renk + ">");
                strHTMLBuilder.Append(veri.ToString());
                strHTMLBuilder.Append("</a>");
                strHTMLBuilder.Append("</td>");

                strHTMLBuilder.Append("</tr>");
               
                
            }

            strHTMLBuilder.Append("</td>");
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("</table>");
            strHTMLBuilder.Append("</body>");
            strHTMLBuilder.Append("</html>");
            string Htmltext = strHTMLBuilder.ToString();
            System.IO.File.WriteAllText("" + path + "\\hepsinigor.html", strHTMLBuilder.ToString());
        
        }

        private void button4_Click(object sender, EventArgs e)
        {
           
        }


    }
}
