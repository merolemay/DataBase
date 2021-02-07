using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;




namespace ImportarExcelToDatagridview
{
    class Importar
    {
        OleDbConnection conn;
        OleDbDataAdapter MyDataAdapter;
        DataTable dt;

        public void importarExcel(DataGridView dgv,String nombreHoja)
        {
            String ruta = "";
            try
            {
                OpenFileDialog openfile1 = new OpenFileDialog();
                openfile1.Filter = "Excel Files |*.xlsx";
                openfile1.Title = "Seleccione el archivo de Excel";
                if (openfile1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    if (openfile1.FileName.Equals("") == false)
                    {
                        ruta = openfile1.FileName;
                    }
                }
                
                    conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;data source=" + ruta + ";Extended Properties='Excel 12.0 Xml;HDR=Yes'");
                    MyDataAdapter = new OleDbDataAdapter("Select * from [" + nombreHoja + "$]", conn);
                    dt = new DataTable();
                    MyDataAdapter.Fill(dt);
                    dgv.DataSource = dt;
                
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void filtrar(TextBox b1, TextBox b2, DataGridView dgv)
        {
            int floor = int.Parse(b1.Text);
            int top = int.Parse(b2.Text);

            List<String> myClass = dgv.DataSource as List<String>;
            List<String> lf = new List<String>();

            for(int i=0;i<myClass.Count-1;i++)
            {
                String divide = myClass[i];
                String[] d = divide.Split(',');
                int dane = int.Parse(d[1]);
                if (dane>floor && dane < top)
                {
                    lf.Add(myClass[i]);
                }
            }


            dgv.DataSource = lf;
        }
    }
}
