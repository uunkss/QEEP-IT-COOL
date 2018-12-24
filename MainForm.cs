using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Data.OleDb;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Reflection;
namespace Qeep_it_cool
{
	/// <summary>
	/// Description of MainForm.
	/// </summary>
	public partial class MainForm : Form
	{
		public MainForm()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
		}
		void TextBox4KeyPress(object sender, KeyPressEventArgs e)
		{
	char ch = e.KeyChar;
			if (!Char.IsDigit(ch) && ch !=8)
			{
				e.Handled = true;
			}
		}
		void TextBox3KeyPress(object sender, KeyPressEventArgs e)
		{
	char ch = e.KeyChar;
			if (!Char.IsDigit(ch) && ch !=8)
			{
				e.Handled = true;
			}
		}
		void Button1Click(object sender, EventArgs e)
		{
			dataGridView1.ColumnCount = 5;
			dataGridView1.Columns[0].Name = "NAME";
			dataGridView1.Columns[1].Name = "ORDER";
			dataGridView1.Columns[2].Name = "PRICE";
			dataGridView1.Columns[3].Name = "LOCATE";
			dataGridView1.Columns[4].Name = "DATE";
			
			dataGridView1.Columns[0].Width = 75;
			dataGridView1.Columns[1].Width = 299;
			dataGridView1.Columns[2].Width = 58;
			dataGridView1.Columns[3].Width = 100;
			dataGridView1.Columns[4].Width = 68;
			
	int n = dataGridView1.Rows.Add();
			dataGridView1.Rows[n].Cells[0].Value = textBox1.Text;
			dataGridView1.Rows[n].Cells[1].Value = textBox2.Text;
			dataGridView1.Rows[n].Cells[2].Value = textBox3.Text;
			dataGridView1.Rows[n].Cells[3].Value = comboBox1.Text;
			dataGridView1.Rows[n].Cells[4].Value = comboBox2.Text + comboBox3.Text + textBox4.Text;
			//textbox clear code
			textBox1.Text = "";
			textBox2.Text = "";
			textBox3.Text = "";
			textBox4.Text = "";
		}
		void Button2Click(object sender, EventArgs e)
		{
	int rowIndex =  dataGridView1.CurrentCell.RowIndex;
			dataGridView1.Rows.RemoveAt(rowIndex);
		}
		void Button3Click(object sender, EventArgs e)
		{
	Microsoft.Office.Interop.Excel.Application xla = new Microsoft.Office.Interop.Excel.ApplicationClass();
			Workbook wb = xla.Workbooks.Add(XlSheetType.xlWorksheet);
			Worksheet ws = (Worksheet)xla.ActiveSheet;
			
			xla.Visible = true;
			
			ws.Cells[1,1] = "NAME";
			ws.Cells[1,2] = "ORDER";
			ws.Cells[1,3] = "PRICE";
			ws.Cells[1,4] = "LOCATE";
			ws.Cells[1,5] = "DATE";
			
			
				for (int i = 0; i < dataGridView1.Rows.Count; i++)
				{
					for (int j = 0; j < dataGridView1.Columns.Count; j++)
					{
						ws.Cells[i+2,j+1]=dataGridView1.Rows[i].Cells[j].Value;
					}
				}
				}
		
		void Button4Click(object sender, EventArgs e)
		{
	OpenFileDialog openFileDialog1 = new OpenFileDialog();
			if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
			{
				this.textBox5.Text = openFileDialog1.FileName;
			}
		}
		void Button5Click(object sender, EventArgs e)
		{
	string PathConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + textBox5.Text + ";Extended Properties=Excel 12.0;";
            OleDbConnection conn = new OleDbConnection(PathConn);

            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter("SELECT * FROM ["+textBox6.Text+"$]", conn);
            System.Data.DataTable dt = new System.Data.DataTable();
            myDataAdapter.Fill(dt);

            dataGridView1.DataSource = dt;
		}
	}
}
