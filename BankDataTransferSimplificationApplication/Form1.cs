using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace BankDataTransferSimplificationApplication
{
	public class Form1 : Form
	{
		public string fileNameSavePath = string.Format("{0}\\Result_" + DateTime.Now.ToString("dd-MM-yyyy_HH.mm.ss") + ".xls", Directory.GetCurrentDirectory());

		public DataTable dtReadedData;

		public DataTable dtDataForSave;

		private IContainer components = null;

		private OpenFileDialog openFileDialogDosyaSec;

		private SaveFileDialog saveFileDialogDosya;

		private Label label1;

		private DataGridView dataGridView1;

		private Button buttonSave;

		private LinkLabel labelFileName;

		public Form1()
		{
			InitializeComponent();
		}

		private void buttonSadelestir_Click(object sender, EventArgs e)
		{
			buttonSave.Enabled = true;
		}

		private void labelFileName_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
		{
			openFileDialogDosyaSec.FileName = string.Empty;
			openFileDialogDosyaSec.Filter = "Excel Files | *.xls*";
			openFileDialogDosyaSec.FileName = "Template.xls";
			if (openFileDialogDosyaSec.ShowDialog() == DialogResult.OK)
			{
				string sheetName = "Sayfa1";
				labelFileName.Text = openFileDialogDosyaSec.FileName;
				try
				{
					ExcelOku(openFileDialogDosyaSec.FileName, sheetName);
					buttonSave.Enabled = true;
				}
				catch (Exception)
				{
					buttonSave.Enabled = false;
					MessageBox.Show("Error Reading Data!\r\nCheck File Format!", "An error occurred!", MessageBoxButtons.OK, MessageBoxIcon.Hand);
				}
			}
		}

		private void ExcelOku(string fileName, string sheetName)
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			dictionary["Provider"] = "Microsoft.ACE.OLEDB.12.0";
			dictionary["Data Source"] = fileName;
			dictionary["Extended Properties"] = "Excel 12.0";
			StringBuilder stringBuilder = new StringBuilder();
			foreach (KeyValuePair<string, string> item in dictionary)
			{
				stringBuilder.Append(item.Key);
				stringBuilder.Append('=');
				stringBuilder.Append(item.Value);
				stringBuilder.Append(';');
			}
			string connectionString = stringBuilder.ToString();
			using (OleDbConnection oleDbConnection = new OleDbConnection(connectionString))
			{
				try
				{
					oleDbConnection.Open();
					using (OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter("SELECT *  FROM [" + sheetName + "$] ", oleDbConnection))
					{
						dtReadedData = new DataTable(sheetName);
						oleDbDataAdapter.Fill(dtReadedData);
						ArrayList arrayList = new ArrayList();
						for (int i = 0; i < dtReadedData.Rows.Count; i++)
						{
							if (!arrayList.Contains(dtReadedData.Rows[i]["TCKN"].ToString()))
							{
								arrayList.Add(dtReadedData.Rows[i]["TCKN"].ToString());
							}
						}
						dtDataForSave = dtReadedData.Clone();
						for (int i = 0; i < arrayList.Count; i++)
						{
							dtDataForSave.ImportRow(dtReadedData.Select(string.Concat("TCKN='", arrayList[i], "'"))[0]);
							dtDataForSave.Rows[dtDataForSave.Rows.Count - 1]["AMOUNT"] = (double)dtReadedData.Compute("Sum(AMOUNT)", string.Concat("TCKN='", arrayList[i], "'"));
						}
						dataGridView1.AutoGenerateColumns = true;
						dataGridView1.DataSource = dtDataForSave;
					}
				}
				catch (Exception ex)
				{
					throw ex;
				}
			}
		}

		private void DataTableToExcel(DataTable dt, string filePath)
		{
			string str = "";
			string text = "";
			for (int i = 0; i < dt.Columns.Count; i++)
			{
				text = text.ToString() + Convert.ToString(dt.Columns[i].ColumnName) + "\t";
			}
			str = str + text + "\r\n";
			for (int j = 0; j < dt.Rows.Count; j++)
			{
				string text2 = "";
				for (int i = 0; i < dt.Columns.Count; i++)
				{
					text2 = text2.ToString() + Convert.ToString(dt.Rows[j][i].ToString()) + "\t";
				}
				str = str + text2 + "\r\n";
			}
			Encoding encoding = Encoding.GetEncoding(1254);
			byte[] bytes = encoding.GetBytes(str);
			FileStream fileStream = new FileStream(filePath, FileMode.Create);
			BinaryWriter binaryWriter = new BinaryWriter(fileStream);
			binaryWriter.Write(bytes, 0, bytes.Length);
			binaryWriter.Flush();
			binaryWriter.Close();
			fileStream.Close();
		}

		private void buttonSave_Click(object sender, EventArgs e)
		{
			saveFileDialogDosya.Filter = "Excel Files | *.xls*";
			saveFileDialogDosya.FileName = "BankDataTransferSimplificationApplicationResult_" + DateTime.Now.ToString("dd-MM-yyyy_HH.mm.ss") + ".xls";
			if (saveFileDialogDosya.ShowDialog() == DialogResult.OK)
			{
				fileNameSavePath = saveFileDialogDosya.FileName;
				try
				{
					DataTableToExcel(dtDataForSave, fileNameSavePath);
					MessageBox.Show("File Successfully Created.", "Operation Successful", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
				}
				catch (Exception ex)
				{
					throw ex;
				}
			}
		}

		protected override void Dispose(bool disposing)
		{
			if (disposing && components != null)
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		private void InitializeComponent()
		{
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(BankDataTransferSimplificationApplication.Form1));
			openFileDialogDosyaSec = new System.Windows.Forms.OpenFileDialog();
			saveFileDialogDosya = new System.Windows.Forms.SaveFileDialog();
			label1 = new System.Windows.Forms.Label();
			dataGridView1 = new System.Windows.Forms.DataGridView();
			buttonSave = new System.Windows.Forms.Button();
			labelFileName = new System.Windows.Forms.LinkLabel();
			((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
			SuspendLayout();
			label1.AutoSize = true;
			label1.Location = new System.Drawing.Point(12, 21);
			label1.Name = "label1";
			label1.Size = new System.Drawing.Size(81, 13);
			label1.TabIndex = 1;
			label1.Text = "Selected File :";
			dataGridView1.Anchor = (System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right);
			dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			dataGridView1.Location = new System.Drawing.Point(12, 86);
			dataGridView1.Name = "dataGridView1";
			dataGridView1.Size = new System.Drawing.Size(554, 175);
			dataGridView1.TabIndex = 3;
			buttonSave.Anchor = (System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right);
			buttonSave.Enabled = false;
			buttonSave.Location = new System.Drawing.Point(365, 12);
			buttonSave.Name = "buttonSave";
			buttonSave.Size = new System.Drawing.Size(201, 31);
			buttonSave.TabIndex = 4;
			buttonSave.Text = "Create File";
			buttonSave.UseVisualStyleBackColor = true;
			buttonSave.Click += new System.EventHandler(buttonSave_Click);
			labelFileName.AutoSize = true;
			labelFileName.Location = new System.Drawing.Point(99, 21);
			labelFileName.Name = "labelFileName";
			labelFileName.Size = new System.Drawing.Size(80, 13);
			labelFileName.TabIndex = 5;
			labelFileName.TabStop = true;
			labelFileName.Text = "Choose File!";
			labelFileName.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(labelFileName_LinkClicked);
			base.AutoScaleDimensions = new System.Drawing.SizeF(6f, 13f);
			base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			base.ClientSize = new System.Drawing.Size(578, 261);
			base.Controls.Add(labelFileName);
			base.Controls.Add(buttonSave);
			base.Controls.Add(dataGridView1);
			base.Controls.Add(label1);
			base.Icon = (System.Drawing.Icon)resources.GetObject("$this.Icon");
			base.Name = "Form1";
			Text = "Bank Data Transfer Simplification Application";
			base.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
			ResumeLayout(false);
			PerformLayout();
		}
	}
}
