using System;
using System.ComponentModel;
using System.Windows.Forms;

namespace ExcelApplication
{
	partial class Main
	{
		/// <summary>
		///   Required designer variable.
		/// </summary>
		private IContainer components = null;

		/// <summary>
		///   Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}

			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.components         = new System.ComponentModel.Container();
			this.contextMenuStrip1  = new System.Windows.Forms.ContextMenuStrip(this.components);
			this.fToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.tabPage1           = new System.Windows.Forms.TabPage();
			this.button1            = new System.Windows.Forms.Button();
			this.groupBox1          = new System.Windows.Forms.GroupBox();
			this.ModeLoadRadio      = new System.Windows.Forms.RadioButton();
			this.ModeWriteRadio     = new System.Windows.Forms.RadioButton();
			this.LoadTable          = new System.Windows.Forms.Button();
			this.ToExcel            = new System.Windows.Forms.Button();
			this.AppendColl         = new System.Windows.Forms.Button();
			this.SelfCreateTable    = new System.Windows.Forms.DataGridView();
			this.tabControl1        = new System.Windows.Forms.TabControl();
			this.tabPage2           = new System.Windows.Forms.TabPage();
			this.OpenExcelDialog    = new System.Windows.Forms.OpenFileDialog();
			this.contextMenuStrip1.SuspendLayout();
			this.tabPage1.SuspendLayout();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize) (this.SelfCreateTable)).BeginInit();
			this.tabControl1.SuspendLayout();
			this.SuspendLayout();

			// 
			// contextMenuStrip1
			// 
			this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[]
												{
														this.fToolStripMenuItem
												});

			this.contextMenuStrip1.Name = "contextMenuStrip1";
			this.contextMenuStrip1.Size = new System.Drawing.Size(81, 26);

			// 
			// fToolStripMenuItem
			// 
			this.fToolStripMenuItem.Name = "fToolStripMenuItem";
			this.fToolStripMenuItem.Size = new System.Drawing.Size(80, 22);
			this.fToolStripMenuItem.Text = "F";

			// 
			// tabPage1
			// 
			this.tabPage1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.tabPage1.Controls.Add(this.button1);
			this.tabPage1.Controls.Add(this.groupBox1);
			this.tabPage1.Controls.Add(this.LoadTable);
			this.tabPage1.Controls.Add(this.ToExcel);
			this.tabPage1.Controls.Add(this.AppendColl);
			this.tabPage1.Controls.Add(this.SelfCreateTable);
			this.tabPage1.Location                = new System.Drawing.Point(4, 25);
			this.tabPage1.Name                    = "tabPage1";
			this.tabPage1.Padding                 = new System.Windows.Forms.Padding(3);
			this.tabPage1.Size                    = new System.Drawing.Size(899, 441);
			this.tabPage1.TabIndex                = 0;
			this.tabPage1.Text                    = "Лист 1";
			this.tabPage1.UseVisualStyleBackColor = true;

			// 
			// button1
			// 
			this.button1.Location                = new System.Drawing.Point(627, 94);
			this.button1.Margin                  = new System.Windows.Forms.Padding(4);
			this.button1.Name                    = "button1";
			this.button1.Size                    = new System.Drawing.Size(212, 35);
			this.button1.TabIndex                = 7;
			this.button1.Text                    = "Зформувати діаграму";
			this.button1.UseVisualStyleBackColor = true;

			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.ModeLoadRadio);
			this.groupBox1.Controls.Add(this.ModeWriteRadio);
			this.groupBox1.Dock     = System.Windows.Forms.DockStyle.Left;
			this.groupBox1.Location = new System.Drawing.Point(3, 3);
			this.groupBox1.Name     = "groupBox1";
			this.groupBox1.Size     = new System.Drawing.Size(243, 146);
			this.groupBox1.TabIndex = 6;
			this.groupBox1.TabStop  = false;
			this.groupBox1.Text     = "Режим роботи з листом";

			// 
			// ModeLoadRadio
			// 
			this.ModeLoadRadio.AutoSize                = true;
			this.ModeLoadRadio.Location                = new System.Drawing.Point(6, 83);
			this.ModeLoadRadio.Name                    = "ModeLoadRadio";
			this.ModeLoadRadio.Size                    = new System.Drawing.Size(229, 21);
			this.ModeLoadRadio.TabIndex                = 1;
			this.ModeLoadRadio.TabStop                 = true;
			this.ModeLoadRadio.Text                    = "Завантаження зовнішніх даних";
			this.ModeLoadRadio.UseVisualStyleBackColor = true;

			this.ModeLoadRadio.CheckedChanged +=
					new System.EventHandler(this.ModeLoadRadio_CheckedChanged);

			// 
			// ModeWriteRadio
			// 
			this.ModeWriteRadio.AutoSize                = true;
			this.ModeWriteRadio.Location                = new System.Drawing.Point(6, 43);
			this.ModeWriteRadio.Name                    = "ModeWriteRadio";
			this.ModeWriteRadio.Size                    = new System.Drawing.Size(195, 21);
			this.ModeWriteRadio.TabIndex                = 0;
			this.ModeWriteRadio.TabStop                 = true;
			this.ModeWriteRadio.Text                    = "Власноручне заповнення";
			this.ModeWriteRadio.UseVisualStyleBackColor = true;

			this.ModeWriteRadio.CheckedChanged +=
					new System.EventHandler(this.ModeWriteRadio_CheckedChanged);

			// 
			// LoadTable
			// 
			this.LoadTable.Location                =  new System.Drawing.Point(627, 22);
			this.LoadTable.Margin                  =  new System.Windows.Forms.Padding(4);
			this.LoadTable.Name                    =  "LoadTable";
			this.LoadTable.Size                    =  new System.Drawing.Size(212, 35);
			this.LoadTable.TabIndex                =  5;
			this.LoadTable.Text                    =  "Завантажити таблицю";
			this.LoadTable.UseVisualStyleBackColor =  true;
			this.LoadTable.Click                   += new System.EventHandler(this.LoadTable_Click);

			// 
			// ToExcel
			// 
			this.ToExcel.Location                =  new System.Drawing.Point(290, 22);
			this.ToExcel.Margin                  =  new System.Windows.Forms.Padding(4);
			this.ToExcel.Name                    =  "ToExcel";
			this.ToExcel.Size                    =  new System.Drawing.Size(212, 35);
			this.ToExcel.TabIndex                =  4;
			this.ToExcel.Text                    =  "Зберегти таблицю";
			this.ToExcel.UseVisualStyleBackColor =  true;
			this.ToExcel.Click                   += new System.EventHandler(this.ToExcel_Click);

			// 
			// AppendColl
			// 
			this.AppendColl.Location                = new System.Drawing.Point(290, 86);
			this.AppendColl.Margin                  = new System.Windows.Forms.Padding(4);
			this.AppendColl.Name                    = "AppendColl";
			this.AppendColl.Size                    = new System.Drawing.Size(212, 35);
			this.AppendColl.TabIndex                = 3;
			this.AppendColl.Text                    = "Додати стовпець";
			this.AppendColl.UseVisualStyleBackColor = true;

			this.AppendColl.Click +=
					new System.EventHandler(this.AppendColl_Click);

			// 
			// SelfCreateTable
			// 
			this.SelfCreateTable.AllowUserToAddRows    = false;
			this.SelfCreateTable.AllowUserToDeleteRows = false;

			this.SelfCreateTable.ClipboardCopyMode = System
												.Windows.Forms.DataGridViewClipboardCopyMode
												.EnableAlwaysIncludeHeaderText;

			this.SelfCreateTable.ColumnHeadersHeightSizeMode =
					System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;

			this.SelfCreateTable.ContextMenuStrip = this.contextMenuStrip1;
			this.SelfCreateTable.Dock             = System.Windows.Forms.DockStyle.Bottom;
			this.SelfCreateTable.Location         = new System.Drawing.Point(3, 149);
			this.SelfCreateTable.Margin           = new System.Windows.Forms.Padding(4);
			this.SelfCreateTable.Name             = "SelfCreateTable";
			this.SelfCreateTable.ReadOnly         = true;
			this.SelfCreateTable.Size             = new System.Drawing.Size(891, 287);
			this.SelfCreateTable.TabIndex         = 2;

			this.SelfCreateTable.CellContentClick +=
					new System.Windows.Forms.DataGridViewCellEventHandler(this
																			.SelfCreateTable_CellContentClick);

			// 
			// tabControl1
			// 
			this.tabControl1.Controls.Add(this.tabPage1);
			this.tabControl1.Controls.Add(this.tabPage2);
			this.tabControl1.Dock          = System.Windows.Forms.DockStyle.Fill;
			this.tabControl1.Location      = new System.Drawing.Point(0, 0);
			this.tabControl1.Name          = "tabControl1";
			this.tabControl1.SelectedIndex = 0;
			this.tabControl1.Size          = new System.Drawing.Size(907, 470);
			this.tabControl1.TabIndex      = 2;

			this.tabControl1.SelectedIndexChanged +=
					new System.EventHandler(this.WorksheetSelectedIndexChanged);

			// 
			// tabPage2
			// 
			this.tabPage2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;

			this.tabPage2.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F,
														System.Drawing.FontStyle.Bold,
														System.Drawing.GraphicsUnit.Point,
														((byte) (204)));

			this.tabPage2.Location                = new System.Drawing.Point(4, 25);
			this.tabPage2.Name                    = "tabPage2";
			this.tabPage2.Size                    = new System.Drawing.Size(899, 441);
			this.tabPage2.TabIndex                = 1;
			this.tabPage2.Text                    = "+";
			this.tabPage2.UseVisualStyleBackColor = true;

			// 
			// OpenExcelDialog
			// 
			this.OpenExcelDialog.FileOk +=
					new System.ComponentModel.CancelEventHandler(this.OpenExcelDialog_FileOk);

			// 
			// Main
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
			this.AutoScaleMode       = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize          = new System.Drawing.Size(907, 470);
			this.Controls.Add(this.tabControl1);
			this.Font            =  new System.Drawing.Font("Microsoft Sans Serif", 10F);
			this.FormBorderStyle =  System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.Margin          =  new System.Windows.Forms.Padding(4);
			this.MaximizeBox     =  false;
			this.MinimizeBox     =  false;
			this.Name            =  "Main";
			this.StartPosition   =  System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text            =  "Головне вікно";
			this.Load            += new System.EventHandler(this.Main_Load);
			this.contextMenuStrip1.ResumeLayout(false);
			this.tabPage1.ResumeLayout(false);
			this.groupBox1.ResumeLayout(false);
			this.groupBox1.PerformLayout();
			((System.ComponentModel.ISupportInitialize) (this.SelfCreateTable)).EndInit();
			this.tabControl1.ResumeLayout(false);
			this.ResumeLayout(false);
		}

		#endregion

		private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
		private System.Windows.Forms.TabPage tabPage1;
		private System.Windows.Forms.Button ToExcel;
		private System.Windows.Forms.Button AppendColl;
		private System.Windows.Forms.DataGridView SelfCreateTable;
		private System.Windows.Forms.TabControl tabControl1;
		private System.Windows.Forms.TabPage tabPage2;
		private System.Windows.Forms.Button LoadTable;
		private System.Windows.Forms.ToolStripMenuItem fToolStripMenuItem;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.RadioButton ModeLoadRadio;
		private System.Windows.Forms.RadioButton ModeWriteRadio;
		private System.Windows.Forms.OpenFileDialog OpenExcelDialog;
		private System.Windows.Forms.Button button1;
	}
}
