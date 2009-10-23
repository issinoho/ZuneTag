namespace DrunkenBakery.ZuneTag
{
    partial class MDACversions
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MDACversions));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.lvStatus = new System.Windows.Forms.ListView();
            this.imLV = new System.Windows.Forms.ImageList(this.components);
            this.cmdOK = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lvStatus);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(327, 159);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Installed MDAC Versions";
            // 
            // lvStatus
            // 
            this.lvStatus.GridLines = true;
            this.lvStatus.Location = new System.Drawing.Point(6, 19);
            this.lvStatus.Name = "lvStatus";
            this.lvStatus.Size = new System.Drawing.Size(315, 134);
            this.lvStatus.SmallImageList = this.imLV;
            this.lvStatus.TabIndex = 6;
            this.lvStatus.UseCompatibleStateImageBehavior = false;
            this.lvStatus.View = System.Windows.Forms.View.Details;
            // 
            // imLV
            // 
            this.imLV.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imLV.ImageStream")));
            this.imLV.TransparentColor = System.Drawing.Color.Transparent;
            this.imLV.Images.SetKeyName(0, "information.png");
            // 
            // cmdOK
            // 
            this.cmdOK.Location = new System.Drawing.Point(264, 179);
            this.cmdOK.Name = "cmdOK";
            this.cmdOK.Size = new System.Drawing.Size(75, 23);
            this.cmdOK.TabIndex = 1;
            this.cmdOK.Text = "&OK";
            this.cmdOK.UseVisualStyleBackColor = true;
            this.cmdOK.Click += new System.EventHandler(this.cmdOK_Click);
            // 
            // MDACversions
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(351, 210);
            this.Controls.Add(this.cmdOK);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MDACversions";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "MDAC Inspector";
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button cmdOK;
        private System.Windows.Forms.ImageList imLV;
        private System.Windows.Forms.ListView lvStatus;
    }
}