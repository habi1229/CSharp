namespace StockCrawling
{
    partial class CrawlingForm
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnReadExcelData = new System.Windows.Forms.Button();
            this.btnCrawlingKOSPI = new System.Windows.Forms.Button();
            this.btnReadInvestOpinion = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnReadExcelData
            // 
            this.btnReadExcelData.Location = new System.Drawing.Point(12, 72);
            this.btnReadExcelData.Name = "btnReadExcelData";
            this.btnReadExcelData.Size = new System.Drawing.Size(134, 49);
            this.btnReadExcelData.TabIndex = 0;
            this.btnReadExcelData.Text = "ReadExcelData";
            this.btnReadExcelData.UseVisualStyleBackColor = true;
            this.btnReadExcelData.Click += new System.EventHandler(this.btnReadExcelData_Click);
            // 
            // btnCrawlingKOSPI
            // 
            this.btnCrawlingKOSPI.Location = new System.Drawing.Point(12, 12);
            this.btnCrawlingKOSPI.Name = "btnCrawlingKOSPI";
            this.btnCrawlingKOSPI.Size = new System.Drawing.Size(134, 45);
            this.btnCrawlingKOSPI.TabIndex = 1;
            this.btnCrawlingKOSPI.Text = "CrawlingKOSPI";
            this.btnCrawlingKOSPI.UseVisualStyleBackColor = true;
            this.btnCrawlingKOSPI.Click += new System.EventHandler(this.btnCrawlingKOSPI_Click);
            // 
            // btnReadInvestOpinion
            // 
            this.btnReadInvestOpinion.Location = new System.Drawing.Point(12, 138);
            this.btnReadInvestOpinion.Name = "btnReadInvestOpinion";
            this.btnReadInvestOpinion.Size = new System.Drawing.Size(134, 49);
            this.btnReadInvestOpinion.TabIndex = 2;
            this.btnReadInvestOpinion.Text = "ReadInvestOpinion";
            this.btnReadInvestOpinion.UseVisualStyleBackColor = true;
            this.btnReadInvestOpinion.Click += new System.EventHandler(this.btnReadInvestOpinion_Click);
            // 
            // CrawlingForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Controls.Add(this.btnReadInvestOpinion);
            this.Controls.Add(this.btnCrawlingKOSPI);
            this.Controls.Add(this.btnReadExcelData);
            this.Name = "CrawlingForm";
            this.Text = "StockCrawling";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.CrawlingForm_FormClosing);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnReadExcelData;
        private System.Windows.Forms.Button btnCrawlingKOSPI;
        private System.Windows.Forms.Button btnReadInvestOpinion;
    }
}

