using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

namespace Latex4CorelDraw
{    
    public partial class StyleDialog : Form
    {
        private string m_textColor;
        public string TextColor
        {
            get { return m_textColor; }
        }

        public string FontSize 
        {
            get { return comboBoxFontSize.Text; }
        }

        public string LatexFont
        {
            get { return comboBoxFont.Text; }
        }

        public string FontSeries
        {
            get { return comboBoxSeries.Text; }
        }

        public string FontShape
        {
            get { return comboBoxShape.Text; }
        }

        public string MathFont
        {
            get { return comboBoxMathFont.Text; }
        }

        private DialogResult m_result;
        public System.Windows.Forms.DialogResult Result
        {
            get { return m_result; }
        }

        private LatexEquation m_latexEquation;
        public Latex4CorelDraw.LatexEquation LatexEquation
        {
            get { return m_latexEquation; }
            set { m_latexEquation = value; }
        }
        

        private bool m_finishedSuccessfully;

        public StyleDialog()
        {
            InitializeComponent();

            createFontEntries();

            m_finishedSuccessfully = false;
            this.FormClosing += new FormClosingEventHandler(StyleDialog_FormClosing);
        }

        public void init(LatexEquation eq, string title)
        {
            this.Text = title;
            m_finishedSuccessfully = false;
            if (eq != null)
            {
                m_latexEquation = eq;
                if (!comboBoxFontSize.Items.Contains(eq.m_fontSize))
                    comboBoxFontSize.Items.Add(eq.m_fontSize);
                comboBoxFontSize.SelectedItem = eq.m_fontSize;
                comboBoxFont.Text = eq.m_font.fontName;
                comboBoxSeries.Text = eq.m_fontSeries.fontSeries;
                comboBoxShape.Text = eq.m_fontShape.fontShape;
                comboBoxMathFont.Text = eq.m_mathFont.fontName;

                try
                {
                    buttonColor.BackColor = AddinUtilities.stringToColor(eq.m_color);
                    m_textColor = eq.m_color;
                }
                catch
                {
                }
            }
        }

        private void createFontEntries()
        {
            AddinUtilities.initFonts();
            comboBoxFont.Items.AddRange(AddinUtilities.LatexFonts.ToArray());
            comboBoxSeries.Items.AddRange(AddinUtilities.LatexFontSeries.ToArray());
            comboBoxShape.Items.AddRange(AddinUtilities.LatexFontShapes.ToArray());
            comboBoxMathFont.Items.AddRange(AddinUtilities.LatexMathFonts.ToArray());
        }

        void StyleDialog_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            if ((this.DialogResult == DialogResult.OK) && (!m_finishedSuccessfully))
                return;
            m_result = this.DialogResult;

            this.Hide();
        }
      

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void buttonOk_Click(object sender, EventArgs e)
        {
            generateEquation();
        }

        public bool generateEquation()
        {
            // Check paths
            SettingsManager mgr = SettingsManager.getCurrent();

            // Check font size
            string fontSize = comboBoxFontSize.Text;
            float size = 12;
            try
            {
                size = Convert.ToSingle(fontSize);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Font size exception: \n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            // Check Dpi
            float[] systemDPI = AddinUtilities.getSystemDPI();
            float dpiValue = systemDPI[0];
            
            mgr.SettingsData.fontSize = comboBoxFontSize.Text;
            mgr.SettingsData.font = comboBoxFont.Text;
            mgr.SettingsData.fontSeries = comboBoxSeries.Text;
            mgr.SettingsData.fontShape = comboBoxShape.Text;
            mgr.SettingsData.mathFont = comboBoxMathFont.Text;
            mgr.SettingsData.textColor = m_textColor;
            mgr.saveSettings();

            m_latexEquation = new LatexEquation(m_latexEquation.m_code, size, m_textColor, (LatexFont)comboBoxFont.SelectedItem,
                                                      (LatexFontSeries)comboBoxSeries.SelectedItem,
                                                      (LatexFontShape)comboBoxShape.SelectedItem,
                                                      (LatexMathFont)comboBoxMathFont.SelectedItem);

            m_finishedSuccessfully = AddinUtilities.createLatexPdf(m_latexEquation);

            if (m_finishedSuccessfully)
            {
                string imageFile = Path.Combine(AddinUtilities.getAppDataLocation(), "teximport.pdf");
                Corel.Interop.VGCore.StructImportOptions impopt = new Corel.Interop.VGCore.StructImportOptions();
                impopt.MaintainLayers = true;
                Corel.Interop.VGCore.ImportFilter impflt = DockerUI.Current.CorelApp.ActiveLayer.ImportEx(imageFile, Corel.Interop.VGCore.cdrFilter.cdrPDF, impopt);
                impflt.Finish();
                m_latexEquation.m_shape = DockerUI.Current.CorelApp.ActiveShape;
                ShapeTags.setShapeTags(m_latexEquation);
            }

            return m_finishedSuccessfully;
        }

        private void changeOptionsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddinUtilities.changeOptions();
        }

        private void buttonColor_Click(object sender, EventArgs e)
        {
            ColorDialog dialog = new ColorDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                System.Globalization.CultureInfo culture = new System.Globalization.CultureInfo("en-US");
                buttonColor.BackColor = dialog.Color;
                Color col = buttonColor.BackColor;
                float r = (float)col.R / 255.0f;
                float g = (float)col.G / 255.0f;
                float b = (float)col.B / 255.0f;
                string rStr = r.ToString(culture);
                string gStr = g.ToString(culture);
                string bStr = b.ToString(culture);
                m_textColor = rStr + "," + gStr + "," + bStr;

            }
        }

        private void openLatexTemplateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string templateFileName = AddinUtilities.getAppDataLocation() + "\\LatexTemplate.txt";
            System.Diagnostics.Process.Start(templateFileName);
        }

    }


}
