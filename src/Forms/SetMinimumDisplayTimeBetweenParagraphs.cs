﻿using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Nikse.SubtitleEdit.Logic;

namespace Nikse.SubtitleEdit.Forms
{
    public partial class SetMinimumDisplayTimeBetweenParagraphs : Form
    {

        Subtitle _subtitle;
        private Subtitle _fixedSubtitle;
        public int FixCount { get; private set; }

        public SetMinimumDisplayTimeBetweenParagraphs()
        {
            InitializeComponent();

            Text = Configuration.Settings.Language.SetMinimumDisplayTimeBetweenParagraphs.Title;
            labelMaxMillisecondsBetweenLines.Text = Configuration.Settings.Language.SetMinimumDisplayTimeBetweenParagraphs.MinimumMillisecondsBetweenParagraphs;
            checkBoxShowOnlyChangedLines.Text = Configuration.Settings.Language.SetMinimumDisplayTimeBetweenParagraphs.ShowOnlyModifiedLines;
            buttonOK.Text = Configuration.Settings.Language.General.OK;
            buttonCancel.Text = Configuration.Settings.Language.General.Cancel;
            SubtitleListview1.InitializeLanguage(Configuration.Settings.Language.General, Configuration.Settings);
        }

        public Subtitle FixedSubtitle
        {
            get { return _fixedSubtitle; }
            private set { _fixedSubtitle = value; }
        }

        public void Initialize(Subtitle subtitle)
        {
            _subtitle = subtitle;

            GeneratePreview();
        }

        private void GeneratePreview()
        {
            List<int> fixes = new List<int>();
            if (_subtitle == null)
                return;

            FixedSubtitle = new Subtitle(_subtitle);
            var onlyFixedSubtitle = new Subtitle();

            double minumumMillisecondsBetweenLines = (double)numericUpDownMinMillisecondsBetweenLines.Value;
            for (int i = 0; i < FixedSubtitle.Paragraphs.Count - 2; i++)
            {
                Paragraph p = FixedSubtitle.GetParagraphOrDefault(i);
                Paragraph next = FixedSubtitle.GetParagraphOrDefault(i + 1);
                if (next.StartTime.TotalMilliseconds - p.EndTime.TotalMilliseconds < minumumMillisecondsBetweenLines)
                {
                    p.EndTime.TotalMilliseconds = next.StartTime.TotalMilliseconds - minumumMillisecondsBetweenLines;
                    fixes.Add(i);
                    onlyFixedSubtitle.Paragraphs.Add(new Paragraph(p));
                }
            }

            SubtitleListview1.BeginUpdate();
            groupBoxLinesFound.Text = string.Format(Configuration.Settings.Language.SetMinimumDisplayTimeBetweenParagraphs.PreviewLinesModifiedX, fixes.Count);
            if (checkBoxShowOnlyChangedLines.Checked)
            {
                SubtitleListview1.Fill(onlyFixedSubtitle);
            }
            else
            {
                SubtitleListview1.Fill(FixedSubtitle);
                foreach (int index in fixes)
                    SubtitleListview1.SetBackgroundColor(index, Color.Silver);
            }
            SubtitleListview1.EndUpdate();
            FixCount = fixes.Count;
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void SetMinimalDisplayTimeDifference_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                DialogResult = DialogResult.Cancel;
                e.SuppressKeyPress = true;
            }
        }

        private void numericUpDownMinMillisecondsBetweenLines_ValueChanged(object sender, EventArgs e)
        {
            GeneratePreview();
        }

        private void checkBoxShowOnlyChangedLines_CheckedChanged(object sender, EventArgs e)
        {
            GeneratePreview();
        }
    }

}
