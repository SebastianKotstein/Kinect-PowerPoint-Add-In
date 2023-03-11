
using Microsoft.Office.Tools.Ribbon;
using System;

namespace skotstein.app.kinect.powerpoint_add_in
{
    partial class comboBoxVoices : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public comboBoxVoices()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.startStopGroup = this.Factory.CreateRibbonGroup();
            this.buttonStart = this.Factory.CreateRibbonButton();
            this.buttonStop = this.Factory.CreateRibbonButton();
            this.recognitionGroup = this.Factory.CreateRibbonGroup();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.comboBoxLanguage = this.Factory.CreateRibbonComboBox();
            this.voiceGroup = this.Factory.CreateRibbonGroup();
            this.label2 = this.Factory.CreateRibbonLabel();
            this.comboBoxVoice = this.Factory.CreateRibbonComboBox();
            this.tab1.SuspendLayout();
            this.startStopGroup.SuspendLayout();
            this.recognitionGroup.SuspendLayout();
            this.voiceGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.startStopGroup);
            this.tab1.Groups.Add(this.recognitionGroup);
            this.tab1.Groups.Add(this.voiceGroup);
            this.tab1.Label = "Kinect Gesture Control";
            this.tab1.Name = "tab1";
            // 
            // startStopGroup
            // 
            this.startStopGroup.Items.Add(this.buttonStart);
            this.startStopGroup.Items.Add(this.buttonStop);
            this.startStopGroup.Label = "Gesture Recognition";
            this.startStopGroup.Name = "startStopGroup";
            // 
            // buttonStart
            // 
            this.buttonStart.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonStart.Image = global::skotstein.app.kinect.powerpoint_add_in.Properties.Resources.Kinect;
            this.buttonStart.Label = "Start";
            this.buttonStart.Name = "buttonStart";
            this.buttonStart.ShowImage = true;
            this.buttonStart.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonStart_Click);
            // 
            // buttonStop
            // 
            this.buttonStop.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonStop.Image = global::skotstein.app.kinect.powerpoint_add_in.Properties.Resources.Stop;
            this.buttonStop.Label = "Stop";
            this.buttonStop.Name = "buttonStop";
            this.buttonStop.ShowImage = true;
            this.buttonStop.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonStop_Click);
            // 
            // recognitionGroup
            // 
            this.recognitionGroup.Items.Add(this.label1);
            this.recognitionGroup.Items.Add(this.comboBoxLanguage);
            this.recognitionGroup.Label = "Speech Recognition";
            this.recognitionGroup.Name = "recognitionGroup";
            // 
            // label1
            // 
            this.label1.Label = "Input Language:";
            this.label1.Name = "label1";
            // 
            // comboBoxLanguage
            // 
            this.comboBoxLanguage.Label = "Language";
            this.comboBoxLanguage.Name = "comboBoxLanguage";
            this.comboBoxLanguage.ShowLabel = false;
            this.comboBoxLanguage.Text = null;
            // 
            // voiceGroup
            // 
            this.voiceGroup.Items.Add(this.label2);
            this.voiceGroup.Items.Add(this.comboBoxVoice);
            this.voiceGroup.Label = "Voice";
            this.voiceGroup.Name = "voiceGroup";
            // 
            // label2
            // 
            this.label2.Label = "Output Voice:";
            this.label2.Name = "label2";
            // 
            // comboBoxVoice
            // 
            this.comboBoxVoice.Label = "comboBoxVoices";
            this.comboBoxVoice.Name = "comboBoxVoice";
            this.comboBoxVoice.ShowLabel = false;
            this.comboBoxVoice.Text = null;
            // 
            // comboBoxVoices
            // 
            this.Name = "comboBoxVoices";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.KinectRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.startStopGroup.ResumeLayout(false);
            this.startStopGroup.PerformLayout();
            this.recognitionGroup.ResumeLayout(false);
            this.recognitionGroup.PerformLayout();
            this.voiceGroup.ResumeLayout(false);
            this.voiceGroup.PerformLayout();
            this.ResumeLayout(false);

        }



        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup startStopGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonStart;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonStop;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup recognitionGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox comboBoxLanguage;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup voiceGroup;
        internal RibbonLabel label1;
        internal RibbonLabel label2;
        internal RibbonComboBox comboBoxVoice;
    }

    partial class ThisRibbonCollection
    {
        internal comboBoxVoices KinectRibbon
        {
            get { return this.GetRibbon<comboBoxVoices>(); }
        }
    }
}
