using System;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Speech.Synthesis;
using Microsoft.Win32;
using System.IO;
using System.Diagnostics;
//using TextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;

namespace SpeakOut
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        SpeechSynthesizer minion = new SpeechSynthesizer();
        bool isSpeaking = false; // global declaration
        private void Button_Click_1(object sender ,EventArgs e)
        {
            try
            {
                
                if (!isSpeaking)
                {
                    if (this.box_top.Text.Length == 0)
                    {
                        MessageBox.Show("Please Write Something or select a file");
                    }
                    else
                    {
                        isSpeaking = true;
                        minion.Rate = (int)(this.speed.Value);
                        minion.Volume = (int)(this.volumebtn.Value);
                        ComboBoxItem it = (ComboBoxItem)voice_choice.SelectedItem;
                        switch (it.Content.ToString())
                        {
                            case "Male Voice":
                                this.minion.SelectVoiceByHints(VoiceGender.Male, VoiceAge.Teen);
                                break;
                            case "Female Voice":
                                this.minion.SelectVoiceByHints(VoiceGender.Female, VoiceAge.Teen);
                                break;
                            default:
                                this.minion.SelectVoiceByHints(VoiceGender.Neutral, VoiceAge.Teen);
                                break;
                        }
                        minion.SpeakAsync(this.box_top.Text.ToString());
                        minion.SpeakCompleted += new EventHandler<SpeakCompletedEventArgs>(minion_tired);
                        this.btn.Content = "Stop Speaking";
                        this.pause_btn.IsEnabled = true;
                        this.pause_btn.Content = "Pause";
                    }
                }
                else if(isSpeaking)
                {
                    minion.Dispose();
                    minion = new SpeechSynthesizer();
                    this.btn.Content = "Speak Out";
                    isSpeaking = false;
                    this.pause_btn.IsEnabled = false;
                    this.pause_btn.Content = "Resume";
                }
            }
            catch(Exception err)
            {
                MessageBox.Show(err.ToString());
            }


        }

        public void minion_tired(object sender,SpeakCompletedEventArgs e)
        {
            minion.Dispose();
            minion = new SpeechSynthesizer();
            this.btn.Content = "Speak Out";
            isSpeaking = false;
            this.pause_btn.IsEnabled = false;
            this.pause_btn.Content = "Pause";
        }

     

        private void volumebtn_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!isSpeaking)
            {
                minion.Volume = (int)(this.volumebtn.Value);
            }
        }
        private void speed_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!isSpeaking)
            {
                minion.Rate = (int)(this.speed.Value);
            }
            
        }

        private void pause_btn_Click(object sender, RoutedEventArgs e)
        {
            if (isSpeaking)
            {
                minion.Pause();
                this.pause_btn.Content = "Resume";
                isSpeaking = false;
            }
            else
            {
                minion.Resume();
                this.pause_btn.Content = "Pause";
                isSpeaking = true;
            }
        }

        private void record_btn_Click(object sender, RoutedEventArgs e)
        {
            SpeechSynthesizer rec = new SpeechSynthesizer();
            rec.Rate = (int)(this.speed.Value);
            rec.Volume = (int)(this.volumebtn.Value);
            
            SaveFileDialog pop = new SaveFileDialog();
            pop.Filter = "mp3 files | *.mp3";
            pop.ShowDialog();
            string filename = pop.FileName;
            try
            {
                ComboBoxItem it = (ComboBoxItem)voice_choice.SelectedItem;
                switch (it.Content.ToString())
                {
                    case "Male Voice":
                        rec.SelectVoiceByHints(VoiceGender.Male, VoiceAge.Teen);
                        break;
                    case "Female Voice":
                        rec.SelectVoiceByHints(VoiceGender.Female, VoiceAge.Teen);
                        break;
                    default:
                        rec.SelectVoiceByHints(VoiceGender.Neutral, VoiceAge.Teen);
                        break;
                }
                rec.SetOutputToWaveFile(pop.FileName);
                rec.SpeakAsync(this.box_top.Text);
                rec.SpeakCompleted += new EventHandler<SpeakCompletedEventArgs>(save_complete);
                //rec.SetOutputToDefaultAudioDevice();
                
            }
            catch (System.ArgumentException)
            {
                rec.Dispose();
            }
            
        }

        public void save_complete(object sender, SpeakCompletedEventArgs e)
        {
            MessageBox.Show("recording completed");
        }

        private void fileLoader_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog loader = new OpenFileDialog();
            loader.Filter = "Text File |*.txt";
            loader.ShowDialog();
            try
            {
                StreamReader file = new StreamReader(loader.FileName);
                this.box_top.Text = file.ReadToEnd();
            }
            catch (System.ArgumentException)
            {
                //do nothing
            }

            
        }

        private void pdfLoader_Click(object sender, RoutedEventArgs e)
        {             
            /*Process pr = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            pr.StartInfo = startInfo;*/
            StringBuilder text = new StringBuilder();  
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "PDF Files |*.pdf";
            ofd.ShowDialog();
            using (PdfReader reader = new PdfReader(ofd.FileName.ToString()))
            {
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    text.Append(PdfTextExtractor.GetTextFromPage(reader, i));
                }
            }
            this.box_top.Text = text.ToString();
        }

        private void Reset_Click(object sender, RoutedEventArgs e)
        {
            this.minion.Dispose();
            minion = new SpeechSynthesizer();
            this.box_top.Text = "";
            this.btn.Content = "Speak Out";
            this.pause_btn.Content = "Pause";
            this.pause_btn.IsEnabled = false;
            isSpeaking = false;
        }

       
    }
}
