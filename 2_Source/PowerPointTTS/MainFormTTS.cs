using SpeechLib;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PowerPointTTS
{
    public partial class MainFormTTS : Form
    {
        private SpVoice spVoiceTts = null;
        private SpeechVoiceSpeakFlags flg = SpeechVoiceSpeakFlags.SVSFlagsAsync | SpeechVoiceSpeakFlags.SVSFPurgeBeforeSpeak;

        public MainFormTTS()
        {
            InitializeComponent();
        }

        private void MainFormTTS_Load(object sender, EventArgs e)
        {
            if (null == this.spVoiceTts)
            {
                this.spVoiceTts = new SpVoice();      // 音声合成のオブジェクト
                ISpeechObjectTokens VoiceInfo;
                VoiceInfo = this.spVoiceTts.GetVoices("", "");
                for (int i = 0; i < VoiceInfo.Count; i++)
                {
                    this.listBox2Language.Items.Add(VoiceInfo.Item(i).GetAttribute("Name"));
                }
                this.listBox2Language.SelectedIndex = 0;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            if (null != this.spVoiceTts)
            {
                this.spVoiceTts.Pause();
            }

            this.spVoiceTts = new SpVoice();      // 音声合成のオブジェクト

            this.spVoiceTts.Voice = this.spVoiceTts.GetVoices("", "").Item(this.listBox2Language.SelectedIndex);
            this.spVoiceTts.Volume = (int)numericUpDown1Vol.Value;
            this.spVoiceTts.Rate = (int)numericUpDown2Speed.Value;


            this.spVoiceTts.Speak(this.textBox1.Text, flg);                   // 文字列の同期読み上げ

        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (null != this.spVoiceTts)
            {
                this.spVoiceTts.Pause();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (this.textBox1.Text == "")
            {
                MessageBox.Show("内容がないため、録音作業を中止しました。");
                return;
            }
            bool isSucess = true;

            
            string nowtimestr = DateTime.Now.ToFileTime().ToString();
            string fileName = nowtimestr;

            string appPath = Directory.GetCurrentDirectory();
            string strtemp = Globals.ThisAddIn.Application.ActivePresentation.Name.Replace(".pptx", "");
            appPath = appPath + @"\" + strtemp + @"\Recode\";


            if (Directory.Exists(appPath) == false)//フォルダがなければ、フォルダを新規する
            {
                Directory.CreateDirectory(appPath);
            }

            if (null != this.spVoiceTts)
            {
                this.spVoiceTts.Pause();
            }
            SpVoice saveVoiceTts = new SpVoice();      // 音声合成のオブジェクト
            saveVoiceTts.Voice = saveVoiceTts.GetVoices("", "").Item(this.listBox2Language.SelectedIndex);
            saveVoiceTts.Volume = (int)numericUpDown1Vol.Value;
            saveVoiceTts.Rate = (int)numericUpDown2Speed.Value;

            string strFullFileName = appPath + @"\" + fileName + @".wav";


            SpFileStream fs = null;
            try
            {
                fs = new SpFileStream();
                fs.Open(strFullFileName, SpeechStreamFileMode.SSFMCreateForWrite, false);

                saveVoiceTts.AudioOutputStream = fs;

                saveVoiceTts.Speak(this.textBox1.Text, SpeechVoiceSpeakFlags.SVSFDefault);
            }
            catch(Exception ex)
            {
                isSucess = false;
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                }
            }

            if (isSucess == false)
            {
                MessageBox.Show("作成失敗しました。音声設定の設定を見直してから、もう一回試して見てください。");
                return;
            }


            //Powerpoint File Object
            Microsoft.Office.Interop.PowerPoint.Application pptApplication = Globals.ThisAddIn.Application;

            //現在のスライド番号
            int slideNumber = pptApplication.ActiveWindow.Selection.SlideRange.SlideNumber;

            Microsoft.Office.Interop.PowerPoint.SlideRange slideNow = pptApplication.ActivePresentation.Slides.Range(slideNumber);

            string strTemp = "\n";

            //音声ファイルを追加
            slideNow.Shapes.AddMediaObject2(strFullFileName, Microsoft.Office.Core.MsoTriState.msoCTrue, Microsoft.Office.Core.MsoTriState.msoCTrue, -10, -10, 50, 50);
            //セリフ文字をノート欄に追加
            slideNow.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text += strTemp + "\n" +"※"+ this.textBox1.Text + strTemp;

            MessageBox.Show("音声ファイルを追加しました！\n"+strFullFileName);
            //成功した場合、フォルダを開ける
            //System.Diagnostics.Process.Start(appPath);
        }


        private void label5_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("ver 1.0.2 音声ファイル自動挿入機能を追加しました。");
        }
    }
}
