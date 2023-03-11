// MIT License
//Copyright (c) 2023 Sebastian Kotstein
//
//Permission is hereby granted, free of charge, to any person obtaining a copy
//of this software and associated documentation files (the "Software"), to deal
//in the Software without restriction, including without limitation the rights
//to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
//copies of the Software, and to permit persons to whom the Software is
//furnished to do so, subject to the following conditions:
//
//The above copyright notice and this permission notice shall be included in all
//copies or substantial portions of the Software.
//
//THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
//IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
//FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
//AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
//LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
//OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
//SOFTWARE.
using Microsoft.Kinect;
using Microsoft.Office.Tools.Ribbon;
using SKotstein.Kinect.API.Core;
using SKotstein.Kinect.API.Core.Root;
using SKotstein.Kinect.API.Gestures;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.Text;
using System.Windows;
using System.Windows.Forms;

namespace skotstein.app.kinect.powerpoint_add_in
{

    public partial class comboBoxVoices
    {
        private HumanApi _humanApi;

        private SpeechRecognitionEngine _recEngine = null;
        private SpeechSynthesizer _synthesizer = new SpeechSynthesizer();
        private ResourceManager _resourceManagerForCommands;
        private ResourceManager _resourceManagerForVoice;


        private void KinectRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            //initialize human API
            _humanApi = HumanApi.GetInstance(KinectSensor.GetDefault());

            //prepare speech recognition engine
            this.comboBoxLanguage.Items.Clear();
            RibbonDropDownItem disabledItem = this.Factory.CreateRibbonDropDownItem();
            disabledItem.Label = "Disabled";
            this.comboBoxLanguage.Items.Add(disabledItem);
            foreach (RecognizerInfo config in SpeechRecognitionEngine.InstalledRecognizers())
            {
                RibbonDropDownItem item = this.Factory.CreateRibbonDropDownItem();
                item.Label = config.Culture.ToString();
                this.comboBoxLanguage.Items.Add(item);
            }
            this.comboBoxLanguage.Text = this.comboBoxLanguage.Items[0].Label;
            this.comboBoxLanguage.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.comboBoxLanguage_TextChanged);

            this.comboBoxVoice.Items.Clear();
            foreach(InstalledVoice installedVoices in _synthesizer.GetInstalledVoices())
            {

                RibbonDropDownItem item = this.Factory.CreateRibbonDropDownItem();
                item.Label = installedVoices.VoiceInfo.Culture.ToString() + ", " + installedVoices.VoiceInfo.Gender.ToString()+ ", "+installedVoices.VoiceInfo.Age.ToString();
                this.comboBoxVoice.Items.Add(item);
            }
            if(this.comboBoxVoice.Items.Count > 0)
            {
                this.comboBoxVoice.Text = this.comboBoxVoice.Items[0].Label;
                evalVoiceComboBox();
            }
            this.comboBoxVoice.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.comboBoxVoice_TextChanged);

        }

        /// <summary>
        /// Returns the voice command having the passed key in the loaded command resource dictionary.
        /// The method returns null, if no command resource dictionary has been loaded or no command is associated with the specified key.
        /// </summary>
        /// <param name="key">Key of the command</param>
        /// <returns>command or null</returns>
        private string GetCmdResource(string key)
        {
            if (_resourceManagerForCommands != null)
            {
                return _resourceManagerForCommands.GetString(key);
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Returns the voice response having the passed key in the loaded voice resource dictionary.
        /// The method returns null, if no voice resource dictionary has been loaded or no response is associated with the specified key.
        /// </summary>
        /// <param name="key">Key of the voice response</param>
        /// <returns>voice response or null</returns>
        private string GetVoiceResource(string key)
        {
            if (_resourceManagerForVoice != null)
            {
                return _resourceManagerForVoice.GetString(key);
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// This method is executed whenever the user makes a selection in the voice command recognizer combo box (labeled as "Input Language"), i.e., changes the input language.
        /// The method unloads the current voice command recognier and loads the selected recognizer and its grammar, unless the user disables voice command recognition. 
        /// </summary>
        private void evalLanguageComboBox()
        {
            UnloadGrammar();
            string language = this.comboBoxLanguage.Text;
            if (language.CompareTo("Disabled") != 0)
            {
                if (!LoadGrammar(language))
                {
                    this.comboBoxLanguage.TextChanged -= new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.comboBoxLanguage_TextChanged);
                    this.comboBoxLanguage.Text = "Disabled";
                    this.comboBoxLanguage.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.comboBoxLanguage_TextChanged);
                }
            }
        }

        /// <summary>
        /// Loads the grammar being associated with the passed language code and prepares the respective voice recognizer.
        /// The method returns true, if the operation can be completed successfully, else false.
        /// </summary>
        /// <param name="languageCode">ISO language code, e.g. "en-US". The method returns false, if the passed language code is unknown to the system.</param>
        /// <returns>true or false</returns>
        private bool LoadGrammar(String languageCode)
        {
            if (languageCode.StartsWith("en-"))
            {
                _resourceManagerForCommands = new ResourceManager("skotstein.app.kinect.powerpoint_add_in.commands", Assembly.GetExecutingAssembly());
            }
            else if (languageCode.StartsWith("de-"))
            {
                _resourceManagerForCommands = new ResourceManager("skotstein.app.kinect.powerpoint_add_in.commands_de_DE", Assembly.GetExecutingAssembly());
            }/*
            else if (languageCode.StartsWith("es-"))
            {
                _resourceManagerForCommands = new ResourceManager("skotstein.app.kinect.powerpoint_add_in.commands_es_ES", Assembly.GetExecutingAssembly());
            }
            else if (languageCode.StartsWith("fr-"))
            {
                _resourceManagerForCommands = new ResourceManager("skotstein.app.kinect.powerpoint_add_in.commands_fr_FR", Assembly.GetExecutingAssembly());
            }
            else if (languageCode.StartsWith("it-"))
            {
                _resourceManagerForCommands = new ResourceManager("skotstein.app.kinect.powerpoint_add_in.commands_it_IT", Assembly.GetExecutingAssembly());
            }*/
            else
            {
                return false;
            }

            if(_resourceManagerForCommands != null)
            {
                Choices choices = new Choices();
                choices.Add(new String[] { GetCmdResource("cmd_start_slide_show"), GetCmdResource("cmd_end_slide_show"), GetCmdResource("cmd_start_recognition"), GetCmdResource("cmd_end_recognition"), GetCmdResource("cmd_is_body_tracked"), GetCmdResource("cmd_voice_options") });
                GrammarBuilder grammarBuilder = new GrammarBuilder();
                grammarBuilder.Culture = new System.Globalization.CultureInfo(languageCode);
                grammarBuilder.Append(choices);
                Grammar grammar = new Grammar(grammarBuilder);

                try
                {
                    _recEngine = new SpeechRecognitionEngine(new System.Globalization.CultureInfo(languageCode));
                    _recEngine.LoadGrammarAsync(grammar);
                    _recEngine.SetInputToDefaultAudioDevice();
                    _recEngine.SpeechRecognized += _recEngine_SpeechRecognized;
                    _recEngine.RecognizeAsync(RecognizeMode.Multiple);
                    return true;
                }
                catch (Exception e)
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
             
        }

        /// <summary>
        /// The method unloads the current voice command recognier and its grammar.
        /// </summary>
        private void UnloadGrammar()
        {
            try
            {
                if (_recEngine != null)
                {
                    _recEngine.UnloadAllGrammars();
                    _recEngine.SetInputToNull();
                    _recEngine.SpeechRecognized -= _recEngine_SpeechRecognized;
                    _recEngine.Dispose();
                }
            }
            catch (Exception e)
            {

            }
        }

        /// <summary>
        /// This method is executed whenever the user makes a selection in the voice response combo box (labeled as "Output Voice"), i.e., changes the output language.
        /// The method loads the selected voice response synthesizer.
        /// </summary>
        private void evalVoiceComboBox()
        {
            string voice = this.comboBoxVoice.Text;
            if (!String.IsNullOrWhiteSpace(voice) && voice.Split(',').Length == 3)
            {
                string language = voice.Split(',')[0].Trim();
                string gender = voice.Split(',')[1].Trim();
                string age = voice.Split(',')[2].Trim();

                foreach (InstalledVoice installedVoice in _synthesizer.GetInstalledVoices())
                {
                    if (installedVoice.VoiceInfo.Culture.ToString().CompareTo(language) == 0
                        && installedVoice.VoiceInfo.Gender.ToString().CompareTo(gender) == 0
                        && installedVoice.VoiceInfo.Age.ToString().CompareTo(age) == 0)
                    {
                        _synthesizer.SelectVoice(installedVoice.VoiceInfo.Name);
                    }
                }
                if (language.StartsWith("en-"))
                {
                    _resourceManagerForVoice = new ResourceManager("skotstein.app.kinect.powerpoint_add_in.voice", Assembly.GetExecutingAssembly());
                }
                else if (language.StartsWith("de-"))
                {
                    _resourceManagerForVoice = new ResourceManager("skotstein.app.kinect.powerpoint_add_in.voice_de_DE", Assembly.GetExecutingAssembly());
                }
                /*
                else if (language.StartsWith("es-"))
                {
                    _resourceManagerForVoice = new ResourceManager("skotstein.app.kinect.powerpoint_add_in.voice_es_ES", Assembly.GetExecutingAssembly());
                }
                else if (language.StartsWith("fr-"))
                {
                    _resourceManagerForVoice = new ResourceManager("skotstein.app.kinect.powerpoint_add_in.voice_fr_FR", Assembly.GetExecutingAssembly());
                }
                else if (language.StartsWith("it-"))
                {
                    _resourceManagerForVoice = new ResourceManager("skotstein.app.kinect.powerpoint_add_in.voice_it_IT", Assembly.GetExecutingAssembly());
                }*/
                else
                {
                    _resourceManagerForVoice = null;
                }
            }
        }

        /// <summary>
        /// Starts gesture recognition
        /// </summary>
        private void Start()
        {
            //human API
            _humanApi.BodyDetected += Body_Detected;
            //_humanApi.BodyLost += Api_BodyLost;
            _humanApi.Start();
        }

        /*
        private void Api_BodyLost(object sender, BodyEventArgs e)
        {
            if (_humanApi.AmountOfBodiesDetected > 0)
            {
                
            }
        }
        */

        /// <summary>
        /// Stops gesture recognition
        /// </summary>
        private void Stop()
        {
            _humanApi.BodyDetected -= Body_Detected;
            _humanApi.Stop();
        }

        /// <summary>
        /// Registers pre-defined gestures whenever a (new) body is detected
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Body_Detected(object sender, BodyEventArgs e)
        {

            IGestureContainer motionContainer = new AutomatedGestureContainer(MotionGestureFactory.GetInstance());
            IGestureContainer handContainer = new AutomatedGestureContainer(HandGestureFactory.GetInstance());
            IBodyController bc = e.BodyController;
            bc.LoadGestureContainer(motionContainer);
            bc.LoadGestureContainer(handContainer);
            bc.AddGestureEventHandler(Gesture_Handler, GestureIdentifier.LEFT_HAND_CLOSED_GESTURE);
            bc.AddGestureEventHandler(Gesture_Handler, GestureIdentifier.RIGHT_HAND_CLOSED_GESTURE);
            bc.AddGestureEventHandler(Gesture_Handler, GestureIdentifier.SWIPE_TO_LEFT_GESTURE);
            bc.AddGestureEventHandler(Gesture_Handler, GestureIdentifier.CIRCLE_COUNTER_CLOCKWISE_GESTURE);
        }

        /// <summary>
        /// Translates a recognized gesture into a PowerPoint command, technically a key stroke (e.g. "ESC" for ending slide show).
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Gesture_Handler(object sender, GestureEventArgs e)
        {
            switch (e.GestureIdentifier)
            {

                case GestureIdentifier.CIRCLE_COUNTER_CLOCKWISE_GESTURE:
                    if (_humanApi.IsClosestBody(e.TrackingId))
                    {
                        SendKeys.SendWait("{LEFT}");
                    }

                    break;
                case GestureIdentifier.SWIPE_TO_LEFT_GESTURE:
                    if (_humanApi.IsClosestBody(e.TrackingId))
                    {
                        SendKeys.SendWait("{RIGHT}");
                    }
                    break;
                case GestureIdentifier.RIGHT_HAND_CLOSED_GESTURE:
                    if (_humanApi.IsClosestBody(e.TrackingId))
                    {
                        SendKeys.SendWait("{ESC}");
                    }
                    break;
                case GestureIdentifier.LEFT_HAND_CLOSED_GESTURE:
                    if (_humanApi.IsClosestBody(e.TrackingId))
                    {
                        SendKeys.SendWait("{F5}");
                    }
                    break;


            }
        }

        /// <summary>
        /// Translates a recognized voice command into a PowerPoint command, technically a key stroke (e.g. "ESC" for ending slide show).
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void _recEngine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            if (String.IsNullOrWhiteSpace(e.Result.Text))
            {

            }
            else if (e.Result.Text.CompareTo(GetCmdResource("cmd_start_recognition")) == 0)
            {
                if(_resourceManagerForVoice != null)
                {
                    _synthesizer.SpeakAsync(GetVoiceResource("voice_recognition_started"));
                }
                this.Start();
            }
            else if (e.Result.Text.CompareTo(GetCmdResource("cmd_end_recognition")) == 0)
            {
                if (_resourceManagerForVoice != null)
                {
                    _synthesizer.SpeakAsync(GetVoiceResource("voice_recognition_ended"));
                }
                this.Stop();
            }
            else if (e.Result.Text.CompareTo(GetCmdResource("cmd_start_slide_show")) == 0)
            {
                if (_resourceManagerForVoice != null)
                {
                    _synthesizer.SpeakAsync(GetVoiceResource("voice_slide_show_started"));
                }
                SendKeys.SendWait("{F5}");
            }
            else if (e.Result.Text.CompareTo(GetCmdResource("cmd_end_slide_show")) == 0)
            {
                if (_resourceManagerForVoice != null)
                {
                    _synthesizer.SpeakAsync(GetVoiceResource("voice_slide_show_ended"));
                }
                SendKeys.SendWait("{ESC}");
            }
            else if (e.Result.Text.CompareTo(GetCmdResource("cmd_is_body_tracked")) == 0)
            {
                if (_humanApi.AmountOfBodiesDetected > 0 && _resourceManagerForVoice != null)
                {
                    _synthesizer.SpeakAsync(GetVoiceResource("voice_body_is_tracked"));
                }
                else
                {
                    _synthesizer.SpeakAsync(GetVoiceResource("voice_body_is_not_tracked"));
                }
            }
            else if (e.Result.Text.CompareTo(GetCmdResource("cmd_voice_options")) == 0)
            {
                if (_resourceManagerForVoice != null)
                {
                    _synthesizer.SpeakAsync(GetVoiceResource("voice_options")+" "+ GetCmdResource("cmd_start_recognition")+", "+ GetCmdResource("cmd_end_recognition")+", "+ GetCmdResource("cmd_start_slide_show")+", "+ GetCmdResource("cmd_end_slide_show")+", "+ GetCmdResource("cmd_is_body_tracked"));
                }
                //_synthesizer.SpeakAsync("You can say Start Presentation, Finish Presentation, Start Motion Control, Finish Motion Control or ask Am I detected?");
            }
        }

        private void checkBoxGesture_Click(object sender, RibbonControlEventArgs e)
        {
            
        }

        private void buttonStart_Click(object sender, RibbonControlEventArgs e)
        {
            Start();
        }

        private void buttonStop_Click(object sender, RibbonControlEventArgs e)
        {
            Stop();
        }

        private void comboBoxLanguage_TextChanged(object sender, RibbonControlEventArgs e)
        {
            evalLanguageComboBox();
        }

        private void comboBoxVoice_TextChanged(object sender, RibbonControlEventArgs e)
        {
            evalVoiceComboBox();
        }
    }
}
