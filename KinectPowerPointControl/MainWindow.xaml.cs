using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Kinect;
using Microsoft.Speech.Recognition;
using System.Threading;
using System.IO;
using Microsoft.Speech.AudioFormat;
using System.Diagnostics;
using System.Windows.Threading; 

namespace KinectPowerPointControl
{
    public partial class MainWindow : Window
    {
        /// <summary>
        /// kinect传感器对象
        /// </summary>
        KinectSensor sensor;

        /// <summary>
        /// 语音识别对象
        /// </summary>
        SpeechRecognitionEngine speechRecognizer;

        /// <summary>
        /// 计时器
        /// </summary>
        DispatcherTimer readyTimer;

        /// <summary>
        /// 颜色数据对象
        /// </summary>
        byte[] colorBytes;

        /// <summary>
        /// 骨骼数据对象
        /// </summary>
        Skeleton[] skeletons;

        /// <summary>
        /// 是否显示左上角圆形图案对象
        /// </summary>
        bool isCirclesVisible = true;

        /// <summary>
        /// forword状态变量
        /// </summary>
        bool isForwardGestureActive = false;

        /// <summary>
        /// back状态变量
        /// </summary>
        bool isBackGestureActive = false;

        /// <summary>
        /// 屏幕是否变黑标志
        /// </summary>
        bool isBlackScreenActive = false;

        /// <summary>
        /// PPT 放映标志
        /// </summary>
        bool isPresent = false;

        /// <summary>
        /// 手臂水平伸展的阈值
        /// </summary>
        private const double ArmStretchedThreshold = 0.45;

        /// <summary>
        /// 手臂垂直上举的阈值
        /// </summary>
        private const double ArmRaisedThreshold = 0.20;

        /// <summary>
        /// 头离双手距离的阈值
        /// </summary>
        private const double DistanceThreshold = 0.05; 

        /// <summary>
        /// 动作被激活时的颜色笔刷
        /// </summary>
        SolidColorBrush activeBrush = new SolidColorBrush(Colors.Green);

        /// <summary>
        /// 动作未被激活的颜色笔刷
        /// </summary>
        SolidColorBrush inactiveBrush = new SolidColorBrush(Colors.Red);

        public MainWindow()
        {
            InitializeComponent(); 
             
            //当窗体被打开，运行初始化方法，窗体关闭时，去初始化 
            this.Loaded += new RoutedEventHandler(MainWindow_Loaded);

            //Handle the content obtained from the video camera, once received. 
            this.KeyDown += new KeyEventHandler(MainWindow_KeyDown);
        }

        /// <summary>
        /// 开启传感器，获取数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            sensor = KinectSensor.KinectSensors.FirstOrDefault();

            if (sensor == null)
            {
                MessageBox.Show("This application requires a Kinect sensor.");
                this.Close();
            }

            sensor.Start();

            sensor.ColorStream.Enable(ColorImageFormat.RgbResolution640x480Fps30);
            sensor.ColorFrameReady += new EventHandler<ColorImageFrameReadyEventArgs>(sensor_ColorFrameReady);

            sensor.DepthStream.Enable(DepthImageFormat.Resolution320x240Fps30);

            sensor.SkeletonStream.Enable();
            sensor.SkeletonFrameReady += new EventHandler<SkeletonFrameReadyEventArgs>(sensor_SkeletonFrameReady);

            //sensor.ElevationAngle = 10;

            Application.Current.Exit += new ExitEventHandler(Current_Exit);

            InitializeSpeechRecognition();
        }

        /// <summary>
        /// 应用退出事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void Current_Exit(object sender, ExitEventArgs e)
        {
            if (speechRecognizer != null)
            {
                speechRecognizer.RecognizeAsyncCancel();
                speechRecognizer.RecognizeAsyncStop();
            }
            if (sensor != null)
            {
                sensor.AudioSource.Stop();
                sensor.Stop();
                sensor.Dispose();
                sensor = null;
            }
        }

        /// <summary>
        /// 键盘监听事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void MainWindow_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.C)
            {
                //点击键盘的C键，左上角的圆形图案消失
                ToggleCircles();
            }
        }

        /// <summary>
        /// 彩色摄像头事件处理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void sensor_ColorFrameReady(object sender, ColorImageFrameReadyEventArgs e)
        {
            using (var image = e.OpenColorImageFrame())
            {
                if (image == null)
                    return;

                if (colorBytes == null ||
                    colorBytes.Length != image.PixelDataLength)
                {
                    colorBytes = new byte[image.PixelDataLength];
                }

                image.CopyPixelDataTo(colorBytes);

                //You could use PixelFormats.Bgr32 below to ignore the alpha,
                //or if you need to set the alpha you would loop through the bytes 
                //as in this loop below
                int length = colorBytes.Length;
                for (int i = 0; i < length; i += 4)
                {
                    colorBytes[i + 3] = 255;
                }

                BitmapSource source = BitmapSource.Create(image.Width,
                    image.Height,
                    96,
                    96,
                    PixelFormats.Bgra32,
                    null,
                    colorBytes,
                    image.Width * image.BytesPerPixel);
                videoImage.Source = source;
            }
        }
        
        /// <summary>
        /// 骨骼事件处理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void sensor_SkeletonFrameReady(object sender, SkeletonFrameReadyEventArgs e)
        {
            using (var skeletonFrame = e.OpenSkeletonFrame())
            {
                if (skeletonFrame == null)
                    return;

                if (skeletons == null ||
                    skeletons.Length != skeletonFrame.SkeletonArrayLength)
                {
                    skeletons = new Skeleton[skeletonFrame.SkeletonArrayLength];
                }

                skeletonFrame.CopySkeletonDataTo(skeletons);
            }

            Skeleton closestSkeleton = skeletons.Where(s => s.TrackingState == SkeletonTrackingState.Tracked)
                                                .OrderBy(s => s.Position.Z * Math.Abs(s.Position.X))
                                                .FirstOrDefault();

            if (closestSkeleton == null)
                return;

            var head = closestSkeleton.Joints[JointType.Head];
            var rightHand = closestSkeleton.Joints[JointType.HandRight];
            var leftHand = closestSkeleton.Joints[JointType.HandLeft];

            if (head.TrackingState == JointTrackingState.NotTracked ||
                rightHand.TrackingState == JointTrackingState.NotTracked ||
                leftHand.TrackingState == JointTrackingState.NotTracked)
            {
                //Don't have a good read on the joints so we cannot process gestures
                return;
            }

            //调用填充头和双手位置图案的的方法
            SetEllipsePosition(ellipseHead, head, false);
            SetEllipsePosition(ellipseLeftHand, leftHand, isBackGestureActive);
            SetEllipsePosition(ellipseRightHand, rightHand, isForwardGestureActive);

            //调用处理手势的方法
            ProcessForwardBackGesture(head, rightHand, leftHand);
        }
         
        /// <summary>
        /// 该方法根据关节运动的跟踪数据，是用来定位的椭圆在画布上的位置 
        /// </summary>
        /// <param name="ellipse"></param>
        /// <param name="joint"></param>
        /// <param name="isHighlighted"></param>
        private void SetEllipsePosition(Ellipse ellipse, Joint joint, bool isHighlighted)
        {
            if (isHighlighted)
            {
                ellipse.Width = 60;
                ellipse.Height = 60;
                ellipse.Fill = activeBrush;
            }
            else
            {
                ellipse.Width = 20;
                ellipse.Height = 20;
                ellipse.Fill = inactiveBrush;
            }

            CoordinateMapper mapper = sensor.CoordinateMapper;
            
            //将三维空间坐标转化为UV平面坐标
            var point = mapper.MapSkeletonPointToColorPoint(joint.Position, sensor.ColorStream.Format);

            //调整绘制图案的位置
            Canvas.SetLeft(ellipse, point.X - ellipse.ActualWidth / 2);
            Canvas.SetTop(ellipse, point.Y - ellipse.ActualHeight / 2);
        }
        
        /// <summary>
        /// 处理手势的方法
        /// </summary>
        /// <param name="head"></param>
        /// <param name="rightHand"></param>
        /// <param name="leftHand"></param>
        private void ProcessForwardBackGesture(Joint head, Joint rightHand, Joint leftHand)
        {
            //若右手位置的横坐标值超过设定的阈值，除法PPT下一页命令
            if (rightHand.Position.X > head.Position.X + ArmStretchedThreshold)
            {
                if (!isForwardGestureActive)
                {
                    //激活forward命令，确保每次操作执行一次命令
                    isForwardGestureActive = true;

                    //模拟鼠标按下“右”方向键
                    System.Windows.Forms.SendKeys.SendWait("{Right}");
                }
            }
            else
            {
                isForwardGestureActive = false;
            }

            //若左手位置的横坐标超过设定的阈值，触发PPT上一页命令
            if (leftHand.Position.X < head.Position.X - ArmStretchedThreshold)
            {
                if (!isBackGestureActive)
                {
                    //激活back命令，确保每次操作执行一次命令
                    isBackGestureActive = true;

                    //模拟鼠标按下“左”方向键
                    System.Windows.Forms.SendKeys.SendWait("{Left}");
                }
            }
            else
            {
                isBackGestureActive = false;
            }


            //双手同时上举，在控制PPT时让屏幕变黑
            if ((leftHand.Position.Y > head.Position.Y - ArmRaisedThreshold) && (rightHand.Position.Y > head.Position.Y - ArmRaisedThreshold))
            {
                if (!isBlackScreenActive)
                {
                    isBlackScreenActive = true;
                    System.Windows.Forms.SendKeys.SendWait("{B}");
                }
            }
            else
                {
                    isBlackScreenActive = false;
                }

            //判断双手靠近头部，触发PPT放映
            if (Math.Abs(head.Position.Y - rightHand.Position.Y) < DistanceThreshold && (Math.Abs(head.Position.Y - leftHand.Position.Y) < DistanceThreshold && !isForwardGestureActive &&!isBackGestureActive))
            {
                if (!isPresent)
                { 
                    isPresent = true;
                    System.Windows.Forms.SendKeys.SendWait("{F5}");
                }
            }
            else
            {
                isPresent = false;
            }
        }
        
        /// <summary>
        /// 切换圆形图案的方法
        /// </summary>
        void ToggleCircles()
        {
            if (isCirclesVisible)
                HideCircles();
            else
                ShowCircles();
        }

        /// <summary>
        /// 隐藏圆图案的方法
        /// </summary>
        void HideCircles()
        {
            isCirclesVisible = false;
            ellipseHead.Visibility = System.Windows.Visibility.Collapsed;
            ellipseLeftHand.Visibility = System.Windows.Visibility.Collapsed;
            ellipseRightHand.Visibility = System.Windows.Visibility.Collapsed;
        }

        /// <summary>
        /// 显示圆形图案的方法
        /// </summary>
        void ShowCircles()
        {
            isCirclesVisible = true;
            ellipseHead.Visibility = System.Windows.Visibility.Visible;
            ellipseLeftHand.Visibility = System.Windows.Visibility.Visible;
            ellipseRightHand.Visibility = System.Windows.Visibility.Visible;
        }

        /// <summary>
        /// 显示窗体的方法
        /// </summary>
        private void ShowWindow()
        {
            this.Topmost = true;
            this.WindowState = System.Windows.WindowState.Maximized;
        }

        /// <summary>
        /// 隐藏窗体的方法
        /// </summary>
        private void HideWindow()
        {
            this.Topmost = false;
            this.WindowState = System.Windows.WindowState.Minimized;
        }

        #region Speech Recognition Methods

        private static RecognizerInfo GetKinectRecognizer()
        {
            Func<RecognizerInfo, bool> matchingFunc = r =>
            {
                string value;
                r.AdditionalInfo.TryGetValue("Kinect", out value);
                return "True".Equals(value, StringComparison.InvariantCultureIgnoreCase) && "en-US".Equals(r.Culture.Name, StringComparison.InvariantCultureIgnoreCase);
            };
            return SpeechRecognitionEngine.InstalledRecognizers().Where(matchingFunc).FirstOrDefault();
        }

        private void InitializeSpeechRecognition()
        {
            RecognizerInfo ri = GetKinectRecognizer();
            if (ri == null)
            {
                MessageBox.Show(
                    @"There was a problem initializing Speech Recognition.
Ensure you have the Microsoft Speech SDK installed.",
                    "Failed to load Speech SDK",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }

            try
            {
                speechRecognizer = new SpeechRecognitionEngine(ri.Id);
            }
            catch
            {
                MessageBox.Show(
                    @"There was a problem initializing Speech Recognition.
Ensure you have the Microsoft Speech SDK installed and configured.",
                    "Failed to load Speech SDK",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }

            var phrases = new Choices();
            phrases.Add("computer show window");
            phrases.Add("computer hide window");
            phrases.Add("computer show circles");
            phrases.Add("computer hide circles");

            var gb = new GrammarBuilder();
            //Specify the culture to match the recognizer in case we are running in a different culture.                                 
            gb.Culture = ri.Culture;
            gb.Append(phrases);

            // Create the actual Grammar instance, and then load it into the speech recognizer.
            var g = new Grammar(gb);

            speechRecognizer.LoadGrammar(g);
            speechRecognizer.SpeechRecognized += SreSpeechRecognized;
            speechRecognizer.SpeechHypothesized += SreSpeechHypothesized;
            speechRecognizer.SpeechRecognitionRejected += SreSpeechRecognitionRejected;

            this.readyTimer = new DispatcherTimer();
            this.readyTimer.Tick += this.ReadyTimerTick;
            this.readyTimer.Interval = new TimeSpan(0, 0, 4);
            this.readyTimer.Start();

        }

        private void ReadyTimerTick(object sender, EventArgs e)
        {
            this.StartSpeechRecognition();
            this.readyTimer.Stop();
            this.readyTimer.Tick -= ReadyTimerTick;
            this.readyTimer = null;
        }

        private void StartSpeechRecognition()
        {
            if (sensor == null || speechRecognizer == null)
                return;

            var audioSource = this.sensor.AudioSource;
            audioSource.BeamAngleMode = BeamAngleMode.Adaptive;
            var kinectStream = audioSource.Start();

            speechRecognizer.SetInputToAudioStream(
                    kinectStream, new SpeechAudioFormatInfo(EncodingFormat.Pcm, 16000, 16, 1, 32000, 2, null));
            speechRecognizer.RecognizeAsync(RecognizeMode.Multiple);

        }

        void SreSpeechRecognitionRejected(object sender, SpeechRecognitionRejectedEventArgs e)
        {
            Trace.WriteLine("\nSpeech Rejected, confidence: " + e.Result.Confidence);
        }

        void SreSpeechHypothesized(object sender, SpeechHypothesizedEventArgs e)
        {
            Trace.Write("\rSpeech Hypothesized: \t{0}", e.Result.Text);
        }

        void SreSpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            //This first release of the Kinect language pack doesn't have a reliable confidence model, so 
            //we don't use e.Result.Confidence here.
            if (e.Result.Confidence < 0.70)
            {
                Trace.WriteLine("\nSpeech Rejected filtered, confidence: " + e.Result.Confidence);
                return;
            }

            Trace.WriteLine("\nSpeech Recognized, confidence: " + e.Result.Confidence + ": \t{0}", e.Result.Text);

            if (e.Result.Text == "computer show window")
            {
                this.Dispatcher.BeginInvoke((Action)delegate
                    {
                        ShowWindow();
                    });
            }
            else if (e.Result.Text == "computer hide window")
            {
                this.Dispatcher.BeginInvoke((Action)delegate
                {
                    HideWindow();
                });
            }
            else if (e.Result.Text == "computer hide circles")
            {
                this.Dispatcher.BeginInvoke((Action)delegate
                {
                    this.HideCircles();
                });
            }
            else if (e.Result.Text == "computer show circles")
            {
                this.Dispatcher.BeginInvoke((Action)delegate
                {
                    this.ShowCircles();
                });
            }
        }

        #endregion

    }
}
