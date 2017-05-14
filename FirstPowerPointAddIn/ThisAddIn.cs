using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Kinect;
using Microsoft.Kinect.VisualGestureBuilder;
using Microsoft.Samples.Kinect.DiscreteGestureBasics;
using Microsoft.Office.Interop.PowerPoint;

namespace FirstPowerPointAddIn
{
    public partial class ThisAddIn
    {

        /// <summary> Active Kinect sensor </summary>
        private static KinectSensor kinectSensor = null;

        /// <summary> Array for the bodies (Kinect will track up to 6 people simultaneously) </summary>
        private static Body[] bodies = null;

        /// <summary> Reader for body frames </summary>
        private static BodyFrameReader bodyFrameReader = null;

        /// <summary> List of gesture detectors, there will be one detector created for each potential body (max of 6) </summary>
        private static List<GestureDetector> gestureDetectorList = null;

        public static bool SwipeOnRight = false;
        public static bool SwipeOnLeft = false;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            // only one sensor is currently supported
            kinectSensor = KinectSensor.GetDefault();

            // set IsAvailableChanged event notifier
            kinectSensor.IsAvailableChanged += Sensor_IsAvailableChanged;

            // open the sensor
            kinectSensor.Open();

            // open the reader for the body frames
            bodyFrameReader = kinectSensor.BodyFrameSource.OpenReader();

            // initialize the gesture detection objects for our gestures
            gestureDetectorList = new List<GestureDetector>();

            // create a gesture detector for each body (6 bodies => 6 detectors) and create content controls to display results in the UI
            int maxBodies = 6;
            for (int i = 0; i < maxBodies; ++i)
            {
                GestureResultView result = new GestureResultView(i, false, false, 0.0f);
                GestureDetector detector = new GestureDetector(kinectSensor, result);
                gestureDetectorList.Add(detector);
            }

            // set the BodyFramedArrived event notifier
            bodyFrameReader.FrameArrived += Reader_BodyFrameArrived;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (kinectSensor != null)
            {
                kinectSensor.Close();
                kinectSensor = null;
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        private static void ExecuteOnSlideShow(Action<SlideShowView> toExecute)
        {
            try
            {

                var view = GetSlideShowView();
                if (view == null) return;
                if (view.State == PpSlideShowState.ppSlideShowRunning)
                {
                    toExecute(view);
                }
                else
                {
                }
            }
            catch (Exception ex)
            {
            }
        }

        private static SlideShowView GetSlideShowView()
        {
            try
            {
                var presentation = Globals.ThisAddIn.Application.ActivePresentation;
                if (presentation == null)
                {
                    return null;
                }
                if (presentation.SlideShowWindow == null)
                {
                    return null;
                }
                if (presentation.SlideShowWindow.View == null)
                {
                    return null;
                }
                return presentation.SlideShowWindow.View;
            }
            catch (Exception ex)
            {

            }
            return null;
        }

        public static void Monitor_SwipeOnRight(bool mode)
        {
            if (SwipeOnRight == false && mode == true)
            {
                ExecuteOnSlideShow(view => view.Next());
                SwipeOnRight = true;
            }
        }

        public static void Monitor_SwipeOnLeft(bool mode)
        {
            if (SwipeOnLeft == false && mode == true)
            {
                ExecuteOnSlideShow(view => view.Previous());
                SwipeOnLeft = true;
            }
        }

        /// <summary>
        /// Handles the event when the sensor becomes unavailable (e.g. paused, closed, unplugged).
        /// </summary>
        /// <param name="sender">object sending the event</param>
        /// <param name="e">event arguments</param>
        private static void Sensor_IsAvailableChanged(object sender, IsAvailableChangedEventArgs e)
        {
            // on failure, set the status text
        }

        void Application_PresentationNewSlide(PowerPoint.Slide Sld)
        {
            PowerPoint.Shape textBox = Sld.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 500, 50);
            textBox.TextFrame.TextRange.InsertAfter("This text was added by using code.");
        }

        private static void Reader_BodyFrameArrived(object sender, BodyFrameArrivedEventArgs e)
        {
            bool dataReceived = false;

            using (BodyFrame bodyFrame = e.FrameReference.AcquireFrame())
            {
                if (bodyFrame != null)
                {
                    if (bodies == null)
                    {
                        // creates an array of 6 bodies, which is the max number of bodies that Kinect can track simultaneously
                        bodies = new Body[bodyFrame.BodyCount];
                    }

                    // The first time GetAndRefreshBodyData is called, Kinect will allocate each Body in the array.
                    // As long as those body objects are not disposed and not set to null in the array,
                    // those body objects will be re-used.
                    bodyFrame.GetAndRefreshBodyData(bodies);
                    dataReceived = true;
                }
            }

            if (dataReceived)
            {
                // we may have lost/acquired bodies, so update the corresponding gesture detectors
                if (bodies != null)
                {
                    // loop through all bodies to see if any of the gesture detectors need to be updated
                    int maxBodies = kinectSensor.BodyFrameSource.BodyCount;
                    for (int i = 0; i < maxBodies; ++i)
                    {
                        Body body = bodies[i];
                        ulong trackingId = body.TrackingId;

                        // if the current body TrackingId changed, update the corresponding gesture detector with the new value
                        if (trackingId != gestureDetectorList[i].TrackingId)
                        {
                            gestureDetectorList[i].TrackingId = trackingId;

                            // if the current body is tracked, unpause its detector to get VisualGestureBuilderFrameArrived events
                            // if the current body is not tracked, pause its detector so we don't waste resources trying to get invalid gesture results
                            gestureDetectorList[i].IsPaused = trackingId == 0;
                        }
                    }
                }
            }
        }
    }
}
