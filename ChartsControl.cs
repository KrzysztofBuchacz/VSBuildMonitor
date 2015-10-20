using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;

namespace VSBuildMonitor
{
   public partial class ChartsControl : UserControl
   {
      public Connect host;
      public bool scrollLast = true;
      public string intFormat = "D3";
      public bool isBuilding = false;
      private bool isDark = false;

      Brush blueSolidBrush = new SolidBrush(System.Drawing.Color.DarkBlue);
      Pen blackPen = new Pen(Color.Black, 1);
      Brush blackBrush = new SolidBrush(Color.Black);
      Brush greenSolidBrush = new SolidBrush(System.Drawing.Color.DarkGreen);
      Brush redSolidBrush = new SolidBrush(System.Drawing.Color.DarkRed);
      Brush whiteBrush = new SolidBrush(Color.White);
      Pen grid = new Pen(Color.LightGray);

      public ChartsControl()
      {
         InitializeComponent();

         isDark = GetThemeId().ToLower() == "1ded0138-47ce-435e-84ef-9ec1f439b749";

         blackBrush = new SolidBrush(isDark ? System.Drawing.Color.White : System.Drawing.Color.Black);
         BackColor = isDark ? System.Drawing.Color.Black : System.Drawing.Color.White;
         greenSolidBrush = new SolidBrush(isDark ? System.Drawing.Color.LightGreen : System.Drawing.Color.DarkGreen);
         redSolidBrush = new SolidBrush(isDark ? Color.FromArgb(165, 33, 33) : System.Drawing.Color.DarkRed);
         blueSolidBrush = new SolidBrush(isDark ? Color.FromArgb(50, 152, 204) : System.Drawing.Color.DarkBlue);
         blackPen = new Pen(isDark ? Color.White : Color.Black, 1);
         grid = new Pen(isDark ? Color.FromArgb(66, 66, 66) : Color.LightGray);
      }

      protected override void OnPaint(PaintEventArgs e)
      {
         if (DisplayRectangle.Size.Width < 10 || DisplayRectangle.Size.Height < 10)
         {
            return;
         }

         Graphics graphicsObj = e.Graphics;

         int rowHeight = Font.Height + 1;
         int linesCount = host.currentBuilds.Count + host.finishedBuilds.Count + 1;
         int totalHeight = rowHeight * linesCount;

         AutoScrollMinSize = new Size(this.AutoScrollMinSize.Width, totalHeight);
         if (scrollLast)
         {
            int newPos = totalHeight - Size.Height;
            AutoScrollPosition = new Point(0, newPos);
         }

         Matrix mx = new Matrix(1, 0, 0, 1, AutoScrollPosition.X, AutoScrollPosition.Y);
         e.Graphics.Transform = mx;

         int maxStringLength = 0;
         long tickStep = 100000000;
         long maxTick = tickStep;
         long nowTick = DateTime.Now.Ticks;
         long t = nowTick - host.buildTime.Ticks;
         if (host.currentBuilds.Count > 0)
         {
            if (t > maxTick)
            {
               maxTick = t;
            }
         }
         int i;
         for (i = 0; i < host.finishedBuilds.Count; i++)
         {
            int l = graphicsObj.MeasureString(host.finishedBuilds[i].name, Font).ToSize().Width;
            t = host.finishedBuilds[i].end;
            if (t > maxTick)
            {
               maxTick = t;
            }
            if (l > maxStringLength)
            {
               maxStringLength = l;
            }
         }
         foreach (KeyValuePair<string, DateTime> item in host.currentBuilds)
         {
            int l = graphicsObj.MeasureString(item.Key, Font).ToSize().Width;
            if (l > maxStringLength)
            {
               maxStringLength = l;
            }
         }
         if (isBuilding)
         {
            maxTick = (maxTick / tickStep + 1) * tickStep; // round up
         }
         maxStringLength += 5 + graphicsObj.MeasureString(i.ToString(intFormat) + " ", Font).ToSize().Width;

         Brush greenGradientBrush = new LinearGradientBrush(new Point(0, 0), new Point(0, rowHeight), System.Drawing.Color.MediumSeaGreen, System.Drawing.Color.DarkGreen);
         Brush redGradientBrush = new LinearGradientBrush(new Point(0, 0), new Point(0, rowHeight), System.Drawing.Color.IndianRed, System.Drawing.Color.DarkRed);
         for (i = 0; i < host.finishedBuilds.Count; i++)
         {
            Brush solidBrush = host.finishedBuilds[i].success ? greenSolidBrush : redSolidBrush;
            Brush gradientBrush = host.finishedBuilds[i].success ? greenGradientBrush : redGradientBrush;
            DateTime span = new DateTime(host.finishedBuilds[i].end - host.finishedBuilds[i].begin);
            string time = span.ToString(Connect.timeFormat);
            graphicsObj.DrawString((i + 1).ToString(intFormat) + " " + host.finishedBuilds[i].name, Font, solidBrush, 1, i * rowHeight);
            Rectangle r = new Rectangle();
            r.X = maxStringLength + (int)((host.finishedBuilds[i].begin) * (long)(DisplayRectangle.Size.Width - maxStringLength) / maxTick);
            r.Width = maxStringLength + (int)((host.finishedBuilds[i].end) * (long)(DisplayRectangle.Size.Width - maxStringLength) / maxTick) - r.X;
            if (r.Width == 0)
            {
               r.Width = 1;
            }
            r.Y = i * rowHeight + 1;
            r.Height = rowHeight - 1;
            graphicsObj.FillRectangle(gradientBrush, r);
            int timeLen = graphicsObj.MeasureString(time, Font).ToSize().Width;
            if (r.Width > timeLen)
            {
               graphicsObj.DrawString(time, Font, whiteBrush, r.Right - timeLen, i * rowHeight);
            }
            graphicsObj.DrawLine(grid, new Point(0, i * rowHeight), new Point(DisplayRectangle.Size.Width, i * rowHeight));
         }

         Brush blueGradientBrush = new LinearGradientBrush(new Point(0, 0), new Point(0, rowHeight), System.Drawing.Color.LightBlue, System.Drawing.Color.DarkBlue);
         foreach (KeyValuePair<string, DateTime> item in host.currentBuilds)
         {
            graphicsObj.DrawString((i + 1).ToString(intFormat) + " " + item.Key, Font, blueSolidBrush, 1, i * rowHeight);
            Rectangle r = new Rectangle();
            r.X = maxStringLength + (int)((item.Value.Ticks - host.buildTime.Ticks) * (long)(DisplayRectangle.Size.Width - maxStringLength) / maxTick);
            r.Width = maxStringLength + (int)((nowTick - host.buildTime.Ticks) * (long)(DisplayRectangle.Size.Width - maxStringLength) / maxTick) - r.X;
            if (r.Width == 0)
            {
               r.Width = 1;
            }
            r.Y = i * rowHeight + 1;
            r.Height = rowHeight - 1;
            graphicsObj.FillRectangle(blueGradientBrush, r);
            graphicsObj.DrawLine(grid, new Point(0, i * rowHeight), new Point(DisplayRectangle.Size.Width, i * rowHeight));
            i++;
         }

         if (host.currentBuilds.Count > 0 || host.finishedBuilds.Count > 0)
         {
            string line = "";
            if (isBuilding)
            {
               line = "Building...";
            }
            else
            {
               line = "Done";
            }
            if (host.maxParallelBuilds > 0)
            {
               line += " (" + host.PercentageProcessorUse().ToString() + "% of " + host.maxParallelBuilds.ToString() + " CPUs)";
            }
            graphicsObj.DrawString(line, Font, blackBrush, 1, i * rowHeight);
         }
         graphicsObj.DrawLine(grid, new Point(maxStringLength, 0), new Point(maxStringLength, i * rowHeight));
         graphicsObj.DrawLine(blackPen, new Point(0, i * rowHeight), new Point(DisplayRectangle.Size.Width, i * rowHeight));
         DateTime dt = new DateTime(0);
         graphicsObj.DrawString(dt.ToString(Connect.timeFormat), Font, blackBrush, maxStringLength, i * rowHeight);

         dt = new DateTime(maxTick);
         string s = dt.ToString(Connect.timeFormat);
         int m = graphicsObj.MeasureString(s, Font).ToSize().Width;
         graphicsObj.DrawString(s, Font, blackBrush, DisplayRectangle.Size.Width - m, i * rowHeight);
      }

      private void ChartsControl_Scroll(object sender, ScrollEventArgs e)
      {
         if (e.ScrollOrientation == ScrollOrientation.VerticalScroll)
         {
            if (e.Type == ScrollEventType.LargeIncrement || e.Type == ScrollEventType.SmallIncrement)
            {
               if (e.NewValue == e.OldValue)
               {
                  scrollLast = true;
               }
               else
               {
                  scrollLast = false;
               }
            }
            else
            {
               scrollLast = false;
            }
         }
      }

      private string GetThemeId()
      {
         const string CATEGORY_TEXT_GENERAL = "General";
         const string PROPERTY_NAME_CURRENT_THEME = "CurrentTheme";

         string result;

         result = Microsoft.Win32.Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\VisualStudio\11.0\"
            + CATEGORY_TEXT_GENERAL, PROPERTY_NAME_CURRENT_THEME, "").ToString();

         return result;
      }

   }
}
