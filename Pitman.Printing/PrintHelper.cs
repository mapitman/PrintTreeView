// ------------------------------------------------------------------------
// <copyright file="PrintHelper.cs" company="Mark Pitman">
//    Copyright 2007 - 2013 Mark Pitman
//    
//    Licensed under the Apache License, Version 2.0 (the "License");
//    you may not use this file except in compliance with the License.
//    You may obtain a copy of the License at
// 
//      http://www.apache.org/licenses/LICENSE-2.0
// 
//    Unless required by applicable law or agreed to in writing, software
//    distributed under the License is distributed on an "AS IS" BASIS,
//    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
//    See the License for the specific language governing permissions and
//    limitations under the License.
// </copyright>
// ------------------------------------------------------------------------
// 

namespace Pitman.Printing
{
    using System;
    using System.Drawing;
    using System.Drawing.Printing;
    using System.Globalization;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;

    public class PrintHelper
    {
        #region Fields

        private readonly PrintDocument printDoc;

        private Image controlImage;

        private PrintDirection currentDir;

        private Point lastPrintPosition;

        private int nodeHeight;

        private int pageNumber;

        private int scrollBarHeight;

        private int scrollBarWidth;

        private string title = string.Empty;

        #endregion

        #region Constructors and Destructors

        public PrintHelper()
        {
            this.lastPrintPosition = new Point(0, 0);
            this.printDoc = new PrintDocument();
            this.printDoc.BeginPrint += this.PrintDocBeginPrint;
            this.printDoc.PrintPage += this.PrintDocPrintPage;
            this.printDoc.EndPrint += this.PrintDocEndPrint;
        }

        #endregion

        #region Enums

        private enum PrintDirection
        {
            Horizontal,

            Vertical
        }

        #endregion

        #region Public Methods and Operators

        /// <summary>
        ///     Shows a PrintPreview dialog displaying the Tree control passed in.
        /// </summary>
        /// <param name="tree">TreeView to print preview</param>
        /// <param name="reportTitle"></param>
        public void PrintPreviewTree(TreeView tree, string reportTitle)
        {
            this.title = reportTitle;
            this.PrepareTreeImage(tree);
            var pp = new PrintPreviewDialog { Document = this.printDoc };
            pp.Show();
        }

        /// <summary>
        ///     Prints a tree
        /// </summary>
        /// <param name="tree">TreeView to print</param>
        /// <param name="reportTitle"></param>
        public void PrintTree(TreeView tree, string reportTitle)
        {
            this.title = reportTitle;
            this.PrepareTreeImage(tree);
            var pd = new PrintDialog { Document = this.printDoc };
            if (pd.ShowDialog() == DialogResult.OK)
            {
                this.printDoc.Print();
            }
        }

        #endregion

        #region Methods

        [DllImport("gdi32.dll")]
        private static extern IntPtr CreateCompatibleBitmap(IntPtr hDC, int width, int height);

        [DllImport("user32.dll")]
        private static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, int lParam);

        // Returns an image of the specified width and height, of a control represented by handle.
        private Image GetImage(IntPtr handle, int width, int height)
        {
            IntPtr screenDC = GetDC(IntPtr.Zero);
            IntPtr hbm = CreateCompatibleBitmap(screenDC, width, height);
            Image image = Image.FromHbitmap(hbm);
            Graphics g = Graphics.FromImage(image);
            IntPtr hdc = g.GetHdc();
            SendMessage(handle, 0x0318 /*WM_PRINTCLIENT*/, hdc, (0x00000010 | 0x00000004 | 0x00000002));
            g.ReleaseHdc(hdc);
            ReleaseDC(IntPtr.Zero, screenDC);
            return image;
        }

        /// <summary>
        ///     Gets an image that shows the entire tree, not just what is visible on the form
        /// </summary>
        /// <param name="tree"></param>
        private void PrepareTreeImage(TreeView tree)
        {
            this.scrollBarWidth = tree.Width - tree.ClientSize.Width;
            this.scrollBarHeight = tree.Height - tree.ClientSize.Height;
            tree.Nodes[0].EnsureVisible();
            int height = tree.Nodes[0].Bounds.Height;
            this.nodeHeight = height;
            int width = tree.Nodes[0].Bounds.Right;
            TreeNode node = tree.Nodes[0].NextVisibleNode;
            while (node != null)
            {
                height += node.Bounds.Height;
                if (node.Bounds.Right > width)
                {
                    width = node.Bounds.Right;
                }
                node = node.NextVisibleNode;
            }
            //keep track of the original tree settings
            int tempHeight = tree.Height;
            int tempWidth = tree.Width;
            BorderStyle tempBorder = tree.BorderStyle;
            bool tempScrollable = tree.Scrollable;
            TreeNode selectedNode = tree.SelectedNode;
            //setup the tree to take the snapshot
            tree.SelectedNode = null;
            DockStyle tempDock = tree.Dock;
            tree.Height = height + this.scrollBarHeight;
            tree.Width = width + this.scrollBarWidth;
            tree.BorderStyle = BorderStyle.None;
            tree.Dock = DockStyle.None;
            //get the image of the tree

            // .Net 2.0 supports drawing controls onto bitmaps natively now
            // However, the full tree didn't get drawn when I tried it, so I am
            // sticking with the P/Invoke calls
            //_controlImage = new Bitmap(height, width);
            //Bitmap bmp = _controlImage as Bitmap;
            //tree.DrawToBitmap(bmp, tree.Bounds);

            this.controlImage = this.GetImage(tree.Handle, tree.Width, tree.Height);

            //reset the tree to its original settings
            tree.BorderStyle = tempBorder;
            tree.Width = tempWidth;
            tree.Height = tempHeight;
            tree.Dock = tempDock;
            tree.Scrollable = tempScrollable;
            tree.SelectedNode = selectedNode;
            //give the window time to update
            Application.DoEvents();
        }

        private void PrintDocEndPrint(object sender, PrintEventArgs e)
        {
            this.controlImage.Dispose();
        }

        private void PrintDocBeginPrint(object sender, PrintEventArgs e)
        {
            this.lastPrintPosition = new Point(0, 0);
            this.currentDir = PrintDirection.Horizontal;
            this.pageNumber = 0;
        }

        private void PrintDocPrintPage(object sender, PrintPageEventArgs e)
        {
            this.pageNumber++;
            Graphics g = e.Graphics;
            var sourceRect = new Rectangle(this.lastPrintPosition, e.MarginBounds.Size);
            Rectangle destRect = e.MarginBounds;

            if ((sourceRect.Height % this.nodeHeight) > 0)
            {
                sourceRect.Height -= (sourceRect.Height % this.nodeHeight);
            }
            g.DrawImage(this.controlImage, destRect, sourceRect, GraphicsUnit.Pixel);
            //check to see if we need more pages
            if ((this.controlImage.Height - this.scrollBarHeight) > sourceRect.Bottom
                || (this.controlImage.Width - this.scrollBarWidth) > sourceRect.Right)
            {
                //need more pages
                e.HasMorePages = true;
            }
            if (this.currentDir == PrintDirection.Horizontal)
            {
                if (sourceRect.Right < (this.controlImage.Width - this.scrollBarWidth))
                {
                    //still need to print horizontally
                    this.lastPrintPosition.X += (sourceRect.Width + 1);
                }
                else
                {
                    this.lastPrintPosition.X = 0;
                    this.lastPrintPosition.Y += (sourceRect.Height + 1);
                    this.currentDir = PrintDirection.Vertical;
                }
            }
            else if (this.currentDir == PrintDirection.Vertical
                     && sourceRect.Right < (this.controlImage.Width - this.scrollBarWidth))
            {
                this.currentDir = PrintDirection.Horizontal;
                this.lastPrintPosition.X += (sourceRect.Width + 1);
            }
            else
            {
                this.lastPrintPosition.Y += (sourceRect.Height + 1);
            }

            //print footer
            Brush brush = new SolidBrush(Color.Black);
            string footer = this.pageNumber.ToString(NumberFormatInfo.CurrentInfo);
            var f = new Font(FontFamily.GenericSansSerif, 10f);
            SizeF footerSize = g.MeasureString(footer, f);
            var pageBottomCenter = new PointF(x: e.PageBounds.Width / 2, y: e.MarginBounds.Bottom + ((e.PageBounds.Bottom - e.MarginBounds.Bottom) / 2));
            var footerLocation = new PointF(
                pageBottomCenter.X - (footerSize.Width / 2), pageBottomCenter.Y - (footerSize.Height / 2));
            g.DrawString(footer, f, brush, footerLocation);

            //print header
            if (this.pageNumber == 1 && this.title.Length > 0)
            {
                var headerFont = new Font(FontFamily.GenericSansSerif, 24f, FontStyle.Bold, GraphicsUnit.Point);
                SizeF headerSize = g.MeasureString(this.title, headerFont);
                var headerLocation = new PointF(x: e.MarginBounds.Left, y: ((e.MarginBounds.Top - e.PageBounds.Top) / 2) - (headerSize.Height / 2));
                g.DrawString(this.title, headerFont, brush, headerLocation);
            }
        }

        #endregion

        //External function declarations
    }
}