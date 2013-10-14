// ------------------------------------------------------------------------
// <copyright file="Form1.cs" company="Mark Pitman">
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

namespace PrintTreeViewSample
{
    using System;
    using System.Globalization;
    using System.Windows.Forms;

    using Pitman.Printing;

    public partial class Form1 : Form
    {
        #region Constructors and Destructors

        public Form1()
        {
            this.InitializeComponent();
        }

        #endregion

        #region Methods

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            //Add some nodes to the tree
            var tn = this.treeView1.Nodes.Add("Top Level Node");
            tn.ImageIndex = 1;
            const string NodeText = "After three days without programming, life becomes meaningless. ";
            for (int j = 0; j < 10; j++)
            {
                tn = this.treeView1.Nodes[0].Nodes.Add("Child Node " + j.ToString(NumberFormatInfo.CurrentInfo));
                tn.ImageIndex = 1;
                for (int i = 0; i < 3 * j; i++)
                {
                    tn = tn.Nodes.Add(NodeText + i.ToString(NumberFormatInfo.CurrentInfo));
                    tn.ImageIndex = 2;
                }
            }
            this.treeView1.Nodes[0].ExpandAll();
            this.treeView1.Nodes[0].Nodes[1].ExpandAll();
            this.treeView1.Nodes[0].Nodes[2].ExpandAll();
            this.treeView1.Nodes[0].Nodes[5].ExpandAll();
            this.treeView1.Nodes[0].Nodes[9].ExpandAll();
            this.treeView1.SelectedNode = this.treeView1.Nodes[0];
        }

        private void ToolStripButtonPrintPreviewClick(object sender, EventArgs e)
        {
            var print = new PrintHelper();
            print.PrintPreviewTree(this.treeView1, "My Tree Sample");
        }

        #endregion
    }
}