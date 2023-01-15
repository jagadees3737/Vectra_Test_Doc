using System;
using INFITF;
using MECMOD;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using DRAFTINGITF;
using ProductStructureTypeLib;
namespace Vectra_test
{
    //This class is used to display the output in the Windows forms.
    public partial class Form1 : Form
    {
        static INFITF.Application myCatia;
        public Form1()
        {
            InitializeComponent();
        }

        private void Btn_Drawingview_Click(object sender, EventArgs e)
        {
            Scope_2 scope2 = new Scope_2();
            scope2.startCatia();
        }

        /// <summary>
        /// This Method is used get the Assembly tree from CATIA and display the values in Tree list view in Windows forms.
        /// For this an active product document should be priorly opened in the CATIA.
        /// Then looping throught the assembly tree structure and get the values to display the values in Treelist view.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Btn_Treeview_Click(object sender, EventArgs e)
        {
            //To Initiate or start the CATIA application.
            try
            {
                myCatia = (INFITF.Application)Marshal.GetActiveObject("CATIA.Application");
            }
            catch (Exception)
            {
                myCatia = (INFITF.Application)Activator.CreateInstance(System.Type.GetTypeFromProgID("CATIA.Application"));
            }

            //To get the Active Product Document from CATIA.
            ProductDocument activeDocument = (ProductDocument)Form1.myCatia.ActiveDocument;
            Product root_product = activeDocument.Product;
            Products products = root_product.Products;

            //To Setup the Tree Structure in Windows forms.
            TreeNode node = new TreeNode(root_product.get_Name());
            this.treeView1.Nodes.Add(node);

            //To get the Child and subchild values from the assembly tree and display in Windows forms.
            for (int child_count = 1; child_count <= products.Count; child_count++)
            {
                Products child_products = products;
                Product child_product = child_products.Item(child_count);
                node.Nodes.Add(child_product.get_Name());
                foreach (Product subchild in child_product.Products)
                {
                    node.LastNode.Nodes.Add(subchild.get_Name());
                }
            }
            this.treeView1.ExpandAll();
        }
    }
}
