using Ergo.Ticket_CaseFixer.Models;
using Ergo.Ticket_CaseFixer.Services;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Ergo.Ticket_CaseFixer
{
    public partial class Form1 : Form
    {
        static IOrganizationService _orgService;

        public static bool OrgServiceRetreived = false;
        CRMService crmService = new CRMService();

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!OrgServiceRetreived)
            {
                ADConnect ad = new ADConnect(this);
                ad.Visible = true;


            }
        }

        public static void SetOrg(OrganizationServiceProxy org)
        {
            _orgService =(IOrganizationService) org;
            label3.Text = "Authenticated: " +  org.IsAuthenticated.ToString();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                Guid contractId = new Guid(textBox2.Text);

                QueryExpression query = new QueryExpression("incident");
                query.ColumnSet = new ColumnSet("incidentid", "contractid", "ergo_contractlineid", "ergo_contracttechnology", "ticketnumber", "statecode");
                query.Criteria.AddCondition("statecode", ConditionOperator.Equal, "0");
                query.Criteria.AddCondition("contractid", ConditionOperator.Equal, contractId);

                EntityCollection brokenCases = _orgService.RetrieveMultiple(query);

                List<Case> ReturnedCases = new List<Case>();
                foreach(var _case in brokenCases.Entities)
                {
                    Case cse = new Models.Case();
                    cse.id = _case.GetAttributeValue<Guid>("incidentid");
                    cse.Contract = _case.GetAttributeValue<EntityReference>("contractid");
                    cse.ContractLine = _case.GetAttributeValue<EntityReference>("ergo_contractlineid");
                    cse.ContractTech = _case.GetAttributeValue<EntityReference>("ergo_contracttechnology");
                    cse.RefNumber = _case.GetAttributeValue<string>("ticketnumber");

                    cse.RecordFixed = false;
                    cse.Selected = false;
                    ReturnedCases.Add(cse);
                }

                //Get broken cases for selected customer and contract
                var table = ConvertListToDataTable(ReturnedCases);
                dataGridView1.DataSource = table;
                dataGridView1.Columns[0] = "selected";

                    }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error occured",
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Question);

            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            //contract
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        static DataTable ConvertListToDataTable(List<String> list)
        {
            // New table.
            DataTable table = new DataTable();

            // Get max columns.
            int columns = 0;
            foreach (var array in list)
            {
                if (array.Length > columns)
                {
                    columns = array.Length;
                }
            }

            // Add columns.
            for (int i = 0; i < columns; i++)
            {
                table.Columns.Add();
            }

            // Add rows.
            foreach (var array in list)
            {
                table.Rows.Add(array);
            }

            return table;
        }
    }
}
