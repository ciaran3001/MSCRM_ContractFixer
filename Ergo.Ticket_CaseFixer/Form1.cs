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
        static int trace = 0;
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
            }else
            {
                try
                {
                    Guid contractId = new Guid(textBox2.Text);

                    QueryExpression query = new QueryExpression("incident");
                    query.ColumnSet = new ColumnSet("incidentid", "ergo_contractid", "ergo_contractlineid", "ergo_contracttechnology", "ticketnumber", "statecode");
                    // query.Criteria.AddCondition("statecode", ConditionOperator.Equal, "0");
                    query.Criteria.AddCondition("ergo_contractid", ConditionOperator.Equal, contractId);

                    EntityCollection brokenCases = _orgService.RetrieveMultiple(query);

                    List<Case> ReturnedCases = new List<Case>();
                    foreach (var _case in brokenCases.Entities)
                    {
                        Case cse = new Models.Case(
                            _case.GetAttributeValue<Guid>("incidentid"),
                            _case.GetAttributeValue<EntityReference>("ergo_contractid"),
                            _case.GetAttributeValue<EntityReference>("ergo_contractlineid"),
                            _case.GetAttributeValue<EntityReference>("ergo_contracttechnology"),
                            _case.GetAttributeValue<string>("ticketnumber")
                            );

                        // cse.id = 
                        //cse.Contract = _case.GetAttributeValue<EntityReference>("contractid");
                        //  cse.ContractLine = _case.GetAttributeValue<EntityReference>("ergo_contractlineid");
                        //  cse.ContractTech = _case.GetAttributeValue<EntityReference>("ergo_contracttechnology");
                        //   cse.RefNumber = _case.GetAttributeValue<string>("ticketnumber");

                        //   cse.RecordFixed = false;
                        //  cse.Selected = false;
                        ReturnedCases.Add(cse);
                    }

                    //Get broken cases for selected customer and contract
                    var table = ConvertListToDataTable(ReturnedCases);
                    dataGridView1.DataSource = table;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error occured",
                                    MessageBoxButtons.YesNo,
                                    MessageBoxIcon.Question);

                }
            }
        }

        public static void SetOrg(OrganizationServiceProxy org)
        {
            _orgService = (IOrganizationService)org;
            label3.Text = "Authenticated: " + org.IsAuthenticated.ToString();
            OrgServiceRetreived = true;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                List<Guid> CasesToFix = new List<Guid>();

                //fix selected records
                Int32 selectedCellCount = (dataGridView1.GetCellCount(DataGridViewElementStates.Selected)) / 7;
                var message = MessageBox.Show(selectedCellCount.ToString() + " Records selected", "Selected Records", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                /*  foreach(var row in dataGridView1.SelectedCells)
                  {
                      CasesToFix.Add(new Guid(row.))
                  }*/
                for (var i = 0; i < selectedCellCount * 7; i += 7)
                {
                    MessageBox.Show(dataGridView1.SelectedCells[i + 6].Value.ToString(), "Selected Records", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    CasesToFix.Add(new Guid(dataGridView1.SelectedCells[i + 6].Value.ToString()));


                }
            }catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Selected Records", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            //contract
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        static DataTable ConvertListToDataTable(List<Case> list)
        {

            // New table.
            DataTable table = new DataTable();

            // Get max columns.
            int columns = 7;
            /*      foreach (var array in list)
                  {
                      if (array.Length > columns)
                      {
                          columns = array.Length;
                      }
                  }*/


            table.Columns.Add("selected");
            table.Columns.Add("CRM Ref");
            table.Columns.Add("Contract ID");
            table.Columns.Add("ContractLine");
            table.Columns.Add("ContractTech");
            table.Columns.Add("Fixed");
            table.Columns.Add("Id");
            // Add rows.
            int e = 0;
            EntityReference empty = new EntityReference("EMPTY", "EMPTY", new Guid("00000000-0000-0000-0000-000000000000"));

            foreach (var array in list)
            {
                try
                {


                    if (array.RefNumber == null) array.RefNumber = "EMPTY";
                    if (array.Contract.Name == null) array.Contract = empty;
                    if (array.ContractLine.Name == null) array.ContractLine = empty;
                    if (array.ContractTech.Name == null) array.ContractTech = empty;
                    if (array.id == null) continue;

                    table.Rows.Add(array);
                    table.Rows[e][0] = array.GetSelected();
                    table.Rows[e][1] = array.GetRef();
                    table.Rows[e][2] = array.GetContract().Name.ToString();
                    table.Rows[e][3] = array.GetContractLine().Name.ToString();
                    table.Rows[e][4] = array.GetContractTech().Name.ToString();
                    table.Rows[e][5] = array.GetFixed();
                    table.Rows[e][6] = array.GetID().ToString();
                    trace++;
                    e++;
                }
                catch (Exception ex)
                {
                     var message = MessageBox.Show(trace.ToString() + " : " + ex.Message, "Error occured",
                     MessageBoxButtons.YesNo,
                      MessageBoxIcon.Question);


                    continue;
                }
            }
            return table;
        }
    }
}
