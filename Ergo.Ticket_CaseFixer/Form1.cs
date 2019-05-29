using Ergo.Ticket_CaseFixer.Models;
using Ergo.Ticket_CaseFixer.Services;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
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
        EntityCollection returned_cases;

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
                    query.ColumnSet = new ColumnSet("incidentid", "ergo_contractid", "ergo_contractlineid", "ergo_contracttechnology", "ticketnumber", "statecode", "statuscode");
                    query.Criteria.AddCondition("statuscode", ConditionOperator.In, 1,2, 951350002);
                    query.Criteria.AddCondition("ergo_contractid", ConditionOperator.Equal, contractId);

                    EntityCollection brokenCases = _orgService.RetrieveMultiple(query);
                    returned_cases = brokenCases;
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
                for (var i = 0; i < selectedCellCount * 7; i += 7)
                {
                    //MessageBox.Show(dataGridView1.SelectedCells[i + 6].Value.ToString(), "Selected Records", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    CasesToFix.Add(new Guid(dataGridView1.SelectedCells[i + 6].Value.ToString()));
                }

                foreach(var _case in CasesToFix)
                {
                    Entity target;
                   
                    foreach (var entity in returned_cases.Entities)
                    {
                        if (_case == entity.Id)
                        {
                            target = entity;
                            var statusreason = entity.GetAttributeValue<OptionSetValue>("statuscode");
                            var state = entity.GetAttributeValue<OptionSetValue>("statecode");
                            Guid new_contract = new Guid(textBox1.Text);
                            List<EntityReference> contractLines = GetNewContractLines(new_contract);
                            foreach(var line in contractLines)
                            {
                                if (entity.GetAttributeValue<EntityReference>("ergo_contractlineid").Name.Equals(line.Name))
                                {
                                    entity["ergo_contractlineid"] = line;

                                    List<EntityReference> technologies = GetNewTech(line.Id);
                                    foreach (var tech in technologies)
                                    {
                                        if (entity.GetAttributeValue<EntityReference>("ergo_contracttechnology").Name.Equals(tech.Name))
                                        {
                                            entity["ergo_contracttechnology"] = tech;
                                        }
                                    }
                                }
                             
                            }
                            //status
                            //status reason
                            //owner
                            EntityReference contract = entity.GetAttributeValue<EntityReference>("ergo_contractid");
                            contract.Id = new Guid(textBox1.Text);
                            entity["ergo_contractid"] = contract;
                           // entity["statuscode"] = statusreason;

                           
                            //entity["statecode"] = state;

                            _orgService.Update(entity);

                            SetStateRequest request = new SetStateRequest();
                            request.EntityMoniker = entity.ToEntityReference();
                            request.State = state;
                            request.Status = statusreason;
                            _orgService.Execute(request);



                            // _orgService.Update(entity);              
                        }
                    }
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

        public List<EntityReference> GetNewContractLines(Guid newContract)
        {
            List<EntityReference> tmp = new List<EntityReference>();

            QueryExpression query = new QueryExpression("ergo_contractline");
            query.Criteria.AddCondition("ergo_contractid", ConditionOperator.Equal, newContract);

            var response = _orgService.RetrieveMultiple(query);
            foreach(var line in response.Entities)
            {
                tmp.Add(line.ToEntityReference());
            }
            return tmp;
        }

        public List<EntityReference> GetNewTech(Guid line)
        {
            List<EntityReference> tmp = new List<EntityReference>();

            QueryExpression query = new QueryExpression("ergo_contracttechnology");
            query.Criteria.AddCondition("ergo_contractline", ConditionOperator.Equal, line);

            var response = _orgService.RetrieveMultiple(query);
            foreach (var tech in response.Entities)
            {
                tmp.Add(tech.ToEntityReference());
            }
            return tmp;
        }
    }
}
