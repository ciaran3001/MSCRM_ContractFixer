using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.ServiceModel.Description;
using Microsoft.Xrm.Sdk.Client;

namespace Ergo.Ticket_CaseFixer
{
    public partial class ADConnect : Form
    {
        string username;
        string password;
        public static IOrganizationService orgService;
        Form1 src;
        OrganizationServiceProxy _organizationService_proxy;


        public ADConnect(Form1 source)
        {
            src = source;
            InitializeComponent();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            //Username
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GetOrgService();
            Form1.SetOrg(_organizationService_proxy);
            this.Dispose();
        }

        public bool GetOrgService()
        {
            string crmurl = "https://crmuat.ergogroup.ie/PreProduction";
            string crmuser = username; //"x15322806@student.ncirl.ie";
            string crmpass = password; //"2104hiHI";
            String apiuri = "/XRMServices/2011/Organization.svc";

            Uri organizationUrl = new Uri(crmurl + apiuri);
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            try
            {
                ClientCredentials credentials = new ClientCredentials();
                credentials.UserName.UserName = textBox1.Text;
                credentials.UserName.Password = textBox2.Text;

                var organizationService_proxy = new OrganizationServiceProxy(organizationUrl, null, credentials, null);
                _organizationService_proxy = organizationService_proxy;
                var organisationService = (IOrganizationService)organizationService_proxy;

                organizationService_proxy.Authenticate();
                var authenticated = organizationService_proxy.IsAuthenticated;

                if (!authenticated)
                {
                    return false;
                }

                orgService = organisationService;
                return true;

            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            //password
        }
    }
}
