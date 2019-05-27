using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ergo.Ticket_CaseFixer.Models
{
    class Case
    {
        public Guid id;
        public EntityReference Contract;
        public EntityReference ContractLine;
        public EntityReference ContractTech;

        public string RefNumber;
        public bool Selected = false;
        public bool RecordFixed = false;

    }
}
