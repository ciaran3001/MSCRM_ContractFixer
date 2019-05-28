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

        public Case(Guid _id, EntityReference _contr, EntityReference _line, EntityReference _tech, string _ref)
        {
            id = _id;
            Contract = _contr;
            ContractLine = _line;
            ContractTech = _tech;
            RefNumber = _ref;

            Selected = false;
            RecordFixed = false;

        }

        public Guid GetID()
        {
            return id;

        }

        public EntityReference GetContract()
        {
            return Contract;
        }

        public EntityReference GetContractLine()
        {
            return ContractLine;
        }

        public EntityReference GetContractTech()
        {
            return ContractTech;
        }
        public string GetRef()
        {
            return RefNumber;
        }

        public bool GetSelected()
        {
            return Selected;
        }
        public bool GetFixed()
        {
            return RecordFixed;
        }
    }
}
