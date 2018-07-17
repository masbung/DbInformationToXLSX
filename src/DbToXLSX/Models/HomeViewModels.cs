using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DbToXLSX.Models
{
    public class HomeViewModel
    {
        public HomeViewModel() {
            this.Tables = new List<Table>();
        }

        public List<Table> Tables { get; set; }
    }

    public class Table
    {
        public string TableName { get; set; }
        public string SchemaName { get; set; }
        public bool Print { get; set; }
    }

    public class TableViewModel
    {
        public List<Table> Tables { get; set; }
    }

    public class SchemaViewModel
    {
        public SchemaViewModel() {
            this.Schemas = new List<Schema>();
        }
        public List<Schema> Schemas { get; set; }
    }

    public class Schema
    {
        public Schema() {
            this.Tables = new List<Table>();
        }

        public string Name { get; set; }
        public List<Table> Tables { get; set; }
    }    
}