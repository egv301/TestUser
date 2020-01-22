using System.Data.Linq.Mapping;

namespace L2SApp
{
    [Table(Name = "Users")]
    public class Users
    {
        [Column(IsPrimaryKey = true, IsDbGenerated = true)]
        public int Id { get; set; }
        [Column(Name = "Name")]
        public string Name { get; set; }
        
    }
}