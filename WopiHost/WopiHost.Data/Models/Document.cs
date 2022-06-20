using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace WopiHost.Data.Models
{
    public partial class Document
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int DocumentId { get; set; }
        public Guid DocumentGuid { get; set; }
        public string DocumentName { get; set; }
        public string Extension { get; set; }
        //public string FileName { get; set; }
        public string Blob { get; set; }
        public DateTime? CreatedDate { get; set; } = DateTime.UtcNow;
    }
}
