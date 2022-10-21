

using System.ComponentModel.DataAnnotations.Schema;

namespace ExposureTracker.Models
{
 
    
    public class Insured
    {
        [Key]
        public int id { get; set; }
        [Required]

        public string identifier { get; set; }

        public string? policyno { get; set; }
 
        public string? firstname { get; set; }

        public string? middlename { get; set; }

        public string? lastname { get; set; }

        public string? fullName { get; set; }

        public string? gender { get; set; }

        public string? clientid { get; set; }

        public string dateofbirth { get; set; }

        public string? cedingcompany { get; set; }

        public string? cedantcode { get; set; }

        public string? typeofbusiness { get; set; }

        public string? bordereauxfilename { get; set; }
        public int? bordereauxyear { get; set; }

        public string? soaperiod { get; set; }
        public string? certificate { get; set; }

        public string? plan { get; set; }

        public string? benefittype { get; set; }

        public string? baserider { get; set; }

        public string? currency { get; set; }
        public string planeffectivedate { get; set; }
        public Decimal sumassured { get; set; }

        public Decimal reinsurednetamountatrisk { get; set; }

        public string? mortalityrating { get; set; }

        public string? status { get; set; }

        public string? dateuploaded { get; set; }

        public string? uploadedby { get; set; }
    }

}
