using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace ReadExcelFile.Models
{
    public class Movies
    {
        [Required]
        public String MovieName { get; set; }
        [Required]
        public String Hero { get; set; }
        [Required]
        public String Director { get; set; }
    }
}