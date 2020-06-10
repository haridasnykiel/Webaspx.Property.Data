using System;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using Npoi.Mapper.Attributes;

namespace Webaspx.Property.Data.ConsoleApp.Model
{
    public class Property : IEquatable<Property>
    {
        [Column("WaxID")]
        public int WaxId { get; set; }
        [Column("PostCode")]
        public string Postcode { get; set; }
        [Column("Streetname")]
        public string StreetName { get; set; }
        [Column("Address")]
        public string Address { get; set; }
        [Column("Town")]
        public string Town { get; set; }

        public string GetTownNameFromAddress() 
        {
            if(!string.IsNullOrEmpty(Town)) 
            {
                return "";
            }

            return Address.Split(",").Last().Trim();
        }

        public bool Equals([AllowNull] Property other)
        {
            if (other == null) return false;

            return
                Postcode == other.Postcode &&
                StreetName == other.StreetName &&
                Address == other.Address;
        }
    }
}