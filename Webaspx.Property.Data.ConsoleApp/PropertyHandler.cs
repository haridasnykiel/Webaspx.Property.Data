using System.Collections.Generic;
using System.Linq;
using Npoi.Mapper;

namespace Webaspx.Property.Data.ConsoleApp
{
    public class PropertyHandler
    {
        
        public static IEnumerable<Model.Property> UpdateTownNames(IEnumerable<RowInfo<Model.Property>> properties) 
        {
            List<Model.Property> propertyList = new List<Model.Property>();

            foreach (var row in properties)
            {
                var property = row.Value;

                property.Town = property.GetTownNameFromAddress();

                propertyList.Add(property);
            }

            return propertyList;
        }

        public static IEnumerable<Model.Property> AddMissingWaxId(
            IEnumerable<Model.Property> propertiesNoWaxId, 
            IEnumerable<Model.Property> propertiesWithWaxId) 
        {
            List<Model.Property> updatedProperties = new List<Model.Property>();

            foreach (var propertyNoWaxId in propertiesNoWaxId)
            {
                var property = propertiesWithWaxId.FirstOrDefault(p => p.Equals(propertyNoWaxId));

                if(property != default(Model.Property)) 
                {
                    updatedProperties.Add(property);
                }
                else 
                {
                    updatedProperties.Add(propertyNoWaxId);
                }
            }

            return updatedProperties;
        }   


    }
}