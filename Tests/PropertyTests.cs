using NUnit.Framework;
using Webaspx.Property.Data.ConsoleApp.Model;

namespace Tests
{
    public class PropertyTests
    {
        [Test]
        public void GetTownNames_OneComma() 
        {
            var property = new Property 
            {
                WaxId = 1,
                Postcode = "WD32BQ",
                StreetName = "Hays Close",
                Address = "12 Hays Close, Hay Town"      
            };
            var expected = "Hay Town";

            var actual = property.GetTownNameFromAddress();

            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void GetTownNames_TwoCommas()
        {
            var property = new Property
            {
                WaxId = 1,
                Postcode = "WD32BQ",
                StreetName = "Hays Close",
                Address = "12 Hays Close, something, Hay Town"
            };
            var expected = "Hay Town";

            var actual = property.GetTownNameFromAddress();

            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void GetTownNames_NoTownNameInAddress()
        {
            var property = new Property
            {
                WaxId = 1,
                Postcode = "WD32BQ",
                StreetName = "Hays Close",
                Address = "12 Hays Close",
                Town = "Hay Town"
            };
            var expected = "";

            var actual = property.GetTownNameFromAddress();

            Assert.AreEqual(expected, actual);
        }
    }
}