using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using Virinco.WATS.Interface;

namespace SigmaQuest
{
    [TestClass]
    public class ConverterTests : TDM
    {

        [TestMethod]
        public void ConverterTest()
        {
            InitializeAPI(true);
            string fn = @"Data\test.xml";
            var converter = new GenericXMLConverter();
            using (FileStream file = new FileStream(fn, FileMode.Open))
            {
                SetConversionSource(new FileInfo(fn), converter.ConverterParameters, null);
                Report uut = converter.ImportReport(this, file);
                Submit(uut);
            }
        }
    }
}
