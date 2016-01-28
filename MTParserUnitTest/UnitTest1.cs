using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Collections.Generic;
using System.Linq;
namespace MTParserUnitTest
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestRemoveCodeComment()
        {
            string content = string.Empty;
            string filePath = @"D:\svn\WIP\Sources\trunk\PWI 6.5\Security\MT\clsDataXange.vb";
            content = "ffdsfsdf\"dsad'as\"";

                using (var sr = new StreamReader(filePath))
                {
                    content =  sr.ReadToEnd();
                }

            var lines = content.Split(new[] { "\r\n" }, StringSplitOptions.None).ToList();
            
            string nocomment = MtParser.Utils.RemoveCodeComment(content);
            var lines2 = nocomment.Split(new[] { "\r\n" }, StringSplitOptions.None).ToList();

            //Assert.AreEqual(lines.Count, lines2.Count, "The length not equals");
            Assert.AreEqual("ffdsfsdf\"dsad'as\"", nocomment, "not pass");
        }
    }
}
