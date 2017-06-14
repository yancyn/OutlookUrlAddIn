using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OutlookUrlAddIn;
using Xunit;

namespace OutlookUrlAddIn.Test
{
    public class Utils
    {
        public Utils()
        {

        }
        [Fact]
        public void ExtractValidUrlTest()
        {
            string target = "Welcome to RegExr v2.1 by gskinner.com, proudly hosted by Media Temple!";
            target += "\nEdit the Expression & Text to see matches. Roll over matches or the expression for details. Undo mistakes with ctrl-z. Save Favorites & Share expressions with friends or the Community. Explore your results with Tools. A full Reference & Help is available in the Library, or watch the video Tutorial.";
            target += "\n\nSample text for testing:";
            target += "\nabcdefghijklmnopqrstuvwxyz ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            target += "\n12345 -98.7 3.141 .6180 9,000 +42";
            target += "\n555.123.4567	+1-(800)-555-2468";
            target += "\nfoo@demo.net	bar.ba@test.co.uk";
            target += "\nwww.demo.com	http://foo.co.uk/";
            target += "\nhttp://regexr.com/foo.html?q=bar";
            target += "\nhttps://mediatemple.net";

            string[] expected = new string[] { "www.demo.com", "http://foo.co.uk/", "http://regexr.com/foo.html?q=bar", "https://mediatemple.net" };
            string[] actual = OutlookUrlAddIn.Utils.ExtractUrl(target);
            Assert.Equal(expected, actual);
        }
        [Fact]
        public void ExtractNoneUrlTest()
        {
            string target = "Welcome to RegExr v2.1 by gskinner.com, proudly hosted by Media Temple!";
            target += "\nEdit the Expression & Text to see matches. Roll over matches or the expression for details. Undo mistakes with ctrl-z. Save Favorites & Share expressions with friends or the Community. Explore your results with Tools. A full Reference & Help is available in the Library, or watch the video Tutorial.";
            target += "\n\nSample text for testing:";
            target += "\nabcdefghijklmnopqrstuvwxyz ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            target += "\n12345 -98.7 3.141 .6180 9,000 +42";
            target += "\n555.123.4567	+1-(800)-555-2468";
            target += "\nfoo@demo.net	bar.ba@test.co.uk";

            string[] expected = new string[] { };
            string[] actual = OutlookUrlAddIn.Utils.ExtractUrl(target);
            Assert.Equal(expected, actual);
        }
    }
}