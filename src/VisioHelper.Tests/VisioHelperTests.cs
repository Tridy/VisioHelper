global using Xunit;
using System.Threading.Tasks.Sources;

namespace Helpers.Tests
{
    public class VisioHelperTests
    {
        [Fact]
        public void CanGetShapeTextsWithPrefix()
        {
            string filePath = Path.Combine("visio_file", "visio001.vsdx");

            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException(filePath);
            }

            VisioHelper helper = new();
            var resultTexts = helper.SearchForPrefixTextInVisioFile(filePath, "Note");

            Assert.Equal("MyPageA", resultTexts.ElementAt(0).Key);
            Assert.Equal("MyPageB", resultTexts.ElementAt(1).Key);

            var pageAItems = resultTexts.ElementAt(0).Value;
            Assert.Equal(3, pageAItems.Count());
            Assert.Contains("Note one", pageAItems);
            Assert.Contains("Note two", pageAItems);
            Assert.Contains("Note three", pageAItems);

            var pageBItems = resultTexts.ElementAt(1).Value;
            Assert.Equal(2, pageBItems.Count());
            Assert.Contains("Note four", pageBItems);
            Assert.Contains("Note five", pageBItems);
        }
    }
}