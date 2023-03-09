using System.Text;

namespace ImportExportAPI.Utils;

public static class StringUtils
{
    public static String sliceString(String myString)
    {
        string[] words = myString.Split(' ');
        StringBuilder sb = new StringBuilder();
        int maxLineLength = 16;
        int currLength = 0;
        foreach (string word in words)
        {
            if (currLength + word.Length + 1 < maxLineLength) // +1 accounts for adding a space
            {
                sb.AppendFormat(" {0}", word);
                currLength = (sb.Length + maxLineLength);
            }
            else
            {
                sb.AppendFormat("{0}{1}", Environment.NewLine, word);
                currLength = 0;
            }
        }

        return sb.ToString();
    }

    public static String sliceStringInEverySpace(String myString)
    {
        string[] words = myString.Split(' ');
        StringBuilder sb = new StringBuilder();
        foreach (string word in words)
        {
            sb.AppendFormat("{0}{1}", Environment.NewLine, word);
        }
        sb.AppendFormat("{0}", Environment.NewLine);

        return sb.ToString();
    }
}
