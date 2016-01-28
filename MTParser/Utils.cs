using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace MtParser
{
    public static class Utils
    {
        public static Boolean IsEmpty<T>(this IEnumerable<T> source)
        {
            if (source == null)
                return true; // or throw an exception
            return !source.Any();
        }

 public static string RemoveCodeComment(string content){
        var blockComments = @"/\*(.*?)\*/";
        var lineComments = @"'(.*?)\r?\n";
        var strings = @"""((\\[^\n]|[^""\n])*)""";
        var verbatimStrings = @"@(""[^""]*"")+";
        String noComments = Regex.Replace(content, blockComments + "|" + lineComments + "|" + strings + "|" + verbatimStrings,
                                                me => {
                                if (me.Value.StartsWith("/*") || me.Value.StartsWith("'"))
                                    return me.Value.StartsWith("'") ? Environment.NewLine : "";
                                // Keep the literal strings
                                return me.Value;
                            },
                            RegexOptions.Singleline);
        return noComments;

        }
    }

}
