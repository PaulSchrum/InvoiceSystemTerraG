using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TimeAnalyzerino
{
   public static class Util
   {
      public static bool AreAnyNull(params Object[] p)
      {
         return p.Any(param => null == param);
      }

   }
}
