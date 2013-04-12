using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SpellCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            SpellCheck spellCheck = new SpellCheck();
            spellCheck.Run(args);
        }
    }
}
