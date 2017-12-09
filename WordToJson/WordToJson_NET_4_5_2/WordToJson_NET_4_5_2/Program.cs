using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordToJson_NET_4_5_2;

/*
Copyright 2017 Rafael CATROU

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/

namespace WordToJson_NET_4_5_2
{
    class Program
    {
        static int Main(string[] args)
        {
            // Control ARGS
            if (args.Count() < 1)
            {
                Console.WriteLine("[ERROR] code 1: missing the path(s) of file(s) to process");
                Environment.Exit(1);
            }
            // Do
            WordToJson application = new WordToJson(args.ToList<string>());
            application.Run();
            return 0;
        }
    }
}
