using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;

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
    class WordToJson
    {
        List<string> source_files;
        Microsoft.Office.Interop.Word.Application _ApplicationWord;
        Microsoft.Office.Interop.Word.Document _MonDocument;
        Dictionary<string, List<string>> content;

        public WordToJson(List<string> p_source_files)
        {
            source_files = p_source_files;
        }

        /// <summary>
        /// Main of the application
        /// </summary>
        public void Run()
        {
            Console.WriteLine("Extracting ContentControls to JSON...");
            // Open Microsoft Office Word
            try
            {
                _ApplicationWord = new Microsoft.Office.Interop.Word.Application();
                _ApplicationWord.Visible = false;
            }
            catch
            {
                Console.WriteLine("[ERROR] code 2: Can't open Microsoft Office Word.");
                Environment.Exit(2);
            }
            // Process each file            
            foreach (string source_file in this.source_files)
            {
                // STEP 1: Import data from ContentControls
                bool import_ok = this.ImportContentControlsFromWord(source_file);
                if (!import_ok)
                {
                    _ApplicationWord.Quit();
                    Console.WriteLine("[ERROR] code 3: Can't import ContentControls from \"" + source_file + "\"");
                    Environment.Exit(3);
                }
                // STEP 2: Export data to JSON
                bool export_ok = ExportToJson(source_file);
                if (!export_ok)
                {
                    _ApplicationWord.Quit();
                    Console.WriteLine("[ERROR] code 4: Can't export ContentControls to JSON for \"" + source_file + "\" . Check write access to myFile.json");
                    Environment.Exit(4);
                }
                Console.WriteLine("[NOTE] Json done for \"" + source_file + "\"");
            }
            _ApplicationWord.Quit();
        }

        /// <summary>
        /// Open Word file and import data from ContentControls
        /// </summary>
        /// <param name="p_source_file">Word file used as source for data</param>
        /// <returns></returns>
        private bool ImportContentControlsFromWord(string p_source_file)
        {
            bool import_ok;
            try
            {
                // Open file in Microsoft Office Word
                this._MonDocument = _ApplicationWord.Documents.Open(p_source_file, ReadOnly: true);
                // Extract data
                List<ContentControl> cc = new List<ContentControl>();
                cc = GetAllContentControls(_MonDocument);

                // Build dictonary from ContentControls
                // Note: Each ContentControls can more than once => which explains the list of values
                this.content = new Dictionary<string, List<string>>();
                foreach (var i in cc)
                {
                    string title = i.Title;
                    string value = i.Range.Text;
                    // Check for key existence
                    if (!this.content.ContainsKey(title))
                    {
                        this.content.Add(title, new List<string>());
                    }
                    // Format and Push value
                    value = value.Replace("\\", ""); // \ is escape char in JSON
                    value = value.Replace("\"", "'"); // " is special char in JSON
                    this.content[i.Title].Add(value);
                }
                // Close file
                _MonDocument.Close();
                import_ok = true;
            }
            catch
            {
                import_ok = false;
            }
            return import_ok;
        }

        /// <summary>
        /// Facility to extract data from ContentControls
        /// </summary>
        /// <param name="wordDocument"></param>
        /// <returns></returns>
        private static List<ContentControl> GetAllContentControls(Document wordDocument)
        {
            if (null == wordDocument)
                throw new ArgumentNullException("wordDocument");

            List<ContentControl> ccList = new List<ContentControl>();

            Range rangeStory;
            foreach (Range range in wordDocument.StoryRanges)
            {
                rangeStory = range;
                do
                {
                    try
                    {
                        foreach (ContentControl cc in rangeStory.ContentControls)
                        {
                            ccList.Add(cc);
                        }

                        foreach (Shape shapeRange in rangeStory.ShapeRange)
                        {
                            foreach (ContentControl cc in shapeRange.TextFrame.TextRange.ContentControls)
                            {
                                ccList.Add(cc);
                            }
                        }
                    }
                    catch (COMException) { }
                    rangeStory = rangeStory.NextStoryRange;

                }
                while (rangeStory != null);
            }
            return ccList;
        }

        /// <summary>
        /// Create JSON (without modern library to avoid portability to old school computer)
        /// </summary>
        /// <param name="p_source_file">Word file used as source for data</param>
        /// <returns></returns>
        private bool ExportToJson(string p_source_file)
        {
            bool export_ok = false;
            // 
            using (StreamWriter f = new StreamWriter(p_source_file + ".json"))
            {
                f.WriteLine("{");
                int key_index = 1;
                string indentation = "  ";

                f.WriteLine(indentation + "\"path\": \"" + p_source_file.Replace("\\", "\\\\") + ".json" + "\"" + (content.Count > 0 ? "," : ""));

                foreach (var k in content.Keys)
                {
                    bool last_key = key_index == content.Keys.Count;
                    if (content[k].Count == 0)
                    {
                        // Empty
                        f.WriteLine(indentation + "\"" + k + "\": \"\"" + (last_key ? "" : ","));
                    }
                    if (content[k].Count == 1)
                    {
                        // Single entry
                        f.WriteLine(indentation + "\"" + k + "\": \"" + content[k].First() + "\"" + (last_key ? "" : ","));
                    }
                    if (content[k].Count > 1)
                    {
                        // Multiple entries
                        f.WriteLine(indentation + "\"" + k + "\": [");
                        int value_index = 1;
                        foreach (var v in content[k])
                        {
                            bool last_value = value_index == content[k].Count;
                            f.WriteLine(indentation + indentation + "\"" + v + "\"" + (last_value ? "" : ","));
                        }
                        f.WriteLine(indentation + indentation + "]" + (last_key ? "" : ","));
                    }
                    key_index++;
                }
                f.WriteLine("}");
                export_ok = true;
            }
            return export_ok;
        }


    }
}
