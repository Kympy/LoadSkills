using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LoadSkills
{
    internal class WriteScript
    {
        public bool WriteFile(string savePath, string fileName, string[,] data)
        {
            try
            {
                StringBuilder path = new StringBuilder();
                // 저장 경로 설정 (SkillList 폴더까지)
                path.Append(savePath);
                // 경로에 파일명 추가
                path.Append(fileName);

                LoadSkills.Msg("Saving path : ", ref LoadSkills.logCount, false);
                LoadSkills.Msg(path.ToString(), true);
                // 파일 작성 시작
                StreamWriter writer = File.CreateText(path.ToString());
                LoadSkills.Msg("Write --", ref LoadSkills.logCount, true);
                // 전처리기 작성
                LoadSkills.Msg("Writing Preprocessor...", ref LoadSkills.logCount, true);
                writer.WriteLine("#pragma once");
                writer.WriteLine("#include <iostream>");
                writer.WriteLine("#include <map>");
                writer.WriteLine("using namespace std;");
                writer.WriteLine();
                LoadSkills.Msg("Writing enum class...", ref LoadSkills.logCount, true);
                writer.WriteLine("const enum class SkillType {");
                writer.Write("\t");

                StringBuilder text = new StringBuilder();
                int lineSpan = 1;

                for (int row = 0; row < data.GetLength(0); row++)
                {
                    text.Append(data[row, 0]);
                    text.Append(", ");
                    writer.Write(text.ToString());
                    text.Clear();

                    lineSpan++;
                    if (lineSpan >= 5)
                    {
                        writer.WriteLine();
                        writer.Write("\t");
                        lineSpan = 1;
                    }
                }
                writer.WriteLine();
                writer.WriteLine("};");
                writer.WriteLine();
                LoadSkills.Msg("Writing map...", ref LoadSkills.logCount, true);
                writer.WriteLine("const map<SkillType, string> SkillDictionary = {");

                for (int row = 0; row < data.GetLength(0); row++)
                {
                    text.Append("\t{ SkillType::");
                    text.Append(data[row, 0]);
                    text.Append(", ");
                    text.Append("\"");
                    text.Append(data[row, 1]);
                    text.Append("\" },");
                    writer.WriteLine(text.ToString());
                    text.Clear();
                }
                writer.WriteLine("};");

                writer.WriteLine();
                writer.Write("const int maxCount = ");
                writer.Write(data.GetLength(0).ToString());
                writer.Write(";");
                writer.Close();
                LoadSkills.Msg("Finished writing.", ref LoadSkills.logCount, true);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
