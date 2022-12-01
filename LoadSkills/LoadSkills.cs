using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ObjectiveC;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace LoadSkills
{
    public class LoadSkills
    {
        public static int logCount = 1;
        private static Workbook book;
        private static Worksheet sheet;
        private static Application app;
        static void Main(string[] args)
        {
            Msg("Find file : ", ref logCount, false);
            Msg("Skills.xlsx", true);
            // 파일을 로드할 경로 저장 변수
            StringBuilder targetPath = new StringBuilder();
            // 로드할 파일 명
            string fileName = "Skills.xlsx";
            try
            {
                // 현재 프로그램 폴더 경로 가져오기
                targetPath.Append(Directory.GetCurrentDirectory());
                Msg("Path : ", ref logCount, false);
                Msg(targetPath.ToString(), true);
                // SkillList 이름을 가진 폴더 경로 추가하기
                targetPath = targetPath.Append("\\SkillList\\");
                string savePath = targetPath.ToString();
                Msg("Path : ", ref logCount, false);
                Msg(targetPath.ToString(), true);
                // SkillList 폴더가 존재하지 않는다면 새 폴더 생성
                if (Directory.Exists(targetPath.ToString()) == false)
                {
                    Msg("There's no 'SkillList' folder.", ref logCount, true);
                    Msg("Creating new folder 'SkillList'...>>> ", ref logCount, false);
                    // 폴더 생성
                    Directory.CreateDirectory(targetPath.ToString());
                    // 폴더가 성공적으로 생성되었다면 다음 진행
                    if (Directory.Exists(targetPath.ToString()) == true)
                    {
                        Msg("Success", true);
                    }
                    else
                    {
                        Msg("Failed to create folder. Exit program.", true);
                        return;
                    }
                }
                // 타겟 경로에 파일 명까지 추가하기
                targetPath.Append(fileName);
                Msg(targetPath.ToString(), ref logCount, true);
                // 읽어 올 파일이 존재하지 않는다면 프로그램 종료
                if (File.Exists(targetPath.ToString()) == false)
                {
                    Msg("There's no file ", ref logCount, false);
                    Msg(fileName, true);
                    return;
                }
                Msg("Opening excel file...", ref logCount, true);
                // 파일을 찾았다면, 엑셀 파일을 연다.
                app = new Application();
                book = app.Workbooks.Open(targetPath.ToString());
                sheet = book.Worksheets.Item[1] as Worksheet;
                // 만약 엑셀 열기에 실패하였다면, 프로그램 종료
                if (book == null || sheet == null || app == null)
                {
                    Msg("Error : Failed to open excel file or worksheet.", ref logCount, true);
                    return;
                }
                Msg("Success", ref logCount, true);
                // 엑셀에서 데이터가 있는 셀 범위를 가져온다.
                Excel.Range range = sheet.UsedRange;
                // 엑셀에서 읽어 온 데이터를 저장할 2차원 배열 생성
                string[,] data = new string[range.Rows.Count - 1, range.Columns.Count];

                Msg("Reading excel file...", ref logCount, true);
                Msg("------------------------", ref logCount, true);
                // 하나의 열을 기준으로 행의 데이터를 순차적으로 읽음
                for (int column = 1; column <= range.Columns.Count; column++)
                {
                    // 첫 행 (필드 이름) 은 무시함
                    for (int row = 2; row <= range.Rows.Count; row++)
                    {
                        data[row - 2, column - 1] = (string)(range.Cells[row, column] as Excel.Range).Value2;
                        Msg("Read Data : ", ref logCount, false);
                        Msg(data[row - 2, column - 1], true);
                    }
                    Msg("------------------------", ref logCount, true);
                }
                Msg("Quit excel program...", ref logCount, true);
                // 사용한 엑셀 파일을 올바르게 닫고 종료시킨다.
                book.Close();
                app.Quit();
                // C++ 스크립트 작성을 위한 함수를 호출하며,
                WriteScript writeScript = new WriteScript();
                // 저장할 경로와 파일명, 저장된 Data 를 넘겨준다.
                bool result = writeScript.WriteFile(savePath, "SkillData.h", data);
                // 파일 작성에 실패하면 종료
                if (result == false)
                {
                    Msg("Failed to writing file.", ref logCount, true);
                }
            }
            finally
            {
                Msg("Exiting program...", ref logCount, true);
                try
                {
                    if (sheet != null)
                    {
                        Marshal.ReleaseComObject(sheet);
                    }
                    if (book != null)
                    {
                        Marshal.ReleaseComObject(sheet);
                    }
                    if (app != null)
                    {
                        Marshal.ReleaseComObject(app);
                    }
                }
                finally
                {
                    Msg("Exit", ref logCount, true);
                    GC.Collect();
                }
            }
        }
        /// <summary>
        /// 로그 카운트없이 메세지를 출력한다.
        /// </summary>
        /// <param name="message"></param>
        /// <param name="Enter"></param>
        public static void Msg(string message, bool Enter)
        {
            if (Enter)
            {
                Console.WriteLine("{0}", message);
            }
            else
            {
                Console.Write("{0}", message);
            }
        }
        /// <summary>
        /// 로그 카운트와 개행을 포함하여 메세지를 출력한다.
        /// </summary>
        /// <param name="message">출력 메세지</param>
        /// <param name="logCount">로그 카운트</param>
        /// <param name="Enter">개행 여부</param>
        public static void Msg(string message, ref int logCount, bool Enter)
        {
            if (Enter)
            {
                Console.WriteLine("[{0}] {1}", logCount, message);
                logCount++;
            }
            else
            {
                Console.Write("[{0}] {1}", logCount, message);
                logCount++;
            }
        }
    }
}
