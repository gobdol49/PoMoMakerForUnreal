using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
//using System.Windows.Shapes;
using System;
using System.IO;
using ClosedXML.Excel;
using Microsoft.Win32;
using Karambolo.PO;
using System.Diagnostics;

namespace PoFileGenerator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void OnConvertButtonClick(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls"
            };

            string selectedLanguage = GetSelectedLanguage();
            if (string.IsNullOrEmpty(selectedLanguage))
            {
                MessageBox.Show("언어를 선택하세요.", "오류", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    string poContent = ConvertExcelToPo(openFileDialog.FileName, selectedLanguage);

                    // PO 파일로 저장
                    string poFilePath = Path.Combine(Path.GetDirectoryName(openFileDialog.FileName), "UITextTable.po");
                    string moFilePath = Path.Combine(Path.GetDirectoryName(openFileDialog.FileName), "UITextTable.mo");

                    var utf8WithoutBom = new UTF8Encoding(false); // false는 BOM 포함 안 함
                    File.WriteAllText(poFilePath, poContent, utf8WithoutBom);

                    WriteMoFile(poFilePath, moFilePath);

                    MessageBox.Show("PO 파일이 생성되었습니다!", "완료", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"오류 발생: {ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private string GetSelectedLanguage()
        {
            if (radioEnglish.IsChecked == true) return "영어";
            if (radioJapanese.IsChecked == true) return "일본어";
            if (radioChinese.IsChecked == true) return "중국어";
            return string.Empty;
        }

        private string GetSelectedEnLanguage()
        {
            if (radioEnglish.IsChecked == true) return "English";
            if (radioJapanese.IsChecked == true) return "Japenese";
            if (radioChinese.IsChecked == true) return "Chinese (Simplified)";
            return string.Empty;
        }

        private string GetSelectedShortEnLanguage()
        {
            if (radioEnglish.IsChecked == true) return "en";
            if (radioJapanese.IsChecked == true) return "ja";
            if (radioChinese.IsChecked == true) return "zh-Hans";
            return string.Empty;
        }

        private string ConvertExcelToPo(string filePath, string language)
        {
            string shortLanguage = GetSelectedShortEnLanguage();
            string EnLanguage = GetSelectedEnLanguage();

            // Unreal Engine 포맷 헤더
            var poBuilder = new StringBuilder();
            poBuilder.AppendLine($"# UITextTable {EnLanguage} translation.");
            poBuilder.AppendLine("# Copyright Epic Games, Inc. All Rights Reserved.");
            poBuilder.AppendLine("# ");
            poBuilder.AppendLine("msgid \"\"");
            poBuilder.AppendLine("msgstr \"\"");
            poBuilder.AppendLine("\"Project-Id-Version: UITextTable\\n\"");
            poBuilder.AppendLine($"\"POT-Creation-Date: {DateTime.Now:yyyy-MM-dd HH:mm}\\n\"");
            poBuilder.AppendLine($"\"PO-Revision-Date: {DateTime.Now:yyyy-MM-dd HH:mm}\\n\"");
            poBuilder.AppendLine("\"Language-Team: \\n\"");
            poBuilder.AppendLine($"\"Language: {shortLanguage}\\n\"");
            poBuilder.AppendLine("\"MIME-Version: 1.0\\n\"");
            poBuilder.AppendLine("\"Content-Type: text/plain; charset=UTF-8\\n\"");
            poBuilder.AppendLine("\"Content-Transfer-Encoding: 8bit\\n\"");
            poBuilder.AppendLine("\"Plural-Forms: nplurals=2; plural=(n != 1);\\n\"");
            poBuilder.AppendLine();

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (var workbook = new XLWorkbook(stream))

                //using (var workbook = new XLWorkbook(filePath))
                {
                    var worksheet = workbook.Worksheet(1); // 첫 번째 시트를 사용
                    int keyColumnIndex = -1;
                    int sourceColumnIndex = -1;
                    int englishColumnIndex = -1;

                    // 첫 번째 행에서 "Key"와 "영어" 열을 찾음
                    var firstRow = worksheet.Row(1);
                    foreach (var cell in firstRow.CellsUsed())
                    {
                        if (cell.GetString().Equals("Key", StringComparison.OrdinalIgnoreCase))
                        {
                            keyColumnIndex = cell.Address.ColumnNumber;
                        }
                        else if(cell.GetString().Equals("SourceString", StringComparison.OrdinalIgnoreCase))
                        {
                            sourceColumnIndex = cell.Address.ColumnNumber;
                        }
                        else if (cell.GetString().Equals("영어", StringComparison.OrdinalIgnoreCase))
                        {
                            englishColumnIndex = cell.Address.ColumnNumber;
                        }
                    }

                    // 열 인덱스를 찾지 못한 경우 예외 처리
                    if (keyColumnIndex == -1 || englishColumnIndex == -1 || sourceColumnIndex == -1)
                    {
                        throw new InvalidOperationException("Key 또는 영어 열을 찾을 수 없습니다.");
                    }

                    // 두 번째 행부터 데이터 읽기
                    foreach (var row in worksheet.RowsUsed().Skip(1))
                    {
                        string tableName = "UITextTable";
                        string key = row.Cell(keyColumnIndex).GetString();
                        string sourceText = row.Cell(sourceColumnIndex).GetString();
                        string englishText = row.Cell(englishColumnIndex).GetString();

                        if(key == "연습장")
                        {
                            tableName = "CCTextTable";

                            poBuilder.AppendLine($"#. Key: {key}");
                            poBuilder.AppendLine($"#. SourceLocation:\t/Game/DB/Text/{tableName}.{tableName}");
                            poBuilder.AppendLine($"#: /Game/DB/Text/CCTextTable.CCTextTable");
                            poBuilder.AppendLine($"msgctxt \"{tableName},{key}\"");
                            poBuilder.AppendLine($"msgid \"{sourceText}\"");
                            poBuilder.AppendLine($"msgstr \"{englishText}\"");
                            poBuilder.AppendLine();
                        }
                        else
                        {
                            poBuilder.AppendLine($"#. Key: {key}");
                            poBuilder.AppendLine($"#. SourceLocation:\t/Game/DB/Text/{tableName}.{tableName}");
                            poBuilder.AppendLine($"#: /Game/DB/Text/UITextTable.UITextTable");
                            poBuilder.AppendLine($"msgctxt \"{tableName},{key}\"");
                            poBuilder.AppendLine($"msgid \"{sourceText}\"");
                            poBuilder.AppendLine($"msgstr \"{englishText}\"");
                            poBuilder.AppendLine();
                        }
                    }
                }
            }

            return poBuilder.ToString();
        }

        private void WriteMoFile(string poFilePath, string moFilePath)
        {
            // msgfmt.exe 실행 파일 경로 설정
            string msgfmtPath = "msgfmt"; // msgfmt.exe가 PATH에 설정된 경우
                                          // string msgfmtPath = @"C:\path\to\msgfmt.exe"; // msgfmt.exe의 전체 경로를 지정

            // msgfmt.exe 명령줄 인수 설정
            string arguments = $"-o \"{moFilePath}\" \"{poFilePath}\"";

            // ProcessStartInfo 설정
            var processStartInfo = new ProcessStartInfo
            {
                FileName = msgfmtPath,
                Arguments = arguments,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            // 외부 프로세스 실행
            using (var process = new Process { StartInfo = processStartInfo })
            {
                try
                {
                    process.Start();
                    string output = process.StandardOutput.ReadToEnd();
                    string error = process.StandardError.ReadToEnd();
                    process.WaitForExit();

                    if (process.ExitCode != 0)
                    {
                        throw new Exception($"msgfmt 실행 실패: {error}");
                    }

                    Console.WriteLine($"msgfmt 실행 성공: {output}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"오류 발생: {ex.Message}");
                    throw;
                }
            }
        }
    }
}