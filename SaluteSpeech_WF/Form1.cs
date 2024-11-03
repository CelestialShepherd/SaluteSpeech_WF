using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using IronXL;

namespace SaluteSpeech_WF
{
    //Плавный вывод результатов
    public partial class Form1 : Form
    {
        //Источник токена отмены
        CancellationTokenSource _tokenSource;
//        Worksheet worksheet;
        WorkBook workBook;
        WorkSheet workSheet;
        Regex regexTextbox3 = new Regex(@"(\s|\S)+.xl(sx|sm|t)");

        public Form1()
        {
            InitializeComponent();
            //TODO: Убрать дубли
            if (regexTextbox3.IsMatch(textBox3.Text) && File.Exists(textBox3.Text))
            {
                //Получить экземпляр файла Excel
                workBook = WorkBook.Load(textBox3.Text);
                //Получить рабочий лист в файле Excel
                WorkSheet workSheet = workBook.WorkSheets.SingleOrDefault(wc => wc.Name.Equals("Words"));
                if (workSheet == null)
                {
                    textBox12.Text = "";
                    throw new Exception("Ошибка! В выбранном файле отсутствует лист под названием Words");
                }
                else
                {
                    // Получить количество строк
                    textBox12.Text = Convert.ToString(workSheet.RowCount - 1);
                }
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                openFileDialog1.Filter = "Книга Excel (*.xlsx)|*.xlsx|Книга Excel с поддержкой макросов (*.xlsm)|*.xlsm";
                openFileDialog1.Title = "Выбор загружаемого файла";
                openFileDialog1.AddExtension = true;
                openFileDialog1.ShowDialog();
                if (openFileDialog1.FileName.Equals("openFileDialog1"))
                {
                    openFileDialog1.FileName = "";
                }

                if (!openFileDialog1.CheckFileExists)
                    throw new Exception("Ошибка! Указан некорректный файл!");
                else if (!openFileDialog1.CheckPathExists)
                    throw new Exception("Ошибка! Указан некорректный путь!");
                else
                {
                    textBox3.Text = openFileDialog1.FileName;
                }
            }
            catch (Exception ex)
            {
                GenerateLog(ex.Message);
            }

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (regexTextbox3.IsMatch(textBox3.Text) && File.Exists(textBox3.Text))
            {
                //Получить экземпляр файла Excel
                workBook = WorkBook.Load(textBox3.Text);
                //Получить рабочий лист в файле Excel
                WorkSheet workSheet = workBook.WorkSheets.SingleOrDefault(wc => wc.Name.Equals("Words"));
                if (workSheet == null)
                {
                    textBox12.Text = "";
                    throw new Exception("Ошибка! В выбранном файле отсутствует лист под названием Words");
                }
                else 
                {
                    // Получить количество строк
                    textBox12.Text = Convert.ToString(workSheet.RowCount - 1);
                }
            }
            else
            {
                textBox12.Text = "";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                folderBrowserDialog1.ShowDialog();
                if (!openFileDialog1.CheckPathExists)
                    throw new Exception("Ошибка! Указан некорректный путь!");
                else
                {
                    textBox5.Text = folderBrowserDialog1.SelectedPath;
                }
            }
            catch (Exception ex)
            {
                GenerateLog(ex.Message);
            }
        }

        private void DisableUselessElements()
        {
            //Изменение статуса активности текстовых полей
            textBox1.Enabled = !textBox1.Enabled;
            textBox2.Enabled = !textBox2.Enabled;
            textBox3.Enabled = !textBox3.Enabled;
            textBox5.Enabled = !textBox5.Enabled;
            textBox6.Enabled = !textBox6.Enabled;
            textBox7.Enabled = !textBox7.Enabled;
            textBox8.Enabled = !textBox8.Enabled;
            textBox9.Enabled = !textBox9.Enabled;
            textBox10.Enabled = !textBox10.Enabled;
            textBox11.Enabled = !textBox11.Enabled;
            //Изменение статуса активности кнопок
            button1.Enabled = !button1.Enabled;
            button2.Enabled = !button2.Enabled;
            button3.Enabled = !button3.Enabled;
            button4.Enabled = !button4.Enabled;
            //Изменение статуса активности кнопки по остановке процесса генерации
            button5.Enabled = !button5.Enabled;
            //Изменение статусов чекбоксов
            checkBox1.Enabled = !checkBox1.Enabled;
            checkBox2.Enabled = !checkBox2.Enabled;
        }

        //TODO: Сделать запрос на токен с бОльшим интервалом
        //TODO: Добавить генерацию в формате MP3
        private async void button3_Click(object sender, EventArgs e)
        {
            DisableUselessElements();
            textBox4.Clear();
            textBox4.ClearUndo();
            //готовим токен отмены
            _tokenSource = new CancellationTokenSource();
            CancellationToken cancelToken = _tokenSource.Token;
            //TODO: Исправить логирование итогового текста предложения
            try
            {
                /*Назначение переменных*/
                string authBasic, authId, filePath, voiceQuestion, voiceAnswer;
                int maxRow, pauseTimeQuestion, pauseTimeAnswer, rowStart = 0, rowLimit = -1;
                DirectoryInfo dirInfo;
                List<string> sentences;

                /*Проверка сведений полей авторизации*/
                //Проверка ключа авторизации
                if (textBox1.Text.Equals(""))
                    throw new Exception("Ошибка! Не был указан ключ авторизации!");
                else
                    authBasic = textBox1.Text;
                //Проверка идентификатора авторизации
                if (textBox2.Text.Equals(""))
                    throw new Exception("Ошибка! Не был указан идентификатор!");
                else
                    authId = textBox2.Text;

                /*Проверка сведений о путях файлов*/
                //Проверка пути к файлу
                if (textBox3.Text.Equals(""))
                    throw new Exception("Ошибка! Не был указан путь к Excel-файлу!");
                else if (!new FileInfo(textBox3.Text).Exists)
                    throw new Exception("Ошибка! Указанного файла не существует!");
                else
                    filePath = textBox3.Text;
                //Проврека пути для генерации аудиофайлов
                if (textBox5.Text.Equals(""))
                    throw new Exception("Ошибка! Не был указан путь для генерации аудиофайлов!");
                else
                    dirInfo = GenerateAudioPath(textBox5.Text);

                /*Проверка настроек голосов*/
                //Проверка голоса вопроса
                if (textBox6.Text.Equals(""))
                    throw new Exception("Ошибка! Не был указан голос вопроса!");
                else if (IsCorrectVoice(textBox6.Text))
                    voiceQuestion = textBox6.Text;
                else
                    throw new Exception("Ошибка! Указан некорректный голос вопроса!");
                //Проверка голоса ответа
                if (textBox7.Text.Equals(""))
                    throw new Exception("Ошибка! Не был указан голос ответа!");
                else if (IsCorrectVoice(textBox7.Text))
                    voiceAnswer = textBox7.Text;
                else
                    throw new Exception("Ошибка! Указан некорректный голос ответа!");

                /*Проверка настроек пауз*/
                //Проверка значения длины паузы после вопроса
                if (textBox8.Text.Equals(""))
                    throw new Exception("Ошибка! Не была задана длина паузы после вопроса");
                else if (!Int32.TryParse(textBox8.Text, out pauseTimeQuestion))
                    throw new Exception("Ошибка! В поле длины паузы после вопроса указано некорректное значение");
                else if (pauseTimeQuestion <= 0)
                    throw new Exception("Ошибка! В поле длины паузы после вопроса указано отрицательное значение");
                //Проверка значения длины паузы после ответа
                if (textBox9.Text.Equals(""))
                    throw new Exception("Ошибка! Не была задана длина паузы после ответа");
                else if (!Int32.TryParse(textBox9.Text, out pauseTimeAnswer))
                    throw new Exception("Ошибка! В поле длины паузы после ответа указано некорректное значение");
                else if (pauseTimeAnswer <= 0)
                    throw new Exception("Ошибка! В поле длины паузы после ответа указано отрицательное значение");
   
                //Получить рабочий лист в файле Excel
//                worksheet = GetWorksheet(filePath);
                // Получить количество строк
//                maxRow = worksheet.Cells.MaxDataRow;
                
                //Получить экземпляр файла Excel
                workBook = WorkBook.Load(filePath);
                //Получить рабочий лист в файле Excel
                WorkSheet workSheet = workBook.WorkSheets.SingleOrDefault(wc => wc.Name.Equals("Words"));
                //Получить количество строк
                maxRow = workSheet.RowCount - 1;

                /*Проверка настроек значений строк для считывания*/
                //Проверка значения считываемых строк
                if (!checkBox1.Checked)
                    rowLimit = -1; // Ограничение количества строк отключено
                else if (textBox10.Text.Equals(""))
                    throw new Exception("Ошибка! Не было задано ограничение на количество считываемых строк");
                else if (!Int32.TryParse(textBox10.Text, out rowLimit))
                    throw new Exception("Ошибка! В поле количества считываемых строк указано некорректное значение");
                else if (rowLimit <= 0)
                    throw new Exception("Ошибка! В поле количества считываемых строк указано отрицательное или нулевое значение");
                else if (rowLimit > maxRow)
                    throw new Exception("Ошибка! В поле количества считываемых строк указано число, превышающее количество заполненных строк в файле");
                //Проверка значения начала считывания
                if (!checkBox2.Checked)
                    rowStart = 1;
                else if (textBox11.Text.Equals(""))
                    throw new Exception("Ошибка! Не был задан номер строки для начала считывания");
                else if (!Int32.TryParse(textBox11.Text, out rowStart))
                    throw new Exception("Ошибка! В поле номера строки для начала считывания указано некорректное значение");
                else if (rowStart <= 0)
                    throw new Exception("Ошибка! В поле номера строки для начала считывания указано отрицательное или нулевое значение");
                else if (rowStart > maxRow)
                    throw new Exception("Ошибка! В поле номера строки для начала считывания указано число, превышающее количество заполненных строк в файле");

                //Получить столбцы ячеек LID, L1_A
                //                List<object[,]> cellsList = IdentifyCells(worksheet, maxRow);
                List<Cell> cellsListLID = IdentifyColumns("LID", workSheet, maxRow);
                List<Cell> cellsListL1_A = IdentifyColumns("L1_A", workSheet, maxRow);

                //Генерация текстов запросов
                GenerateLog("Начало генерации текстов запросов");
                rowStart--;
                //Получить тексты запросов к генерации аудиофайлов
                //TODO: Переделать условие
                if (rowLimit == -1)
                    sentences = GetSentencesL1_A(cellsListL1_A, voiceQuestion, voiceAnswer, pauseTimeQuestion, pauseTimeAnswer, rowStart, maxRow);
                else if (rowStart + rowLimit > maxRow)
                    throw new Exception("Ошибка! Общее количество запрашиваемых строк превышает общее количество строк в файле");
                else
                    sentences = GetSentencesL1_A(cellsListL1_A, voiceQuestion, voiceAnswer, pauseTimeQuestion, pauseTimeAnswer, rowStart, rowStart + rowLimit);

                //Генерация аудиофайлов
                //GenerateLog("Начало генерации аудиофайлов");
                if (rowLimit == -1)
                    await GenerateAudioLID(cellsListLID, sentences, dirInfo, authBasic, authId, cancelToken, rowStart, maxRow);
                else
                    await GenerateAudioLID(cellsListLID, sentences, dirInfo, authBasic, authId, cancelToken, rowStart, rowLimit + rowStart);

                DisableUselessElements();
                GenerateLog("Генерация прошла успешно!");
            }
            catch (OperationCanceledException)
            {
                GenerateLog("Процесс генерации остановлен");
            }
            catch (Exception ex)
            {
                GenerateLog(ex.Message);
                DisableUselessElements();
            }

        }
        private List<Cell> IdentifyColumns(string header, WorkSheet workSheet, int maxRow)
        {
            string columnName = "";

            List<Cell> columnList;

            var cellHeader = workSheet.FirstOrDefault(cl => cl.Text.Equals(header));

            Regex regexEng = new Regex(@"[A-Za-z]");
            MatchCollection matches = regexEng.Matches(cellHeader.AddressString);

            foreach (var match in matches)
            {
                columnName += match;
            }

            columnList = workSheet.GetColumn(columnName).ToList();
            columnList.RemoveAt(0);

            return columnList;
        }

        //Получение аудозаписи с помощью типа данных application/ssml
        private async Task<byte[]> GetSoundHttpClientSSML(string token, string rawString)
        {
            //Игнорирование ненадежных SSL-сертификатов
            var handler = new HttpClientHandler()
            {
                ServerCertificateCustomValidationCallback = HttpClientHandler.DangerousAcceptAnyServerCertificateValidator
            };

            var client = new HttpClient(handler);
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

            var rawBytes = Encoding.UTF8.GetBytes(rawString);
            var rawContent = new ByteArrayContent(rawBytes);
            rawContent.Headers.Add("Content-Type", "application/ssml");

            //Отправка запроса
            HttpResponseMessage response = await client.PostAsync("https://smartspeech.sber.ru/rest/v1/text:synthesize?format=wav16", rawContent);

            //Обработка ответа
            var contents = await response.Content.ReadAsByteArrayAsync();

            return contents;
        }

        //Получение ответа от сервера авторизации токенов SaluteSpeech
        private async Task<string> GetTokenResponseHttpClient(string authBasic, string authId)
        {
            //Игнорирование ненадежных SSL-сертификатов
            var handler = new HttpClientHandler()
            {
                ServerCertificateCustomValidationCallback = HttpClientHandler.DangerousAcceptAnyServerCertificateValidator
            };

            var client = new HttpClient(handler);
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", authBasic);
            client.DefaultRequestHeaders.Add("RqUID", authId);

            var requestContent = new FormUrlEncodedContent(new[] {
                new KeyValuePair<string, string>("Content-Type","application/x-www-form-urlencoded"),
                new KeyValuePair<string, string>("scope", "SALUTE_SPEECH_CORP")
            });

            HttpResponseMessage response = await client.PostAsync("https://ngw.devices.sberbank.ru:9443/api/v2/oauth", requestContent);
            var contents = await response.Content.ReadAsStringAsync();

            return contents;
        }

        //Извлечение токена авторизации в ответе от сервера
        private string GetTokenHttpClient(string tokenResponse)
        {
            var sb = new StringBuilder(tokenResponse);
            string token = sb.ToString();
            sb.Remove(token.IndexOf("\",\"expires_at\":"), token.Length - token.IndexOf("\",\"expires_at\":")).Replace("{\"access_token\":\"", "");
            token = sb.ToString();

            return token;
        }

        private List<string> GetSentencesL1_A(List<Cell> cellsList, string voiceQuestion, string voiceAnswer, int pauseTimeQuestion, int pauseTimeAnswer, int rowStart, int rowEnd)
        {
            string temp;
            string[] sentences;
            string sentenceTemp;
            StringBuilder stringBuilder = new StringBuilder();
            string symbol = "(*)";
            List<string> resultSentences = new List<string>();

            for (int i = rowStart; i < rowEnd; i++)
            {
                //if (cellsList[i].Text.Trim().Equals("") || cellsList[i].Text.Trim().Equals(null))
                //    continue;
                sentences = cellsList[i].Text.Split(new[] { "?" }, StringSplitOptions.RemoveEmptyEntries);
                for (int j = 0; j < sentences.Length - 1; j++)
                {
                    temp = FilterPhrase(HighlightAccents(sentences[j].Trim() + "?"));
                    stringBuilder.Append(MakeQuestion(temp, voiceQuestion, pauseTimeQuestion));
                }

                if (sentences.Length != 0)
                    sentenceTemp = sentences[sentences.Length - 1];
                else
                    sentenceTemp = "";

                if (sentenceTemp.Contains(symbol))
                    sentenceTemp = sentenceTemp.Substring(0, sentenceTemp.IndexOf(symbol));
                stringBuilder.Append(MakeAnswer(FilterPhrase(HighlightAccents(sentenceTemp.Trim())), voiceAnswer, pauseTimeAnswer));
                GenerateLog(stringBuilder.ToString());
                resultSentences.Add(stringBuilder.ToString());
                stringBuilder.Clear();
            }

            return resultSentences;
        }

        private async Task GenerateAudioLID(List<Cell> cellsList, List<string> sentences, DirectoryInfo dirInfo, string authBasic, string authId, CancellationToken cancelToken, int rowStart, int rowEnd)
        {
            //Путь к папкам с генерируемыми аудиозаписями
            string word = "";
            StringBuilder subPathSb = new StringBuilder("");
            //Используемые переменные
            string token;

            textBox4.Text = "\r\n===================================================\r\n\r\n" + textBox4.Text;

            for (int i = rowStart; i < rowEnd; i++)
            {
                //Проверка на null строки со словом
                if (cellsList[i].Text.Trim().Equals("") || cellsList[i].Text.Equals(null))
                {
                    File.AppendAllText("D:\\SaluteSpeech_Generator_Test\\LogSkipped.txt", $"#{i}. Слово: {cellsList[i]} - Предложение: {sentences.First()}\r\n");
                    continue;
                }
                if (sentences.Equals(""))
                {
                    File.AppendAllText("D:\\SaluteSpeech_Generator_Test\\LogSkipped.txt", $"#{i}. Слово: {cellsList[i]} - Предложение:\r\n");
                    continue;
                }
                //Приведение значения пути к значению по умолчанию
                subPathSb.Clear();
                //Наименование файла
                word = cellsList[i].Text.ToString();

                //Получение токена авторизации
                textBox4.Text = (i + 1) + ". " + word + ". Getting Token - Started\r\n" + textBox4.Text;
                var tokenResponse = await GetTokenResponseHttpClient(authBasic, authId);
                token = GetTokenHttpClient(tokenResponse);

                //Отмена по запросу
                cancelToken.ThrowIfCancellationRequested();

                //Получение звукового файла
                textBox4.Text = (i + 1) + ". " + word + ". Getting Sound - Started\r\n" + textBox4.Text;
                var soundResponse = await GetSoundHttpClientSSML(token, sentences[i - rowStart]);
                System.IO.File.WriteAllBytes(dirInfo.ToString() + @"\L1_A\" + word + ".wav", soundResponse);

                //Отмена по запросу
                cancelToken.ThrowIfCancellationRequested();
            }

            textBox4.Text = "===================================================\r\n\r\n" + textBox4.Text;
        }

        //Убрать из фразы все символы кроме кириллицы и некотрых спецсимволов
        private string FilterPhrase(string phrase)
        {
            StringBuilder stringBuilder = new StringBuilder();
            Regex regex = new Regex(@"[А-Яа-яё,.!?'-]");
            string[] words = phrase.Split(' ');

            foreach (var word in words)
            {
                MatchCollection matches = regex.Matches(word);

                if (matches.Count > 0)
                {
                    foreach (Match match in matches)
                    {
                        if (match.Value != " ")
                            stringBuilder.Append(match);
                    }
                }
                stringBuilder.Append(" ");
            }
            //GenerateLog("Фраза отфильтрована");
            return stringBuilder.ToString().Trim();
        }

        private string PutIntonation(string phrase)
        {
            //TODO: Переписать под размер конкретную длину концов строк
            int[] ids = new int[4] { int.MaxValue, int.MaxValue, int.MaxValue, int.MaxValue };
            string phraseResult = "";

            string[] words;
            StringBuilder sb = new StringBuilder();

            do
            {
                if (phrase.Contains("."))
                    ids[0] = phrase.IndexOf(".");

                if (phrase.Contains("!?"))
                    ids[3] = phrase.IndexOf("!?");
                else if (phrase.Contains("?!"))
                    ids[3] = phrase.IndexOf("?!");
                else if (phrase.Contains("!"))
                    ids[1] = phrase.IndexOf("!");
                else if (phrase.Contains("?"))
                    ids[2] = phrase.IndexOf("?");

                if (ids.Min() == ids[2] || ids.Min() == ids[3])
                {
                    words = phrase.Substring(0, ids.Min() + 1).Split(' ');
                    if (words.Length > 0)
                    {
                        words[0] = "*" + words[0];
                        if (words.Length > 1)
                            words[1] = "*" + words[1];
                    }
                    foreach (string word in words)
                    {
                        sb.Append(word + " ");
                    }
                }
                else
                    sb.Append(phrase.Substring(0, ids.Min() + 1));

                phraseResult += sb.ToString();
                //Подготовка к следующему циклу
                if (ids.Min() != ids[3])
                    phrase = phrase.Substring(ids.Min() + 1).Trim();
                else
                    phrase = phrase.Substring(ids.Min() + 2).Trim();
                sb.Clear();
                ids = new int[4] { int.MaxValue, int.MaxValue, int.MaxValue, int.MaxValue };

            } while (phrase != "");

            return phraseResult.Trim();
        }

        private string MakeQuestion(string phrase, string voiceQuestion, int pauseTime)
        {
            phrase = PutIntonation(phrase);
            string pause = "<break time =\"" + pauseTime + "\" />";
            string question = "<voice name=\"" + voiceQuestion + "\">" + "<speak>" + phrase + "</speak>" + "</voice>" + pause;

            return question;
        }

        private string MakeAnswer(string phrase, string voiceAnswer, int pauseTime)
        {
            string pause = "<break time =\"" + pauseTime + "\" />";
            string answer = "<voice name=\"" + voiceAnswer + "\">" + phrase + "</voice>" + pause;

            return answer;
        }

        private bool IsCorrectVoice(string voice)
        {
            Regex regex = new Regex(@"\w+_(24000|8000)");

            return regex.IsMatch(voice);
        }

        private DirectoryInfo GenerateAudioPath(string path)
        {
            DirectoryInfo dirInfo = new DirectoryInfo(path);
            if (!dirInfo.Exists)
                dirInfo.Create();
            dirInfo.CreateSubdirectory(@"L1_A");

            return dirInfo;
        }

        private void GenerateLog(string message)
        {
            textBox4.Text = "===================================================\r\n\r\n" + message + "\r\n\r\n===================================================\r\n\r\n" + textBox4.Text;
            Application.DoEvents();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            textBox4.Clear();
            textBox4.ClearUndo();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                linkLabel1.LinkVisited = true;
                System.Diagnostics.Process.Start("https://developers.sber.ru/docs/ru/salutespeech/synthesis/voices");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка перехода по ссылке.\r\nТекст ошибки: {ex}");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            _tokenSource.Cancel();
            DisableUselessElements();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            MakeColLimitVisible();
        }

        private void MakeColLimitVisible()
        {
            label12.Enabled = !label12.Enabled;
            label12.Visible = !label12.Visible;
            textBox10.Enabled = !textBox10.Enabled;
            textBox10.Visible = !textBox10.Visible;
        }

        private string HighlightAccents(string sentence)
        {
            char[] letters;
            string wordEdited, sentenceEdited = "";
            string[] words = sentence.Split(' ');

            foreach (string word in words)
            {
                letters = word.ToCharArray();
                wordEdited = "";

                for (int i = 0; i < letters.Length; i++)
                {
                    if (letters[i] == 769)
                        wordEdited += "'";
                    else
                        wordEdited += letters[i];
                }

                sentenceEdited += wordEdited + " ";
            }

            //GenerateLog("Акценты подсвечены!");
            return sentenceEdited;
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            MakeColStartVisible();
        }

        private void MakeColStartVisible()
        {
            label13.Enabled = !label13.Enabled;
            label13.Visible = !label13.Visible;
            textBox11.Enabled = !textBox11.Enabled;
            textBox11.Visible = !textBox11.Visible;
        }
    }
}
