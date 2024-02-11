using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelIsleme
{
        public partial class Form1 : Form
        {
            private string kaynakDizin = @"D:\KaynakExcel";
            private string hedefDizin = @"D:\HedefExcel";

            private System.Timers.Timer timer;
            private Queue<string> dosyaKuyrugu = new Queue<string>();
            private object kuyrukLock = new object();
            private bool islemDevamEdiyor = false;

            private DatabaseManager databaseManager = new DatabaseManager();

            public Form1()
            {
                InitializeComponent();

                timer = new System.Timers.Timer
                {
                    Interval = 5000,
                    AutoReset = true,
                    Enabled = false
                };

                timer.Elapsed += Timer_Elapsed;
            }

            private void Timer_Elapsed(object sender, ElapsedEventArgs e)
            {
                KontrolEtVeTasi();
            }

            private void KontrolEtVeTasi()
            {
                lock (kuyrukLock)
                {
                    if (islemDevamEdiyor)
                        return;

                    string[] excelDosyalari = Directory.GetFiles(kaynakDizin, "*.xlsx");
                    foreach (string excelDosyasi in excelDosyalari)
                    {
                        dosyaKuyrugu.Enqueue(excelDosyasi);
                    }

                    islemDevamEdiyor = true;
                }

                while (dosyaKuyrugu.Count > 0)
                {
                    string excelDosyasi;
                    lock (kuyrukLock)
                    {
                        excelDosyasi = dosyaKuyrugu.Dequeue();
                    }

                    string dosyaAdi = Path.GetFileName(excelDosyasi);
                    string hedefDosyaYolu = Path.Combine(hedefDizin, dosyaAdi);

                    File.Move(excelDosyasi, hedefDosyaYolu);

                    LogMesaji($"Dosya taşındı: {dosyaAdi}");

                    // Veritabanına ekle
                    ImportAndInsertToDatabase(hedefDosyaYolu);
                }

                lock (kuyrukLock)
                {
                    dosyaKuyrugu.Clear();
                }

                islemDevamEdiyor = false;
            }

            private void ImportAndInsertToDatabase(string excelDosyaYolu)
            {
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(excelDosyaYolu);
                Excel.Worksheet excelWorksheet = excelWorkbook.Sheets[1];

                try
                {
                    for (int row = 2; row <= excelWorksheet.UsedRange.Rows.Count; row++)
                    {
                        string mesken = excelWorksheet.Cells[row, 1].Value2?.ToString();
                        string ulke = excelWorksheet.Cells[row, 2].Value2?.ToString();
                        string saat = excelWorksheet.Cells[row, 4].Value2?.ToString();
                        string sehir = excelWorksheet.Cells[row, 5].Value2?.ToString();
                        string hayvan = excelWorksheet.Cells[row, 6].Value2?.ToString();
                        string oyun = excelWorksheet.Cells[row, 8].Value2?.ToString();

                        databaseManager.InsertData(mesken, ulke, saat, sehir, hayvan, oyun);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Hata: " + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    excelWorkbook.Close(false);
                    excelApp.Quit();
                }
            }

            private void LogMesaji(string mesaj)
            {
                if (InvokeRequired)
                {
                    Invoke(new Action(() => LogMesaji(mesaj)));
                }
                else
                {
                    listBoxLog.Items.Add($"{DateTime.Now:HH:mm:ss} - {mesaj}");
                    listBoxLog.SelectedIndex = listBoxLog.Items.Count - 1;
                }
            }

            private void btnBaslat_Click(object sender, EventArgs e)
            {
                timer.Start();
                LogMesaji("Excel dosyalarını bekliyorum.");
            }
            
            private void btnDurdur_Click(object sender, EventArgs e)
            {
                timer.Stop();
                LogMesaji("Kontrol ve taşıma durduruldu.");
            }
        }
    }




    using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelIsleme
{
        public class DatabaseManager
        {
            private string connectionString = "Data Source=DESKTOP-OPQQL1L\\SQLEXPRESS;Initial Catalog=DenemeExcelVT;Integrated Security=True";

            public void InsertData(string mesken, string ulke, string saat, string sehir, string hayvan, string oyun)
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    string insertQuery = "INSERT INTO DenemeExcelVT.dbo.KayitExcel (Mesken, Ulke, Saat, Sehir, Hayvan, Oyun) VALUES (@Mesken, @Ulke, @Saat, @Sehir, @Hayvan, @Oyun)";

                    using (SqlCommand command = new SqlCommand(insertQuery, connection))
                    {
                        command.Parameters.AddWithValue("@Mesken", mesken);
                        command.Parameters.AddWithValue("@Ulke", ulke);
                        command.Parameters.AddWithValue("@Saat", string.IsNullOrEmpty(saat) ? (object)DBNull.Value : saat);
                        command.Parameters.AddWithValue("@Sehir", sehir);
                        command.Parameters.AddWithValue("@Hayvan", hayvan);
                        command.Parameters.AddWithValue("@Oyun", oyun);

                        command.ExecuteNonQuery();
                    }
                }
            }
        }
    }
