using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using ClosedXML.Excel;
using System.Configuration;
using MongoDB.Driver.Core.Configuration;

namespace Excel_to_db
{
    class Program
    {//"C:\\Users\\stajyer\\Desktop\\Loglar - Kopya.xlsx"
        static void Main(string[] args)
        {



            Int64 TransferID = 0;
            Boolean GondermeDurumu = false;
            string BaslamaZamani = "";
            string BitisZamani = "";
            int KameraID = 0;
            string Plaka = "";
            string Zaman = "";
            double Resim = 1;
            double ResimPlaka = 1;


            string FileName = "C:\\Users\\stajyer\\Desktop\\Loglar - Kopya.xlsx";
            var WorkBook = new XLWorkbook(FileName);
            var ws1 = WorkBook.Worksheet(1).RangeUsed().RowsUsed().Skip(1);

            #region Tablo Tanımı
            DataTable dtLoglar = new DataTable("Kamera Logları");
            dtLoglar.Columns.Add("TransferID", typeof(Int64));
            dtLoglar.Columns.Add("Gönderme Durumu", typeof(bool));
            dtLoglar.Columns.Add("Başlama Zamanı", typeof(string));
            dtLoglar.Columns.Add("Bitiş Zamanı", typeof(string));
            dtLoglar.Columns.Add("KameraID", typeof(int));
            dtLoglar.Columns.Add("Plaka", typeof(string));
            dtLoglar.Columns.Add("Zaman", typeof(string));
            dtLoglar.Columns.Add("Resim", typeof(double));
            dtLoglar.Columns.Add("Resim Plaka", typeof(double));
            #endregion

            using (var excelWorkbook = new XLWorkbook(FileName))
            {

                var Ws = excelWorkbook.Worksheet("Sayfa1");
                var SonSatir = Ws.RowsUsed().Count();

                int ilksatir = 1;
                int SiradakiSatir = 1;
                string BaglantiAdresi = "Server=LBT013\\SQLEXPRESS;Database=ekt;User Id=sa;Password=123;";
                SqlConnection Baglanti = new SqlConnection(BaglantiAdresi);
                Baglanti.ConnectionString = BaglantiAdresi;
                String query = "IF NOT EXISTS(SELECT * FROM ekt_table WHERE TransferID = @TransferID AND Plaka = @Plaka)" +
                         " INSERT INTO ekt_table(TransferID,GondermeDurumu,BaslamaZamani,BitisZamani,KameraID,Plaka,Zaman,Resim,ResimPlaka) " +
                         "VALUES(@TransferID,@GondermeDurumu,@BaslamaZamani,@BitisZamani,@KameraID,@Plaka,@Zaman,@Resim,@ResimPlaka) ";

                for (int i = 1; i <= SonSatir; i++)
                {
                    if (SiradakiSatir == 1)
                    {
                        TransferID = Convert.ToInt64(Ws.Cell(i, 2).GetString().Split(':')[1].Trim());

                    }
                    else if (SiradakiSatir == 2)
                    {
                        GondermeDurumu = Convert.ToBoolean(Ws.Cell(i, 2).GetString().Split(':')[1].Trim());

                    }
                    else if (SiradakiSatir == 3)
                    {
                        BaslamaZamani = Convert.ToString(Ws.Cell(i, 2).GetString().Split(' ')[3].Trim());

                    }
                    else if (SiradakiSatir == 4)
                    {
                        BitisZamani = Convert.ToString(Ws.Cell(i, 2).GetString().Split(' ')[3].Trim());

                    }
                    else
                    {
                        if (Ws.Cell(i, 2).GetString().StartsWith("-"))
                        {
                            SiradakiSatir = 1;
                            continue;

                        }
                        else
                        {
                            KameraID = Convert.ToInt32(Ws.Cell(i, 2).GetString().Split(':')[1].Trim());
                            Plaka = Convert.ToString(Ws.Cell(i, 3).GetString().Split(':')[1].Trim());
                            Zaman = Convert.ToString(Ws.Cell(i, 4).GetString().Split(' ')[1].Trim());
                            Resim = Convert.ToDouble(Ws.Cell(i, 5).GetString().Split(':')[1].Trim());
                            ResimPlaka = Convert.ToDouble(Ws.Cell(i, 6).GetString().Split(':')[1].Trim());

                            DataRow drLog = dtLoglar.NewRow();

                            drLog["TransferID"] = TransferID;
                            drLog["Gönderme Durumu"] = GondermeDurumu;
                            drLog["Başlama Zamanı"] = BaslamaZamani;
                            drLog["Bitiş Zamanı"] = BitisZamani;
                            drLog["KameraID"] = KameraID;
                            drLog["Plaka"] = Plaka;
                            drLog["Zaman"] = Zaman;
                            drLog["Resim"] = Resim;
                            drLog["Resim Plaka"] = ResimPlaka;

                            dtLoglar.Rows.Add(drLog);
                            dtLoglar.AcceptChanges();


                        }

                    }

                    SiradakiSatir++;
                    if (SiradakiSatir>=6)
                    {
                        using (SqlCommand command = new SqlCommand(query, Baglanti))
                        {
                            command.Parameters.AddWithValue("@TransferID", TransferID);
                            command.Parameters.AddWithValue("@GondermeDurumu", GondermeDurumu);
                            command.Parameters.AddWithValue("@BaslamaZamani", BaslamaZamani);
                            command.Parameters.AddWithValue("@BitisZamani", BitisZamani);
                            command.Parameters.AddWithValue("@KameraID", KameraID);
                            command.Parameters.AddWithValue("@Plaka", Plaka);
                            command.Parameters.AddWithValue("@Zaman", Zaman);
                            command.Parameters.AddWithValue("@Resim", Resim);
                            command.Parameters.AddWithValue("@ResimPlaka", ResimPlaka);
                            Baglanti.Open();
                            command.ExecuteNonQuery();
                            Baglanti.Close();
                        }
                    }
                }

                XLWorkbook WbYeni = new XLWorkbook();   
                WbYeni.Worksheets.Add(dtLoglar, "Kamera Logları");
                WbYeni.SaveAs("C:\\Users\\stajyer\\Desktop\\Çıktı2.xlsx");
                

                
                


            }

        }

    }
}








