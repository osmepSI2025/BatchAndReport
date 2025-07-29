using System;
using System.Globalization;
using System.Text;

namespace BatchAndReport.DAO
{
    public class CommonDAO
    {
        public static string[] ToThaiDateString(DateTime date)
        {
            int thaiYear = date.Year + 543;
            string thaiMonth = date.ToString("MMMM", new CultureInfo("th-TH"));
            string day = date.Day.ToString();
            return new string[] { day, thaiMonth, thaiYear.ToString() };
        }

        public static string ToThaiDateStringCovert(DateTime date)
        {
            int thaiYear = date.Year + 543;
            string thaiMonth = date.ToString("MMMM", new CultureInfo("th-TH"));
            string day = date.Day.ToString();
            if(thaiYear<2500)
                {
                thaiYear = thaiYear+ 543;
            }
            return $"วันที่ {date.Day} เดือน {thaiMonth} พ.ศ. {thaiYear}";
        }

        // Function: Convert number to Thai text (Baht/Satang)
        public static string NumberToThaiText(decimal amount)
        {
            string[] numText = { "ศูนย์", "หนึ่ง", "สอง", "สาม", "สี่", "ห้า", "หก", "เจ็ด", "แปด", "เก้า" };
            string[] rankText = { "", "สิบ", "ร้อย", "พัน", "หมื่น", "แสน", "ล้าน" };

            string bahtText = "";
            string satangText = "";

            string[] parts = amount.ToString("0.00").Split('.');
            long baht = long.Parse(parts[0]);
            int satang = int.Parse(parts[1]);

            bahtText = ConvertIntegerToThaiText(baht, numText, rankText);
            if (baht == 0) bahtText = "ศูนย์";

            if (satang > 0)
            {
                satangText = ConvertIntegerToThaiText(satang, numText, rankText) + "สตางค์";
            }
            else
            {
                satangText = "ถ้วน";
            }

            return bahtText + "บาท" + satangText;
        }

        private static string ConvertIntegerToThaiText(long number, string[] numText, string[] rankText)
        {
            StringBuilder result = new StringBuilder();
            string numStr = number.ToString();
            int len = numStr.Length;

            for (int i = 0; i < len; i++)
            {
                int digit = int.Parse(numStr[i].ToString());
                int rank = len - i - 1;

                if (digit == 0) continue;

                if (rank == 0 && digit == 1 && len > 1)
                    result.Append("เอ็ด");
                else if (rank == 1 && digit == 2)
                    result.Append("ยี่");
                else if (rank == 1 && digit == 1)
                    result.Append("");
                else
                    result.Append(numText[digit]);

                result.Append(rankText[rank]);
            }
            return result.ToString();
        }
    }
}

